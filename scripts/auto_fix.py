#!/usr/bin/env python3
"""自動修正ループ・AI臭除去"""

import os
import re
import json
from pathlib import Path

from docx import Document

from models import HearingData
from config import Config
from document_writer import generate_business_plan_1_2
from plan3_writer import generate_business_plan_3
from other_documents import generate_other_documents


def _run_generation(data: HearingData, output_dir: str, template_dir, diagrams: dict):
    """書類一式を生成する（1回分の実行）"""
    template_dir = Path(template_dir)
    Path(output_dir).mkdir(exist_ok=True, parents=True)

    t = template_dir / "事業計画書_その1その2_様式.docx"
    if t.exists():
        generate_business_plan_1_2(data, diagrams, str(output_dir), t)

    t = template_dir / "事業計画書_その3_様式.xlsx"
    if t.exists():
        generate_business_plan_3(data, str(output_dir), t)

    generate_other_documents(data, str(output_dir), template_dir)


def _apply_fixes(issues: list, data: HearingData) -> list:
    """スコアリング結果のissuesを解析し、パラメータを自動修正する。
    適用した修正のリストを返す。"""
    fixes_applied = []

    for issue in issues:
        action = issue.get("action", "")

        if action == "increase_growth_rate":
            old = Config.GROWTH_RATE
            Config.GROWTH_RATE = min(Config.GROWTH_RATE + 0.005, 1.10)  # 上限10%
            if Config.GROWTH_RATE != old:
                fixes_applied.append(f"GROWTH_RATE: {old} -> {Config.GROWTH_RATE}")

        elif action == "increase_salary_rate":
            old = Config.SALARY_GROWTH_RATE
            Config.SALARY_GROWTH_RATE = min(Config.SALARY_GROWTH_RATE + 0.005, 1.05)  # 上限5%
            if Config.SALARY_GROWTH_RATE != old:
                fixes_applied.append(f"SALARY_GROWTH_RATE: {old} -> {Config.SALARY_GROWTH_RATE}")

        elif action == "increase_text" or action == "increase_section_text":
            # テキスト不足はテンプレートで対応済みのため、再生成で解決を試みる
            if "テキスト再生成" not in [f.split(":")[0] for f in fixes_applied]:
                fixes_applied.append("テキスト再生成: リトライ")

    return fixes_applied


def _extract_docx_text(output_dir: str) -> str:
    """事業計画書docxから全テキストを抽出する"""
    docx_path = Path(output_dir) / "事業計画書_その1その2_完成版.docx"
    if not docx_path.exists():
        return ""
    doc = Document(str(docx_path))
    texts = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    t = para.text.strip()
                    if t:
                        texts.append(t)
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            texts.append(t)
    return "\n\n".join(texts)


def _write_text_to_docx(output_dir: str, rewritten_text: str):
    """リライト済みテキストを事業計画書docxのテーブルセルに書き戻す"""
    docx_path = Path(output_dir) / "事業計画書_その1その2_完成版.docx"
    if not docx_path.exists():
        return

    doc = Document(str(docx_path))

    # セクション番号→リライト済みテキストのマッピングを構築
    # リライト済みテキストをセクションヘッダーで分割
    section_map = {}
    current_key = None
    current_lines = []

    for line in rewritten_text.split("\n"):
        # セクションヘッダー検出（【...】パターン）
        header_match = re.match(r"^【(.+?)】", line.strip())
        if header_match:
            if current_key and current_lines:
                section_map[current_key] = "\n".join(current_lines).strip()
            current_key = header_match.group(1)
            current_lines = [line]
        elif current_key:
            current_lines.append(line)

    if current_key and current_lines:
        section_map[current_key] = "\n".join(current_lines).strip()

    if not section_map:
        # セクション分割できない場合、全体を最大のテーブルセルに書き込む
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if len(cell.text) > 500:
                        cell.text = rewritten_text
                        doc.save(str(docx_path))
                        return
        return

    # テーブルセルをスキャンし、対応するセクションのテキストを置換
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text.strip()
                for key, new_text in section_map.items():
                    if key in cell_text and len(cell_text) > 200:
                        cell.text = new_text
                        break

    doc.save(str(docx_path))


def _run_deai_phase(
    output_dir: str,
    industry: str,
    target_ai_score: int = 85,
    max_rounds: int = 3,
    on_progress=None,
) -> dict:
    """AI臭除去フェーズ: docxテキスト抽出→スコアリング→リライト→書き戻し

    Returns:
        dict: {ai_score, ai_rounds, ai_history, skipped}
    """
    # ai_smell_score をインポート
    skill_scripts = Path.home() / ".claude" / "skills" / "shoryokuka-review-deai" / "scripts"
    if not skill_scripts.exists():
        print("  AI臭除去スキルが未インストール。スキップします。")
        return {"ai_score": None, "ai_rounds": 0, "ai_history": [], "skipped": True}

    import importlib.util
    spec = importlib.util.spec_from_file_location("ai_smell_score", str(skill_scripts / "ai_smell_score.py"))
    ai_smell = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(ai_smell)

    # テキスト抽出
    text = _extract_docx_text(output_dir)
    if not text or len(text) < 100:
        print("  事業計画書テキストが短すぎます。AI臭除去をスキップ。")
        return {"ai_score": None, "ai_rounds": 0, "ai_history": [], "skipped": True}

    # 初回スコアリング
    result = ai_smell.calculate_score(text)
    ai_score = result["total_score"]
    ai_history = [{"round": 0, "score": ai_score, "grade": result["grade"]}]
    print(f"\n  AI臭スコア（初回）: {ai_score}/100 ({result['grade']})")

    if on_progress:
        on_progress("ai_smell_initial", ai_score, result)

    if ai_score >= target_ai_score:
        print(f"  AI臭スコア {ai_score} >= {target_ai_score}。リライト不要。")
        return {"ai_score": ai_score, "ai_rounds": 0, "ai_history": ai_history, "skipped": False}

    # auto_rewrite のコア関数をインポート
    spec2 = importlib.util.spec_from_file_location("auto_rewrite", str(skill_scripts / "auto_rewrite.py"))
    auto_rw = importlib.util.module_from_spec(spec2)
    spec2.loader.exec_module(auto_rw)

    # ANTHROPIC_API_KEY チェック
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("  ANTHROPIC_API_KEY 未設定。AI臭除去のリライトをスキップ。")
        return {"ai_score": ai_score, "ai_rounds": 0, "ai_history": ai_history, "skipped": True}

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
    except ImportError:
        print("  anthropic パッケージ未インストール。AI臭除去のリライトをスキップ。")
        return {"ai_score": ai_score, "ai_rounds": 0, "ai_history": ai_history, "skipped": True}

    # 参照ファイル読み込み
    skill_root = skill_scripts.parent
    system_prompt = ""
    rewrite_prompt_path = skill_root / "prompts" / "rewrite_system.txt"
    if rewrite_prompt_path.exists():
        system_prompt = rewrite_prompt_path.read_text(encoding="utf-8")

    patterns_path = skill_root / "reference" / "ai_smell_patterns.md"
    patterns_text = patterns_path.read_text(encoding="utf-8") if patterns_path.exists() else ""

    good_examples_path = skill_root / "reference" / "good_examples.md"
    good_examples_text = good_examples_path.read_text(encoding="utf-8") if good_examples_path.exists() else ""

    vocab_path = skill_root / "reference" / "industry_vocab.json"
    vocab_data = json.loads(vocab_path.read_text(encoding="utf-8")) if vocab_path.exists() else {}

    full_system = f"{system_prompt}\n\n---\n\n## 参照: AI臭パターン辞典\n\n{patterns_text}\n\n---\n\n## 参照: 採択済み申請書の文体サンプル\n\n{good_examples_text}"

    # リライトループ
    current_text = text
    for round_num in range(1, max_rounds + 1):
        print(f"\n  AI臭除去 ラウンド {round_num}/{max_rounds}...")

        weak_areas = auto_rw.identify_weak_areas(result)
        instruction = auto_rw.build_rewrite_instruction(
            weak_areas, industry, round_num, vocab_data, None,
        )

        try:
            rewritten = auto_rw.rewrite_with_claude(
                client, current_text, full_system, instruction,
                auto_rw.DEFAULT_MODEL,
            )
        except Exception as e:
            print(f"  リライトAPI失敗: {e}")
            break

        result = ai_smell.calculate_score(rewritten)
        ai_score = result["total_score"]
        ai_history.append({"round": round_num, "score": ai_score, "grade": result["grade"]})
        print(f"  AI臭スコア（ラウンド{round_num}）: {ai_score}/100 ({result['grade']})")

        if on_progress:
            on_progress(f"ai_smell_round_{round_num}", ai_score, result)

        current_text = rewritten

        if ai_score >= target_ai_score:
            print(f"  AI臭スコア目標達成！ {ai_score} >= {target_ai_score}")
            break

        # スコアが改善しなかったら終了
        if round_num >= 2 and ai_history[-1]["score"] <= ai_history[-2]["score"]:
            print(f"  スコア改善なし。ループ終了。")
            break

    # リライト結果をdocxに書き戻し
    if len(ai_history) > 1:
        print(f"  リライト結果をdocxに書き戻し中...")
        _write_text_to_docx(output_dir, current_text)
        # リライト済みテキストも保存
        rewrite_path = Path(output_dir) / "事業計画書_リライト済み.txt"
        rewrite_path.write_text(current_text, encoding="utf-8")
        print(f"  保存: {rewrite_path}")

    return {"ai_score": ai_score, "ai_rounds": len(ai_history) - 1, "ai_history": ai_history, "skipped": False}


def generate_with_auto_fix(
    data: HearingData,
    output_dir: str,
    template_dir,
    diagrams: dict = None,
    target_score: int = 85,
    max_iterations: int = 5,
    skip_diagrams: bool = False,
    deai: bool = True,
    target_ai_score: int = 85,
    max_ai_rounds: int = 3,
    on_progress=None,
) -> dict:
    """スコアが目標に達するまで生成→検証→修正を繰り返し、
    品質スコア達成後にAI臭除去フェーズを実行する。
    """
    from validate import calculate_score

    if diagrams is None:
        diagrams = {}

    history = []

    # === Phase 1: 書類品質ループ ===
    for iteration in range(1, max_iterations + 1):
        # --- 生成 ---
        _run_generation(data, output_dir, template_dir, diagrams)

        # --- スコアリング ---
        result = calculate_score(Path(output_dir), skip_diagrams=skip_diagrams)
        current_score = result["score"]
        history.append({
            "iteration": iteration,
            "score": current_score,
            "breakdown": result["breakdown"],
            "issues": [i["detail"] for i in result["issues"]],
        })

        if on_progress:
            on_progress(iteration, current_score, history[-1])

        print(f"\n{'='*50}")
        print(f"  イテレーション {iteration}/{max_iterations}: 品質スコア {current_score}/100")
        for cat, info in result["breakdown"].items():
            print(f"    {cat}: {info['score']}/{info['max']}")

        # --- 目標達成チェック ---
        if current_score >= target_score:
            print(f"  品質スコア {target_score} を達成！")
            break

        # --- 最終イテレーションなら終了 ---
        if iteration >= max_iterations:
            print(f"  最大イテレーション {max_iterations} に到達。最終スコア: {current_score}")
            break

        # --- 自動修正 ---
        fixes = _apply_fixes(result["issues"], data)
        if not fixes:
            print(f"  追加の自動修正なし。最終スコア: {current_score}")
            break

        print(f"  自動修正を適用:")
        for fix in fixes:
            print(f"    - {fix}")

        # 出力ディレクトリをクリーンアップして再生成
        out_path = Path(output_dir)
        for f in out_path.glob("*_完成版.*"):
            f.unlink()

    # === Phase 2: AI臭除去 ===
    ai_result = {"ai_score": None, "ai_rounds": 0, "ai_history": [], "skipped": True}
    if deai:
        industry = data.company.industry or "サービス"
        print(f"\n{'='*50}")
        print(f"  Phase 2: AI臭除去（業種: {industry}）")
        ai_result = _run_deai_phase(
            output_dir=output_dir,
            industry=industry,
            target_ai_score=target_ai_score,
            max_rounds=max_ai_rounds,
            on_progress=on_progress,
        )

    final = calculate_score(Path(output_dir), skip_diagrams=skip_diagrams)
    return {
        "score": final["score"],
        "iterations": len(history),
        "history": history,
        "result": final,
        "ai_result": ai_result,
    }
