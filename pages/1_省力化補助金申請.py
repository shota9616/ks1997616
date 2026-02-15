#!/usr/bin/env python3
"""
省力化補助金 申請書類生成ツール — ページ版
ヒアリングシート（Excel）+ 決算書PDF + 登記簿PDF から全11種の申請書類を自動生成する。
"""

import io
import os
import sys
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

# scripts/ を import path に追加
sys.path.insert(0, str(Path(__file__).parents[1] / "scripts"))

from main import (
    read_hearing_sheet,
    generate_diagrams,
    generate_business_plan_1_2,
    generate_business_plan_3,
    generate_other_documents,
    generate_with_auto_fix,
    OfficerInfo,
    Config,
    validate_hearing_data,
)
from transcription_to_hearing import (
    extract_from_transcription,
    validate_extracted_data,
    build_hearing_data,
    write_hearing_excel,
    ANTHROPIC_AVAILABLE,
)
from validate import check_files, check_diagrams, check_docx_text, check_plan3_values, calculate_score
from pdf_extractor import extract_financial_statements, extract_corporate_registry

# lib/ を import path に追加
sys.path.insert(0, str(Path(__file__).parents[1]))
from lib.auth import check_auth
from lib.styles import apply_styles, page_header, footer

# ---------------------------------------------------------------------------
# ページ設定
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="省力化補助金 申請書類生成",
    page_icon=":clipboard:",
    layout="centered",
)

apply_styles()

# --- Authentication ---
if not check_auth():
    st.stop()

page_header(
    ":clipboard: 省力化補助金 申請書類生成ツール",
    "ヒアリングシート + 決算書 + 登記簿 から申請に必要な全11種の書類を自動生成します。",
)

st.divider()

# ---------------------------------------------------------------------------
# セクション0: 議事録からヒアリングシート自動生成（オプション）
# ---------------------------------------------------------------------------
st.subheader("0. 議事録からヒアリングシート自動生成（オプション）")
st.caption("ミーティングの議事録テキストからヒアリングシートを自動生成できます。")

uploaded_transcript = st.file_uploader(
    "議事録テキストファイル（.txt）",
    type=["txt"],
    help="ミーティングの文字起こしテキスト。Claude APIで構造化データを抽出します。",
)

# 議事録 → ヒアリングシート生成結果を保持
generated_hearing_bytes = None
generated_hearing_data = None
use_generated_hearing = False

if uploaded_transcript is not None:
    # ANTHROPIC_API_KEY 取得
    transcript_api_key = ""
    env_anthropic = os.environ.get("ANTHROPIC_API_KEY", "")
    try:
        secrets_anthropic = st.secrets.get("ANTHROPIC_API_KEY", "")
    except Exception:
        secrets_anthropic = ""
    if secrets_anthropic == "your-anthropic-api-key-here":
        secrets_anthropic = ""

    if env_anthropic:
        st.info("環境変数の ANTHROPIC_API_KEY を使用します。")
        transcript_api_key = env_anthropic
    elif secrets_anthropic:
        st.info("Secrets に設定済みの ANTHROPIC_API_KEY を使用します。")
        transcript_api_key = secrets_anthropic
    else:
        transcript_api_key = st.text_input(
            "ANTHROPIC_API_KEY（議事録読み取り用）",
            type="password",
            help="Claude APIキーを入力してください。",
        )

    if not ANTHROPIC_AVAILABLE:
        st.error("anthropic パッケージがインストールされていません。pip install anthropic を実行してください。")
    elif transcript_api_key and st.button("ヒアリングシートを生成", type="secondary"):
        transcript_text = uploaded_transcript.getvalue().decode("utf-8")
        with st.status("議事録からデータを抽出中（Claude API×4回）...", expanded=True) as status:
            try:
                raw = extract_from_transcription(transcript_text, transcript_api_key)
                extraction_result = validate_extracted_data(raw)

                # 警告表示
                for w in extraction_result.warnings:
                    st.warning(f"抽出警告: {w}")

                hearing_data = build_hearing_data(extraction_result)
                st.write(f"企業名: **{hearing_data.company.name}**")
                st.write(f"業種: {hearing_data.company.industry}")
                st.write(f"設備: {hearing_data.equipment.name}")
                st.write(f"投資額: {hearing_data.equipment.total_price:,}円")

                # 一時ファイルにExcel書き出し
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                    write_hearing_excel(hearing_data, tmp.name)
                    with open(tmp.name, "rb") as f:
                        generated_hearing_bytes = f.read()
                    os.unlink(tmp.name)

                generated_hearing_data = hearing_data
                st.session_state["generated_hearing_bytes"] = generated_hearing_bytes
                st.session_state["generated_hearing_data"] = generated_hearing_data
                status.update(label="ヒアリングシート生成完了", state="complete")
            except Exception as e:
                status.update(label="生成エラー", state="error")
                st.error(f"議事録からの生成に失敗しました: {e}")

    # セッションステートから復元
    if "generated_hearing_bytes" in st.session_state:
        generated_hearing_bytes = st.session_state["generated_hearing_bytes"]
        generated_hearing_data = st.session_state["generated_hearing_data"]

    if generated_hearing_bytes:
        col_dl, col_use = st.columns(2)
        with col_dl:
            company_name = ""
            if generated_hearing_data:
                company_name = generated_hearing_data.company.name or "output"
            st.download_button(
                label="生成したヒアリングシートをダウンロード",
                data=generated_hearing_bytes,
                file_name=f"hearing_{company_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with col_use:
            use_generated_hearing = st.checkbox("このまま書類生成に進む", value=False)

st.divider()

# ---------------------------------------------------------------------------
# ファイルアップロード
# ---------------------------------------------------------------------------
st.subheader("1. ファイルアップロード")

uploaded = st.file_uploader(
    "ヒアリングシート（必須）",
    type=["xlsx"],
    help="10シート＋財務情報シートを含むExcelファイル",
)

col1, col2 = st.columns(2)
with col1:
    uploaded_financial = st.file_uploader(
        "決算書 PDF（任意）",
        type=["pdf"],
        help="損益計算書を含む決算書。アップロードすると参考書式の財務データが正確になります。",
    )
with col2:
    uploaded_registry = st.file_uploader(
        "履歴事項全部証明書 PDF（任意）",
        type=["pdf"],
        help="法人登記簿。アップロードすると役員情報・会社情報が正確になります。",
    )

# サンプルダウンロード
sample_path = Path(__file__).parents[1] / "examples" / "sample_hearing.xlsx"
if sample_path.exists():
    with open(sample_path, "rb") as f:
        st.download_button(
            label="サンプルヒアリングシートをダウンロード",
            data=f.read(),
            file_name="sample_hearing.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.divider()

# ---------------------------------------------------------------------------
# オプション設定
# ---------------------------------------------------------------------------
st.subheader("2. オプション設定")

use_diagrams = st.checkbox("図解も生成する（Gemini API）", value=False)

gemini_api_key = ""
if use_diagrams:
    env_key = os.environ.get("GEMINI_API_KEY", "")
    try:
        secrets_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        secrets_key = ""
    if secrets_key == "your-gemini-api-key-here":
        secrets_key = ""

    if env_key:
        st.info("環境変数の GEMINI_API_KEY を使用します。")
        gemini_api_key = env_key
    elif secrets_key:
        st.info("Secrets に設定済みの GEMINI_API_KEY を使用します。")
        gemini_api_key = secrets_key
    else:
        gemini_api_key = st.text_input(
            "GEMINI_API_KEY",
            type="password",
            help="Google AI Studio で取得した API キーを入力してください。",
        )

# PDF読み取り用にも GEMINI_API_KEY を使用
pdf_api_key = ""
if uploaded_financial or uploaded_registry:
    env_key = os.environ.get("GEMINI_API_KEY", "")
    try:
        secrets_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        secrets_key = ""
    if secrets_key == "your-gemini-api-key-here":
        secrets_key = ""

    if gemini_api_key:
        pdf_api_key = gemini_api_key
    elif env_key:
        pdf_api_key = env_key
    elif secrets_key:
        pdf_api_key = secrets_key
    else:
        pdf_api_key = st.text_input(
            "GEMINI_API_KEY（PDF読み取り用）",
            type="password",
            help="PDF読み取りに必要です。Google AI Studio で取得してください。",
        )

st.divider()

# ---------------------------------------------------------------------------
# 書類生成
# ---------------------------------------------------------------------------
st.subheader("3. 書類生成")

can_generate = (uploaded is not None) or use_generated_hearing
if st.button("書類を生成する", type="primary", disabled=(not can_generate)):
    if not can_generate:
        st.warning("ヒアリングシートをアップロードするか、議事録から生成してください。")
        st.stop()

    template_dir = Path(__file__).parents[1] / "templates"
    if not template_dir.exists():
        st.error("templates/ ディレクトリが見つかりません。")
        st.stop()

    # --- 一時ディレクトリで作業 ---
    with tempfile.TemporaryDirectory() as tmpdir:
        output_dir = os.path.join(tmpdir, "output")
        os.makedirs(output_dir, exist_ok=True)

        # 1. アップロードファイルを一時保存
        hearing_path = os.path.join(tmpdir, "hearing.xlsx")
        with open(hearing_path, "wb") as f:
            if use_generated_hearing and generated_hearing_bytes:
                f.write(generated_hearing_bytes)
            else:
                f.write(uploaded.getvalue())

        # 2. ヒアリングシート読み込み
        with st.status("ヒアリングシートを読み込み中...", expanded=True) as status:
            try:
                data = read_hearing_sheet(hearing_path)
                st.write(f"企業名: **{data.company.name}**")
                st.write(f"業種: {data.company.industry}")
                st.write(f"設備: {data.equipment.name}")
                st.write(f"投資額: {data.equipment.total_price:,}円")
                status.update(label="ヒアリングシート読み込み完了", state="complete")
            except Exception as e:
                status.update(label="読み込みエラー", state="error")
                st.error(f"ヒアリングシートの読み込みに失敗しました: {e}")
                st.stop()

        # 2.5. データバリデーション
        data_issues = validate_hearing_data(data)
        if data_issues:
            for issue in data_issues:
                st.warning(f"データ警告: {issue}")

        # 3. 決算書PDF読み取り（Claude API）
        if uploaded_financial and pdf_api_key:
            with st.status("決算書PDFを読み取り中（Claude API）...", expanded=True) as status:
                try:
                    fin_data = extract_financial_statements(
                        uploaded_financial.getvalue(), pdf_api_key
                    )
                    if fin_data:
                        # HearingData の財務情報を上書き
                        if fin_data.get("売上高", 0) > 0:
                            data.company.revenue_2024 = fin_data["売上高"]
                            data.company.revenue_2023 = int(fin_data["売上高"] / Config.GROWTH_RATE)
                            data.company.revenue_2022 = int(fin_data["売上高"] / Config.GROWTH_RATE / Config.GROWTH_RATE)
                        if fin_data.get("売上総利益", 0) > 0:
                            data.company.gross_profit_2024 = fin_data["売上総利益"]
                            data.company.gross_profit_2023 = int(fin_data["売上総利益"] / Config.GROWTH_RATE)
                            data.company.gross_profit_2022 = int(fin_data["売上総利益"] / Config.GROWTH_RATE / Config.GROWTH_RATE)
                        if "営業利益" in fin_data and fin_data["営業利益"] != 0:
                            data.company.operating_profit_2024 = fin_data["営業利益"]
                            data.company.operating_profit_2023 = int(fin_data["営業利益"] / Config.PROFIT_GROWTH_RATE)
                            data.company.operating_profit_2022 = int(fin_data["営業利益"] / Config.PROFIT_GROWTH_RATE / Config.PROFIT_GROWTH_RATE)
                        if fin_data.get("人件費", 0) > 0:
                            data.company.labor_cost = fin_data["人件費"]
                        if fin_data.get("減価償却費", 0) > 0:
                            data.company.depreciation = fin_data["減価償却費"]
                        if fin_data.get("給与支給総額", 0) > 0:
                            data.company.total_salary = fin_data["給与支給総額"]

                        st.write(f"売上高: **{fin_data.get('売上高', 0):,}円**")
                        st.write(f"営業利益: **{fin_data.get('営業利益', 0):,}円**")
                        st.write(f"人件費: **{fin_data.get('人件費', 0):,}円**")
                        st.write(f"減価償却費: **{fin_data.get('減価償却費', 0):,}円**")
                        st.write(f"給与支給総額: **{fin_data.get('給与支給総額', 0):,}円**")
                        status.update(label="決算書PDF読み取り完了", state="complete")
                    else:
                        status.update(label="決算書PDF: データ抽出できず", state="error")
                        st.warning("決算書PDFからデータを抽出できませんでした。ヒアリングシートの値を使用します。")
                except Exception as e:
                    status.update(label="決算書PDF読み取りエラー", state="error")
                    st.warning(f"決算書PDF読み取りエラー: {e}")
        elif uploaded_financial and not pdf_api_key:
            st.warning("GEMINI_API_KEY が未設定のため、決算書PDFの読み取りをスキップします。")

        # 4. 登記簿PDF読み取り（Claude API）
        if uploaded_registry and pdf_api_key:
            with st.status("履歴事項全部証明書を読み取り中（Claude API）...", expanded=True) as status:
                try:
                    reg_data = extract_corporate_registry(
                        uploaded_registry.getvalue(), pdf_api_key
                    )
                    if reg_data:
                        # HearingData の会社情報を上書き
                        if reg_data.get("会社名"):
                            data.company.name = reg_data["会社名"]
                        if reg_data.get("本店所在地"):
                            addr = reg_data["本店所在地"]
                            data.company.address = addr
                            # 都道府県を抽出
                            pref_found = False
                            for pref in ["東京都", "北海道", "大阪府", "京都府"]:
                                if addr.startswith(pref):
                                    data.company.prefecture = pref
                                    pref_found = True
                                    break
                            if not pref_found:
                                for i, ch in enumerate(addr):
                                    if ch == "県" and i <= 4:
                                        data.company.prefecture = addr[:i+1]
                                        break
                        if reg_data.get("設立年月日"):
                            data.company.established_date = reg_data["設立年月日"]
                        if reg_data.get("資本金", 0) > 0:
                            data.company.capital = reg_data["資本金"]
                        if reg_data.get("事業目的") and not data.company.business_description:
                            import re
                            purpose_text = reg_data["事業目的"]
                            first_line = purpose_text.split("\n")[0].strip()
                            first_line = re.sub(r"^[\d０-９]+[.．、)\s]+", "", first_line)
                            if len(first_line) > 50:
                                first_line = first_line[:50]
                            data.company.business_description = first_line

                        # 役員情報を上書き
                        officers = reg_data.get("役員", [])
                        if officers:
                            data.officers = [
                                OfficerInfo(
                                    name=o.get("氏名", ""),
                                    position=o.get("役職", "役員"),
                                    birth_date=o.get("就任日", ""),
                                )
                                for o in officers
                            ]
                            data.company.officer_count = len(data.officers)

                        st.write(f"会社名: **{reg_data.get('会社名', '')}**")
                        st.write(f"所在地: {reg_data.get('本店所在地', '')}")
                        st.write(f"設立: {reg_data.get('設立年月日', '')}")
                        st.write(f"資本金: {reg_data.get('資本金', 0):,}円")
                        st.write(f"役員数: {len(officers)}名")
                        status.update(label="履歴事項全部証明書読み取り完了", state="complete")
                    else:
                        status.update(label="登記簿PDF: データ抽出できず", state="error")
                        st.warning("登記簿PDFからデータを抽出できませんでした。")
                except Exception as e:
                    status.update(label="登記簿PDF読み取りエラー", state="error")
                    st.warning(f"登記簿PDF読み取りエラー: {e}")
        elif uploaded_registry and not pdf_api_key:
            st.warning("GEMINI_API_KEY が未設定のため、登記簿PDFの読み取りをスキップします。")

        # 5. 図解生成（オプション）
        diagrams = {}
        if use_diagrams and gemini_api_key:
            with st.status("図解を生成中（13枚）...", expanded=True) as status:
                os.environ["GEMINI_API_KEY"] = gemini_api_key
                try:
                    diagrams = generate_diagrams(data, output_dir)
                    st.write(f"生成完了: {len(diagrams)}/13 枚")
                    status.update(label=f"図解生成完了（{len(diagrams)}枚）", state="complete")
                except Exception as e:
                    status.update(label="図解生成エラー", state="error")
                    st.warning(f"図解生成中にエラーが発生しました: {e}")
        elif use_diagrams and not gemini_api_key:
            st.warning("GEMINI_API_KEY が未設定のため、図解生成をスキップします。")

        # 6. 書類生成（自動修正ループ）
        skip_diags = not bool(diagrams)
        target_score = 85
        max_iters = 5

        progress = st.progress(0, text="書類を生成中...")
        score_placeholder = st.empty()
        iteration_log = st.container()

        def on_progress(iteration, score, entry):
            pct = min(int((iteration / max_iters) * 80) + 10, 90)
            progress.progress(pct, text=f"イテレーション {iteration}/{max_iters} — スコア {score}/100")
            with iteration_log:
                if score >= target_score:
                    st.success(f"#{iteration}: {score}/100 — 目標達成！")
                else:
                    st.info(f"#{iteration}: {score}/100 — 自動修正して再生成...")

        # ANTHROPIC_API_KEY があればAI臭除去も実行
        has_anthropic_key = bool(os.environ.get("ANTHROPIC_API_KEY"))

        try:
            result = generate_with_auto_fix(
                data=data,
                output_dir=output_dir,
                template_dir=template_dir,
                diagrams=diagrams,
                target_score=target_score,
                max_iterations=max_iters,
                skip_diagrams=skip_diags,
                deai=has_anthropic_key,
                target_ai_score=85,
                max_ai_rounds=3,
                on_progress=on_progress,
            )
        except Exception as e:
            st.error(f"書類生成エラー: {e}")
            progress.progress(100, text="エラー")
            st.stop()

        progress.progress(90, text="検証中...")

        final_score = result["score"]
        iterations_used = result["iterations"]
        score_result = result["result"]
        ai_result = result.get("ai_result", {})

        # 7. スコア表示
        st.subheader("品質スコア")
        if final_score >= target_score:
            st.success(f"品質スコア: {final_score}/100 （{iterations_used}回で目標{target_score}点を達成）")
        else:
            st.warning(f"品質スコア: {final_score}/100 （{iterations_used}回実行、目標{target_score}点に未到達）")

        # AI臭スコア表示
        if ai_result and not ai_result.get("skipped"):
            ai_score = ai_result.get("ai_score")
            ai_rounds = ai_result.get("ai_rounds", 0)
            if ai_score is not None:
                if ai_score >= 85:
                    st.success(f"AI臭スコア: {ai_score}/100 （{ai_rounds}回リライト、自然な文章）")
                elif ai_score >= 70:
                    st.info(f"AI臭スコア: {ai_score}/100 （{ai_rounds}回リライト、概ね自然）")
                else:
                    st.warning(f"AI臭スコア: {ai_score}/100 （{ai_rounds}回リライト）")
                with st.expander("AI臭除去履歴"):
                    for ah in ai_result.get("ai_history", []):
                        label = "初回" if ah["round"] == 0 else f"ラウンド{ah['round']}"
                        st.write(f"{label}: **{ah['score']}点** ({ah['grade']})")
        elif not has_anthropic_key:
            st.info("ANTHROPIC_API_KEY を設定するとAI臭除去（自動リライト）が有効になります")

        # スコア内訳
        breakdown = score_result.get("breakdown", {})
        with st.expander("スコア内訳"):
            for cat, info in breakdown.items():
                label = {"files": "ファイル", "diagrams": "図解", "text_total": "総文字数",
                         "sections": "セクション文字数", "values": "数値要件"}.get(cat, cat)
                bar_pct = info["score"] / info["max"] if info["max"] > 0 else 0
                st.write(f"{label}: **{info['score']}/{info['max']}** {info.get('detail', '')}")
                st.progress(min(bar_pct, 1.0))

        # イテレーション履歴
        if len(result["history"]) > 1:
            with st.expander("自動修正履歴"):
                for h in result["history"]:
                    issues_str = ", ".join(h.get("issues", [])[:3])
                    if h["score"] >= target_score:
                        st.write(f"#{h['iteration']}: **{h['score']}点** — 合格")
                    else:
                        st.write(f"#{h['iteration']}: **{h['score']}点** — {issues_str}")

        # 8. 詳細検証結果
        st.subheader("検証結果")
        raw = score_result.get("raw", {})
        file_results = raw.get("files", [])
        diagram_results = raw.get("diagrams", {})
        text_results = raw.get("text", {})
        value_results = raw.get("values", {})

        # ファイル存在チェック
        file_ok = sum(1 for r in file_results if r["ok"])
        file_total = len(file_results)
        if file_ok == file_total:
            st.success(f"ファイル: {file_ok}/{file_total} 全ファイル生成済み")
        else:
            st.warning(f"ファイル: {file_ok}/{file_total}")
            for r in file_results:
                if not r["ok"]:
                    st.write(f"  - 未生成: {r['file']}")

        # 文字数チェック
        if "error" not in text_results:
            total_chars = text_results.get("total_chars", 0)
            if text_results.get("ok"):
                st.success(f"文字数: {total_chars:,}字（基準: {text_results['min_required']:,}字以上）")
            else:
                st.warning(f"文字数: {total_chars:,}字（基準: {text_results['min_required']:,}字以上）")
            sections = text_results.get("sections", {})
            if sections:
                with st.expander("セクション別文字数"):
                    for key in sorted(sections.keys()):
                        sec = sections[key]
                        mark = "OK" if sec["ok"] else "NG"
                        st.write(f"[{mark}] {key}: {sec['chars']:,}字 / {sec['min_required']:,}字")

        # 図解チェック
        if diagrams:
            d_found = diagram_results.get("found", 0)
            d_expected = diagram_results.get("expected", 13)
            if diagram_results.get("ok"):
                st.success(f"図解: {d_found}/{d_expected} 枚")
            else:
                st.warning(f"図解: {d_found}/{d_expected} 枚")

        # 数値チェック
        if isinstance(value_results, dict) and "error" not in value_results:
            has_issues = False
            for key, val in value_results.items():
                if isinstance(val, dict) and "ok" in val:
                    if val["ok"]:
                        st.success(f"{key}: {val.get('年率', '')} （基準: {val.get('基準', '')}）")
                    else:
                        st.error(f"{key}: {val.get('年率', '')} （基準: {val.get('基準', '')}）")
                        has_issues = True
            if not has_issues and value_results:
                st.success("基本要件（付加価値額+4%以上、給与支給総額+2%以上）をクリアしています")

        progress.progress(100, text="完了")

        # 9. ZIPダウンロード
        st.divider()
        st.subheader("ダウンロード")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, output_dir)
                    zf.write(file_path, arcname)
        zip_buffer.seek(0)

        company_name = data.company.name if data.company.name else "output"
        st.download_button(
            label="全書類をZIPでダウンロード",
            data=zip_buffer.getvalue(),
            file_name=f"省力化補助金_{company_name}_申請書類.zip",
            mime="application/zip",
            type="primary",
        )

elif not can_generate:
    st.info("ヒアリングシート（.xlsx）をアップロードするか、議事録から自動生成して「書類を生成する」を押してください。")

footer()
