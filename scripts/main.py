#!/usr/bin/env python3
"""
省力化補助金（一般型）申請書類生成スクリプト v10.5 完全版

このファイルは後方互換性のための再エクスポートハブです。
実際のロジックは以下のモジュールに分割されています：

  models.py            - データクラス定義
  config.py            - 設定値（成長率、時給等）
  hearing_reader.py    - ヒアリングシート読み込み
  process_templates.py - 業種別工程テンプレート
  content_generator.py - 事業計画テキスト生成
  diagram_generator.py - 図表生成（Gemini API）
  document_writer.py   - 事業計画書Part1-2 Word生成
  plan3_writer.py      - 事業計画書Part3 Excel生成
  other_documents.py   - その他9種書類生成
  auto_fix.py          - 自動修正ループ・AI臭除去

【使用方法】
python scripts/main.py --hearing ヒアリングシート.xlsx --output ./output --template-dir ./templates
"""

import os
import sys
from pathlib import Path

# --- 再エクスポート: app.py / transcription_to_hearing.py との後方互換性維持 ---
from models import (
    CompanyInfo,
    LaborShortageInfo,
    LaborSavingInfo,
    EquipmentInfo,
    FundingInfo,
    WorkProcess,
    OfficerInfo,
    EmployeeInfo,
    ShareholderInfo,
    HearingData,
)
from config import Config
from hearing_reader import (
    read_hearing_sheet,
    validate_hearing_data,
    _split_name,
    _find_sheet_in_workbook,
)
from process_templates import generate_processes
from content_generator import ContentGenerator
from diagram_generator import generate_diagrams
from document_writer import generate_business_plan_1_2, add_schedule_table
from plan3_writer import generate_business_plan_3
from other_documents import generate_other_documents
from auto_fix import generate_with_auto_fix


# =============================================================================
# メイン処理（CLI）
# =============================================================================

def main():
    import argparse

    parser = argparse.ArgumentParser(description="省力化補助金申請書類生成 v10.5 完全版")
    parser.add_argument("--hearing", "-H", required=False, help="ヒアリングシートのパス")
    parser.add_argument("--from-transcription", help="議事録テキストからヒアリングシートを自動生成して使用")
    parser.add_argument("--output", "-o", default="./output", help="出力ディレクトリ")
    parser.add_argument("--template-dir", "-t", required=True, help="テンプレートディレクトリ")
    parser.add_argument("--no-diagrams", action="store_true", help="図解生成をスキップ")
    parser.add_argument("--auto-fix", action="store_true", help="85点以上になるまで自動修正ループ")
    parser.add_argument("--target-score", type=int, default=85, help="自動修正の目標スコア（デフォルト85）")
    parser.add_argument("--max-iterations", type=int, default=5, help="自動修正の最大リトライ回数（デフォルト5）")
    parser.add_argument("--no-deai", action="store_true", help="AI臭除去フェーズをスキップ")
    parser.add_argument("--target-ai-score", type=int, default=85, help="AI臭除去の目標スコア（デフォルト85）")
    parser.add_argument("--max-ai-rounds", type=int, default=3, help="AI臭除去の最大リライト回数（デフォルト3）")
    args = parser.parse_args()

    # --hearing か --from-transcription のいずれかが必須
    if not args.hearing and not args.from_transcription:
        parser.error("--hearing または --from-transcription のいずれかを指定してください")

    print("=" * 70)
    print("省力化補助金 申請書類生成スクリプト v10.5 完全版")
    print("- 事業者概要テーブル完全対応")
    print("- PREP法による採択レベル文章生成")
    print("- nano-banana-pro-preview 図解生成")
    print("=" * 70)

    template_dir = Path(args.template_dir)
    output_dir = Path(args.output)
    output_dir.mkdir(exist_ok=True, parents=True)

    # 1. データ読み込み
    hearing_path = args.hearing
    if args.from_transcription:
        # 議事録テキストから一時Excelを生成
        import tempfile
        from transcription_to_hearing import transcription_to_hearing as t2h
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        if not api_key:
            print("❌ --from-transcription 使用時は ANTHROPIC_API_KEY 環境変数が必要です")
            sys.exit(1)
        tmp_hearing = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_hearing.close()
        _, _, hearing_path = t2h(
            input_path=args.from_transcription,
            output_path=tmp_hearing.name,
            api_key=api_key,
        )
        print(f"  📄 生成されたヒアリングシート: {hearing_path}")

    data = read_hearing_sheet(hearing_path)

    # 2. 図解生成
    diagrams = {} if args.no_diagrams else generate_diagrams(data, str(output_dir))

    if args.auto_fix:
        # 自動修正ループ
        deai_enabled = not args.no_deai
        print(f"\n🔄 自動修正モード: 品質目標 {args.target_score}点 / 最大 {args.max_iterations}回")
        if deai_enabled:
            print(f"   AI臭除去: 目標 {args.target_ai_score}点 / 最大 {args.max_ai_rounds}回")
        result = generate_with_auto_fix(
            data=data,
            output_dir=str(output_dir),
            template_dir=template_dir,
            diagrams=diagrams,
            target_score=args.target_score,
            max_iterations=args.max_iterations,
            skip_diagrams=args.no_diagrams,
            deai=deai_enabled,
            target_ai_score=args.target_ai_score,
            max_ai_rounds=args.max_ai_rounds,
        )
        print("\n" + "=" * 70)
        print(f"品質スコア: {result['score']}/100 （{result['iterations']}回で完了）")
        for h in result["history"]:
            status = "PASS" if h["score"] >= args.target_score else "----"
            print(f"  [{status}] #{h['iteration']}: {h['score']}点")
        ai_r = result.get("ai_result", {})
        if ai_r and not ai_r.get("skipped"):
            print(f"AI臭スコア: {ai_r['ai_score']}/100 （{ai_r['ai_rounds']}回リライト）")
            for ah in ai_r.get("ai_history", []):
                print(f"  ラウンド{ah['round']}: {ah['score']}点 ({ah['grade']})")
        print(f"📁 出力先: {output_dir}")
        print("=" * 70)
    else:
        # 通常の1回生成
        from auto_fix import _run_generation
        _run_generation(data, str(output_dir), template_dir, diagrams)
        print("\n" + "=" * 70)
        print("✅ 全ての書類生成が完了しました！")
        print(f"📁 出力先: {output_dir}")
        print("=" * 70)


if __name__ == "__main__":
    main()
