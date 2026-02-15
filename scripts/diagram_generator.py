#!/usr/bin/env python3
"""
図表生成（Gemini API）

【編集ガイド】
図解の内容やプロンプトを変更したい場合はこのファイルを編集してください。
specs リスト内の各タプル (ID, プロンプト) が1枚の図解に対応しています。
"""

import os
import base64
import time
from pathlib import Path
from typing import Dict

from models import HearingData
from config import Config

# Gemini API
try:
    from google import genai
    from google.genai import types
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False


def generate_diagrams(data: HearingData, output_dir: str) -> Dict[str, str]:
    """全ての図解を生成（Phase 5: exponential backoff付きリトライ）"""
    if not GEMINI_AVAILABLE:
        print("  ⚠️ Gemini APIが利用できません")
        return {}

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("  ⚠️ GEMINI_API_KEY未設定")
        return {}

    print(f"\n🎨 図解を生成中（{Config.GEMINI_MODEL}）...")

    client = genai.Client(api_key=api_key)
    diagram_dir = Path(output_dir) / "diagrams"
    diagram_dir.mkdir(exist_ok=True)

    c, s, l, e, f = data.company, data.labor_shortage, data.labor_saving, data.equipment, data.funding
    diagrams = {}

    specs = [
        ("01_企業概要", f"企業概要図\n会社名:{c.name}\n業種:{c.industry}\n従業員:{c.employee_count}名\n設立:{c.established_date}\n事業:{c.business_description}"),
        ("02_SWOT分析", f"SWOT分析図（4象限）\n強み:専門技術、経験豊富\n弱み:人手不足、業務効率低下\n機会:省力化設備導入\n脅威:人材確保競争激化"),
        ("03_人手不足", f"人手不足状況図\n必要人員:{s.desired_workers}名\n現在:{s.current_workers}名\n不足:{s.desired_workers-s.current_workers}名\n残業:{s.overtime_hours}時間/月"),
        ("04_課題フロー", f"課題の連鎖図（矢印で連鎖を示す）\n業種:{c.industry}\n対象業務:{s.shortage_tasks}\n\n人手不足（現{s.current_workers}名/必要{s.desired_workers}名）→業務過多（{s.shortage_tasks}に1日{l.current_hours}時間）→残業増加（月{s.overtime_hours}時間）→品質低下・離職リスク→さらなる人手不足\n\n根本原因：手作業中心の業務プロセスが非効率"),
        ("05_設備概要", f"導入設備概要\n名称:{e.name}\n金額:{e.total_price:,}円\n特徴:AI活用、自動化"),
        ("06_ビフォーアフター", f"ビフォーアフター比較図（横棒グラフ形式で工程別に表示）\n設備名:{e.name}\n\n" + "\n".join([f"{bp.name}: 導入前{bp.time_minutes}分→導入後{ap.time_minutes}分" for bp, ap in zip(data.before_processes, data.after_processes)]) + f"\n\n合計: 導入前{l.current_hours}時間→導入後{l.target_hours}時間\n削減:{l.reduction_hours:.1f}時間（{l.reduction_rate:.0f}%削減）"),
        ("07_効果算定", f"省力化効果の定量分析図\n設備名:{e.name}\n\n削減時間:{l.reduction_hours:.1f}時間/日\n月間削減:{l.reduction_hours*22:.0f}時間\n年間削減:{l.reduction_hours*Config.WORKING_DAYS_PER_YEAR:.0f}時間\n削減率:{l.reduction_rate:.0f}%\n人件費換算:年間約{int(l.reduction_hours*Config.WORKING_DAYS_PER_MONTH*12*Config.HOURLY_WAGE):,}円相当"),
        ("12_業務フロー", f"現状の業務フロー図（フローチャート形式・左から右に工程を並べる）\n会社名:{c.name}\n業種:{c.industry}\n対象業務:{s.shortage_tasks}\n\n" + "→".join([f"{p.name}({p.time_minutes}分)" for p in data.before_processes]) + f"\n\n合計所要時間: {sum(p.time_minutes for p in data.before_processes)}分/サイクル\n問題点: 手作業中心で1日{l.current_hours}時間を要する"),
        ("13_工程別比較", f"工程別の省力化効果比較チャート（横棒グラフ：各工程の導入前vs導入後の所要時間を色分けで並べる）\n設備名:{e.name}\n\n" + "\n".join([f"{bp.name}: 導入前{bp.time_minutes}分→導入後{ap.time_minutes}分（{bp.time_minutes-ap.time_minutes}分削減）" for bp, ap in zip(data.before_processes, data.after_processes)]) + f"\n\n全体削減率: {l.reduction_rate:.0f}%"),
        ("08_実施体制", f"実施体制図\n代表者:{c.representative}\n責任者:{f.implementation_manager}\n従業員:{c.employee_count}名"),
        ("09_スケジュール", f"実施スケジュール\n1ヶ月目:契約発注\n2ヶ月目:納品設置\n3ヶ月目:試運転\n4ヶ月目:本格稼働"),
        ("10_5年計画", f"5年計画グラフ\n付加価値額:年率+{(Config.GROWTH_RATE-1)*100:.0f}%成長\n給与支給総額:年率+{(Config.SALARY_GROWTH_RATE-1)*100:.1f}%成長\n投資回収:約2-3年"),
        ("11_実施工程", f"""補助事業のスケジュール表（ガントチャート形式）を作成してください。

【表の構成】
- 縦軸：フェーズとタスク名
- 横軸：補助事業実施期間（3月～翌3月の13ヶ月）＋ 事業計画1～5年目

【フェーズとタスク】
0.構想設計: 事業目的・目標設定(3-5月)、課題・改善方針検討(3-6月)、事業計画作成(4-7月)、社内プロジェクト体制決定(4-6月)、投資採算性・投資規模決定(5-8月)、予算・調達計画策定(6-8月)
1.機能設計: システム要件定義(6-8月)、システム構成策定(7-9月)、機能一覧定義(8-10月)
2.周辺機器の手配: 機械装置発注(8-9月)、部品・原材料調達(8-11月)
3.機能試作・システム組み立て: システム設計(9-11月)、システム発注・開発(10-12月)
4.評価: テスト・リリース(11-12月)、課題・改善方針検討(12-1月)
5.調整改善: システム再設計(1-2月)
6.稼働・実装: セキュリティ対策(2-3月,1-2年目)、保守・管理(3月以降,1-5年目継続)

【デザイン】
- 青系統の配色
- 活動期間は矢印(⇨)で表示
- プロフェッショナルなビジネス文書スタイル
- 会社名:{c.name}
- 設備名:{e.name}"""),
    ]

    for diagram_id, prompt in specs:
        print(f"    📊 {diagram_id}...", end=" ")
        output_path = diagram_dir / f"{diagram_id}.png"

        # Phase 5: exponential backoff 付きリトライ
        success = False
        for attempt in range(Config.GEMINI_RETRY_MAX):
            try:
                response = client.models.generate_content(
                    model=Config.GEMINI_MODEL,
                    contents=f"以下の内容を示すビジネス図解を生成してください。日本語で、青系統の配色で、プロフェッショナルなスタイルで。\n\n{prompt}",
                    config=types.GenerateContentConfig(response_modalities=["IMAGE", "TEXT"])
                )

                for part in response.candidates[0].content.parts:
                    if hasattr(part, 'inline_data') and part.inline_data:
                        image_data = part.inline_data.data
                        if isinstance(image_data, str):
                            image_data = base64.b64decode(image_data)
                        with open(output_path, 'wb') as f_out:
                            f_out.write(image_data)
                        if os.path.getsize(output_path) > 1000:
                            diagrams[diagram_id] = str(output_path)
                            print("✅")
                            success = True
                            break
                if success:
                    break
                if attempt < Config.GEMINI_RETRY_MAX - 1:
                    delay = Config.GEMINI_RETRY_BASE_DELAY * (2 ** attempt)
                    print(f"⏳ リトライ({attempt + 2}/{Config.GEMINI_RETRY_MAX})...", end=" ")
                    time.sleep(delay)
            except Exception as ex:
                if attempt < Config.GEMINI_RETRY_MAX - 1:
                    delay = Config.GEMINI_RETRY_BASE_DELAY * (2 ** attempt)
                    print(f"⏳ エラー、リトライ({attempt + 2}/{Config.GEMINI_RETRY_MAX})...", end=" ")
                    time.sleep(delay)
                else:
                    print(f"❌ ({ex})")

        if not success and diagram_id not in diagrams:
            print("❌")

        time.sleep(Config.GEMINI_INTER_REQUEST_DELAY)

    return diagrams
