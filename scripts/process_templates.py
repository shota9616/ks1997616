#!/usr/bin/env python3
"""
業種別Before/After工程テンプレート

【編集ガイド】
業種ごとの工程テンプレート（導入前・導入後）を変更したい場合はこのファイルを編集してください。
新しい業種を追加する場合は、generate_processes() 内に elif ブロックを追加してください。
"""

from typing import List, Tuple

from models import HearingData, WorkProcess


def generate_processes(data: HearingData) -> Tuple[List[WorkProcess], List[WorkProcess]]:
    """業種に応じた工程データを生成（Phase 3: 6業種対応）"""
    industry = data.company.industry

    if "建設" in industry or "建築" in industry:
        before = [
            WorkProcess("顧客打合せ", 60, "要件ヒアリング"),
            WorkProcess("図面作成", 120, "CAD設計"),
            WorkProcess("数量拾い出し", 90, "手作業計算"),
            WorkProcess("単価確認", 120, "見積依頼"),
            WorkProcess("見積書作成", 60, "書類作成"),
            WorkProcess("顧客説明", 30, "提案"),
        ]
        after = [
            WorkProcess("顧客打合せ", 60, "要件ヒアリング"),
            WorkProcess("図面作成", 120, "CAD設計"),
            WorkProcess("数量拾い出し", 10, "AI自動計算"),
            WorkProcess("単価確認", 15, "AIマッチング"),
            WorkProcess("見積書作成", 10, "自動生成"),
            WorkProcess("顧客説明", 30, "提案"),
        ]
    elif "製造" in industry:
        before = [
            WorkProcess("受注処理", 30, "注文確認・伝票起票"),
            WorkProcess("生産計画", 45, "手動スケジューリング"),
            WorkProcess("部材手配", 40, "在庫確認・発注"),
            WorkProcess("加工", 60, "手動作業"),
            WorkProcess("検品", 45, "目視確認"),
            WorkProcess("出荷準備", 30, "梱包・伝票作成"),
        ]
        after = [
            WorkProcess("受注処理", 10, "自動取り込み"),
            WorkProcess("生産計画", 10, "AI最適化"),
            WorkProcess("部材手配", 10, "自動発注"),
            WorkProcess("加工", 30, "自動化"),
            WorkProcess("検品", 15, "AI検査"),
            WorkProcess("出荷準備", 15, "自動梱包"),
        ]
    elif "IT" in industry or "情報" in industry:
        before = [
            WorkProcess("要件定義", 60, "顧客ヒアリング"),
            WorkProcess("設計", 90, "手動設計書作成"),
            WorkProcess("コーディング", 120, "手動開発"),
            WorkProcess("テスト", 60, "手動テスト"),
            WorkProcess("ドキュメント作成", 45, "手動作成"),
            WorkProcess("デプロイ", 30, "手動デプロイ"),
        ]
        after = [
            WorkProcess("要件定義", 60, "顧客ヒアリング"),
            WorkProcess("設計", 30, "AI支援設計"),
            WorkProcess("コーディング", 40, "AI支援開発"),
            WorkProcess("テスト", 15, "自動テスト"),
            WorkProcess("ドキュメント作成", 10, "自動生成"),
            WorkProcess("デプロイ", 10, "自動デプロイ"),
        ]
    elif "飲食" in industry:
        before = [
            WorkProcess("食材発注", 30, "在庫確認・手動発注"),
            WorkProcess("仕込み", 60, "手作業調理"),
            WorkProcess("注文受付", 20, "口頭・手書き"),
            WorkProcess("調理", 45, "手作業調理"),
            WorkProcess("会計", 15, "手動レジ"),
            WorkProcess("在庫管理", 30, "手動棚卸し"),
        ]
        after = [
            WorkProcess("食材発注", 5, "AI自動発注"),
            WorkProcess("仕込み", 40, "一部自動化"),
            WorkProcess("注文受付", 5, "タブレット注文"),
            WorkProcess("調理", 30, "調理支援機器"),
            WorkProcess("会計", 5, "自動精算"),
            WorkProcess("在庫管理", 5, "自動管理"),
        ]
    elif "サービス" in industry or "介護" in industry:
        before = [
            WorkProcess("予約管理", 30, "手動台帳管理"),
            WorkProcess("顧客対応", 45, "電話・来客対応"),
            WorkProcess("書類作成", 40, "手動作成"),
            WorkProcess("実作業", 60, "手作業"),
            WorkProcess("報告書作成", 30, "手書き"),
            WorkProcess("請求処理", 25, "手動計算"),
        ]
        after = [
            WorkProcess("予約管理", 5, "オンライン自動管理"),
            WorkProcess("顧客対応", 20, "AI自動応答併用"),
            WorkProcess("書類作成", 10, "自動生成"),
            WorkProcess("実作業", 40, "機器支援"),
            WorkProcess("報告書作成", 5, "自動生成"),
            WorkProcess("請求処理", 5, "自動計算"),
        ]
    elif "小売" in industry:
        before = [
            WorkProcess("発注業務", 30, "手動発注・在庫確認"),
            WorkProcess("検品", 25, "目視確認"),
            WorkProcess("陳列", 30, "手作業"),
            WorkProcess("接客", 40, "対面対応"),
            WorkProcess("会計", 20, "手動レジ"),
            WorkProcess("棚卸し", 45, "手動カウント"),
        ]
        after = [
            WorkProcess("発注業務", 5, "AI自動発注"),
            WorkProcess("検品", 10, "バーコード自動検品"),
            WorkProcess("陳列", 20, "最適配置提案"),
            WorkProcess("接客", 30, "セルフ+有人併用"),
            WorkProcess("会計", 5, "セルフレジ"),
            WorkProcess("棚卸し", 10, "自動在庫管理"),
        ]
    else:
        # デフォルト（汎用）
        before = [
            WorkProcess("検査", 30, "品質確認"),
            WorkProcess("準備", 20, "セットアップ"),
            WorkProcess("加工", 60, "手動作業"),
            WorkProcess("検品", 45, "目視確認"),
            WorkProcess("仕上げ", 30, "調整"),
            WorkProcess("梱包", 25, "出荷準備"),
        ]
        after = [
            WorkProcess("検査", 10, "自動検査"),
            WorkProcess("準備", 15, "自動セット"),
            WorkProcess("加工", 30, "自動化"),
            WorkProcess("検品", 15, "AI検査"),
            WorkProcess("仕上げ", 20, "効率化"),
            WorkProcess("梱包", 20, "効率化"),
        ]
    return before, after
