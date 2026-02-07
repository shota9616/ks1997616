#!/usr/bin/env python3
"""
PDF読み取りモジュール（Gemini API）

決算書・履歴事項全部証明書のPDFから構造化データを抽出する。
"""

import base64
import json
import os

from google import genai
from google.genai import types


def _call_gemini(pdf_bytes: bytes, prompt: str, api_key: str) -> dict:
    """Gemini API にPDFを送信してJSON形式のデータを取得する"""
    client = genai.Client(api_key=api_key)
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")

    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=[
            types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf"),
            prompt,
        ],
    )

    text = response.text

    # ```json ... ``` マーカーを除去
    cleaned = text.strip()
    if cleaned.startswith("```"):
        lines = cleaned.split("\n")
        lines = [l for l in lines if not l.strip().startswith("```")]
        cleaned = "\n".join(lines)

    # JSON部分を抽出
    start = cleaned.find("{")
    end = cleaned.rfind("}") + 1
    if start >= 0 and end > start:
        try:
            return json.loads(cleaned[start:end])
        except json.JSONDecodeError:
            return {}
    return {}


def extract_financial_statements(pdf_bytes: bytes, api_key: str) -> dict:
    """
    決算書PDFから財務データを抽出する。

    Returns:
        {
            "売上高": int,
            "売上総利益": int,
            "営業利益": int,
            "人件費": int,
            "減価償却費": int,
            "給与支給総額": int,
            "決算期": str,
        }
    """
    prompt = """この決算書（損益計算書・販売費及び一般管理費内訳等）から以下の数値を抽出してJSON形式で返してください。
数値は円単位の整数で返してください。千円単位で記載されている場合は1000倍してください。
見つからない場合は0としてください。

{
    "売上高": 0,
    "売上総利益": 0,
    "営業利益": 0,
    "人件費": 0,
    "減価償却費": 0,
    "給与支給総額": 0,
    "決算期": "2024年3月期"
}

【抽出ルール】
- 売上高: 「売上高」「売上金額」「営業収益」のいずれか
- 売上総利益: 「売上総利益」「粗利益」
- 営業利益: 「営業利益」（営業損失の場合はマイナスで）
- 人件費: 以下の合計値
  - 役員報酬
  - 給料手当（給料及び手当）
  - 賞与（賞与引当金繰入含む）
  - 法定福利費
  - 福利厚生費
  ※製造原価に含まれる人件費も加算すること
- 減価償却費: 販管費の減価償却費 + 製造原価の減価償却費
- 給与支給総額: 「給料手当」+「賞与」の合計（役員報酬は除く）
- 決算期: 対象となる事業年度

JSONのみ返してください。説明文は不要です。"""

    return _call_gemini(pdf_bytes, prompt, api_key)


def extract_corporate_registry(pdf_bytes: bytes, api_key: str) -> dict:
    """
    履歴事項全部証明書PDFから法人データを抽出する。

    Returns:
        {
            "会社名": str,
            "本店所在地": str,
            "設立年月日": str,
            "資本金": int,
            "事業目的": str,
            "役員": [{"氏名": str, "役職": str, "就任日": str}, ...],
        }
    """
    prompt = """この履歴事項全部証明書から以下の情報を抽出してJSON形式で返してください。
見つからない場合は空文字としてください。

{
    "会社名": "",
    "本店所在地": "",
    "設立年月日": "",
    "資本金": 0,
    "事業目的": "",
    "役員": [
        {"氏名": "", "役職": "", "就任日": ""}
    ]
}

【抽出ルール】
- 会社名: 「商号」欄の値
- 本店所在地: 「本店」欄の値（都道府県から番地まで全部）
- 設立年月日: 「会社成立の年月日」欄の値
- 資本金: 円単位の整数
- 事業目的: 「目的」欄の全文
- 役員: **現任**のもの全員（「退任」「辞任」と記載があるものは除く）
  - 「重任」は現任として含める
  - 役職は「代表取締役」「取締役」「監査役」等
  - 氏名は姓と名の間にスペースを入れる

JSONのみ返してください。説明文は不要です。"""

    return _call_gemini(pdf_bytes, prompt, api_key)
