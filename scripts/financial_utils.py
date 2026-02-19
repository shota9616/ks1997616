#!/usr/bin/env python3
"""
付加価値額の計算ロジックを一元化

【編集ガイド】
付加価値額の計算式やフォールバックロジックを変更したい場合はこのファイルを編集してください。
このファイルが「唯一の真実の源（Single Source of Truth）」です。

付加価値額 = 営業利益 + 人件費 + 減価償却費
"""

from models import HearingData
from config import Config


def calc_base_components(data: HearingData) -> dict:
    """付加価値額の基本構成要素を計算する（唯一の真実の源）

    各構成要素は、決算書PDFから取得した実値を優先し、
    取得できなかった場合のみフォールバック推計値を使用する。

    Args:
        data: ヒアリングデータ

    Returns:
        dict: {
            "labor_cost": int,       # 人件費（実値 or fallback: revenue×0.35）
            "depreciation": int,     # 減価償却費（実値 or fallback: 設備価格÷5年）
            "salary": int,           # 給与支給総額（実値 or fallback: revenue×0.30）
            "added_value": int,      # 付加価値額 = OP + LC + DEP
            "op_profit": int,        # 営業利益（そのまま）
            "revenue": int,          # 売上高（参照用）
        }
    """
    c = data.company
    e = data.equipment

    labor_cost = (
        c.labor_cost
        if c.labor_cost > 0
        else int(c.revenue_2024 * Config.LABOR_COST_RATIO)
    )
    depreciation = (
        c.depreciation
        if c.depreciation > 0
        else int(e.total_price / Config.DEPRECIATION_YEARS)
    )
    salary = (
        c.total_salary
        if c.total_salary > 0
        else int(c.revenue_2024 * Config.SALARY_RATIO)
    )
    added_value = c.operating_profit_2024 + labor_cost + depreciation

    return {
        "labor_cost": labor_cost,
        "depreciation": depreciation,
        "salary": salary,
        "added_value": added_value,
        "op_profit": c.operating_profit_2024,
        "revenue": c.revenue_2024,
    }


def calc_year_added_value(base: dict, year: int) -> int:
    """year年目の付加価値額を計算する

    各構成要素は異なる成長率で推移:
    - 営業利益: Config.GROWTH_RATE（デフォルト5%）
    - 人件費: Config.SALARY_GROWTH_RATE（デフォルト2.5%）
    - 減価償却費: 0%（定額、単一資産前提）

    Args:
        base: calc_base_components() の返り値
        year: 年数（0=基準年、1=1年目、...、5=5年目）

    Returns:
        int: year年目の付加価値額
    """
    av_op = int(base["op_profit"] * Config.GROWTH_RATE ** year)
    av_lc = int(base["labor_cost"] * Config.SALARY_GROWTH_RATE ** year)
    av_dep = int(base["depreciation"])  # 減価償却は定額
    return av_op + av_lc + av_dep


def calc_year_salary(base: dict, year: int) -> int:
    """year年目の給与支給総額を計算する

    Args:
        base: calc_base_components() の返り値
        year: 年数（0=基準年）

    Returns:
        int: year年目の給与支給総額
    """
    return int(base["salary"] * Config.SALARY_GROWTH_RATE ** year)
