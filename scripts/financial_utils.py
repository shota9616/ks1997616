#!/usr/bin/env python3
"""
付加価値額の計算ロジックを一元化

【編集ガイド】
付加価値額の計算式やフォールバックロジックを変更したい場合はこのファイルを編集してください。
このファイルが「唯一の真実の源（Single Source of Truth）」です。

付加価値額 = 営業利益 + 人件費 + 減価償却費
"""

from typing import List
from models import HearingData
from config import Config


def validate_financial_inputs(data: HearingData) -> List[str]:
    """入力段階で財務データの異常を検出する

    生成前に呼び出し、致命的な不整合を早期発見する。

    Args:
        data: ヒアリングデータ

    Returns:
        list[str]: 警告メッセージのリスト（空なら問題なし）
    """
    warnings = []
    c = data.company
    e = data.equipment

    # 販管費（人件費+減価償却費）が売上高を超えていないか
    if c.revenue_2024 > 0:
        estimated_sga = c.labor_cost + c.depreciation
        if estimated_sga > c.revenue_2024:
            warnings.append(
                f"⚠️ 人件費+減価償却費({estimated_sga:,}円)が売上高({c.revenue_2024:,}円)を超えています"
            )

    # 営業利益がマイナスの場合
    if c.operating_profit_2024 < 0:
        warnings.append(
            f"⚠️ 営業利益がマイナス({c.operating_profit_2024:,}円)です。計画値で営業利益が負になる可能性があります"
        )

    # 設備価格がゼロ
    if e.total_price <= 0:
        warnings.append("⚠️ 設備価格が未入力（0円）です")

    # 従業員数がゼロ
    if c.employee_count <= 0:
        warnings.append("⚠️ 従業員数が0です。1人当たり給与支給総額が計算できません")

    # 売上高がゼロ
    if c.revenue_2024 <= 0:
        warnings.append("⚠️ 売上高が未入力（0円）です")

    return warnings


def calc_base_components(data: HearingData) -> dict:
    """付加価値額の基本構成要素を計算する（唯一の真実の源）

    各構成要素は、決算書PDFから取得した実値を優先し、
    取得できなかった場合のみフォールバック推計値を使用する。

    減価償却費は「既存分 + 新規設備分」を合算する。
    既存分が決算書から取得できた場合もその上に新規設備の減価償却費を加算する。

    Args:
        data: ヒアリングデータ

    Returns:
        dict: {
            "labor_cost": int,             # 人件費（実値 or fallback: revenue×0.35）
            "depreciation": int,           # 減価償却費合計（既存 + 新規）
            "existing_depreciation": int,  # 既存減価償却費
            "new_depreciation": int,       # 新規設備減価償却費
            "salary": int,                 # 給与支給総額（実値 or fallback: revenue×0.30）
            "added_value": int,            # 付加価値額 = OP + LC + DEP
            "op_profit": int,              # 営業利益（そのまま）
            "revenue": int,               # 売上高（参照用）
        }
    """
    c = data.company
    e = data.equipment

    labor_cost = (
        c.labor_cost
        if c.labor_cost > 0
        else int(c.revenue_2024 * Config.LABOR_COST_RATIO)
    )

    # 既存減価償却費
    existing_depreciation = c.depreciation if c.depreciation > 0 else 0

    # 新規設備の減価償却費（定額法）
    new_depreciation = int(e.total_price / Config.DEPRECIATION_YEARS) if e.total_price > 0 else 0

    # 合計減価償却費 = 既存 + 新規
    # 既存がゼロの場合でも新規分は加算される
    depreciation = existing_depreciation + new_depreciation

    salary = (
        c.total_salary
        if c.total_salary > 0
        else int(c.revenue_2024 * Config.SALARY_RATIO)
    )
    added_value = c.operating_profit_2024 + labor_cost + depreciation

    return {
        "labor_cost": labor_cost,
        "depreciation": depreciation,
        "existing_depreciation": existing_depreciation,
        "new_depreciation": new_depreciation,
        "salary": salary,
        "added_value": added_value,
        "op_profit": c.operating_profit_2024,
        "revenue": c.revenue_2024,
    }


def calc_year_added_value(base: dict, year: int) -> int:
    """year年目の付加価値額を計算する

    各構成要素は異なる成長率で推移:
    - 営業利益: Config.GROWTH_RATE（デフォルト5%）
    - 人件費: Config.SALARY_GROWTH_RATE（デフォルト4%）
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


def calc_cagr(base_value: float, final_value: float, years: int) -> float:
    """CAGR（年平均成長率）を計算する

    Args:
        base_value: 基準年の値
        final_value: N年目の値
        years: 年数

    Returns:
        float: CAGR（例: 0.05 = 5%）。base_valueが0以下の場合は0を返す。
    """
    if base_value <= 0 or years <= 0:
        return 0.0
    return (final_value / base_value) ** (1 / years) - 1


def calc_per_capita_salary_cagr(
    base_salary: int, n_employees: int, years: int = 5
) -> float:
    """1人当たり給与支給総額のCAGRを計算する

    Args:
        base_salary: 基準年の給与支給総額
        n_employees: 対象従業員数
        years: 計画年数（デフォルト5年）

    Returns:
        float: 1人当たり給与支給総額のCAGR
    """
    if base_salary <= 0 or n_employees <= 0:
        return 0.0
    base_per_capita = base_salary / n_employees
    year5_salary = int(base_salary * Config.SALARY_GROWTH_RATE ** years)
    year5_per_capita = year5_salary / n_employees
    return calc_cagr(base_per_capita, year5_per_capita, years)


def check_requirements(data: HearingData) -> dict:
    """補助金3要件の充足判定を行う

    Args:
        data: ヒアリングデータ

    Returns:
        dict: {
            "added_value_cagr": float,        # 付加価値額CAGR
            "added_value_ok": bool,            # 要件充足か
            "salary_per_capita_cagr": float,   # 1人当たり給与CAGR
            "salary_per_capita_ok": bool,       # 要件充足か
            "warnings": list[str],             # 未達の場合の警告
        }
    """
    base = calc_base_components(data)
    warnings = []

    # 付加価値額CAGR
    year5_av = calc_year_added_value(base, 5)
    av_cagr = calc_cagr(base["added_value"], year5_av, 5)
    av_ok = av_cagr >= Config.REQUIREMENT_ADDED_VALUE_CAGR

    if not av_ok:
        warnings.append(
            f"⚠️ 付加価値額CAGR {av_cagr*100:.2f}% < 要件{Config.REQUIREMENT_ADDED_VALUE_CAGR*100:.1f}%"
        )

    # 1人当たり給与支給総額CAGR
    n_employees = data.company.employee_count
    salary_cagr = calc_per_capita_salary_cagr(base["salary"], n_employees, 5)
    salary_ok = salary_cagr >= Config.REQUIREMENT_SALARY_PER_CAPITA_CAGR

    if not salary_ok:
        warnings.append(
            f"⚠️ 1人当たり給与CAGR {salary_cagr*100:.2f}% < 要件{Config.REQUIREMENT_SALARY_PER_CAPITA_CAGR*100:.1f}%"
        )

    # 営業利益が5年間のどこかでマイナスにならないかチェック
    for yr in range(1, 6):
        projected_op = int(base["op_profit"] * Config.GROWTH_RATE ** yr)
        if projected_op < 0:
            warnings.append(f"⚠️ {yr}年目の営業利益が{projected_op:,}円（マイナス）")

    return {
        "added_value_cagr": av_cagr,
        "added_value_ok": av_ok,
        "salary_per_capita_cagr": salary_cagr,
        "salary_per_capita_ok": salary_ok,
        "warnings": warnings,
    }
