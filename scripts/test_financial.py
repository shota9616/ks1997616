#!/usr/bin/env python3
"""付加価値額の整合性テスト

financial_utils.py が唯一の真実の源として正しく動作することを検証する。
"""

import sys
from pathlib import Path

# scripts/ ディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent))

import pytest
from models import HearingData, CompanyInfo, EquipmentInfo
from config import Config
from financial_utils import (
    calc_base_components,
    calc_year_added_value,
    calc_year_salary,
    calc_cagr,
    calc_per_capita_salary_cagr,
    check_requirements,
    validate_financial_inputs,
)


class TestCalcBaseComponents:
    """calc_base_components のユニットテスト"""

    def _make_data(
        self,
        labor=6_713_298,
        dep=2_822_935,
        op=2_275_980,
        revenue=64_199_095,
        equip_price=14_114_675,
        salary=0,
        employee_count=2,
    ):
        """テスト用のHearingDataを作成"""
        data = HearingData()
        data.company.operating_profit_2024 = op
        data.company.labor_cost = labor
        data.company.depreciation = dep
        data.company.revenue_2024 = revenue
        data.company.total_salary = salary
        data.company.employee_count = employee_count
        data.equipment.total_price = equip_price
        return data

    def test_basic_calculation(self):
        """正常系: 実値がすべて設定されている場合"""
        data = self._make_data()
        base = calc_base_components(data)
        # 減価償却費 = 既存(2,822,935) + 新規(14,114,675/5=2,822,935)
        expected_dep = 2_822_935 + int(14_114_675 / 5)
        expected = 2_275_980 + 6_713_298 + expected_dep
        assert base["added_value"] == expected
        assert base["labor_cost"] == 6_713_298
        assert base["depreciation"] == expected_dep
        assert base["op_profit"] == 2_275_980

    def test_depreciation_includes_new_equipment(self):
        """★リグレッション: 既存減価償却費がある場合も新規設備分が加算される"""
        data = self._make_data(dep=2_499_244, equip_price=11_250_000)
        base = calc_base_components(data)
        expected_existing = 2_499_244
        expected_new = int(11_250_000 / 5)  # 2,250,000
        assert base["existing_depreciation"] == expected_existing
        assert base["new_depreciation"] == expected_new
        assert base["depreciation"] == expected_existing + expected_new
        # 減価償却費は既存だけでなく、必ず既存+新規になること
        assert base["depreciation"] > expected_existing

    def test_depreciation_fallback_when_zero(self):
        """depreciation=0の場合、新規設備分のみになる"""
        data = self._make_data(dep=0, equip_price=10_000_000)
        base = calc_base_components(data)
        assert base["existing_depreciation"] == 0
        assert base["new_depreciation"] == int(10_000_000 / 5)
        assert base["depreciation"] == int(10_000_000 / 5)

    def test_fallback_labor_cost(self):
        """labor_cost=0の場合、revenue*0.35がfallbackされる"""
        data = self._make_data(labor=0)
        base = calc_base_components(data)
        expected_labor = int(64_199_095 * 0.35)
        assert base["labor_cost"] == expected_labor
        assert base["labor_cost"] != 0

    def test_fallback_salary(self):
        """total_salary=0の場合、revenue*0.30がfallbackされる"""
        data = self._make_data(salary=0)
        base = calc_base_components(data)
        expected_salary = int(64_199_095 * 0.30)
        assert base["salary"] == expected_salary

    def test_added_value_formula(self):
        """★リグレッション: 付加価値額 = 営業利益 + 人件費 + 減価償却費"""
        data = self._make_data()
        base = calc_base_components(data)
        assert base["added_value"] == base["op_profit"] + base["labor_cost"] + base["depreciation"]

    def test_consistency_across_calls(self):
        """同じデータで複数回呼んでも同じ結果になる（冪等性）"""
        data = self._make_data()
        base1 = calc_base_components(data)
        base2 = calc_base_components(data)
        assert base1 == base2

    def test_base_has_depreciation_breakdown(self):
        """返り値にexisting_depreciation, new_depreciationが含まれる"""
        data = self._make_data()
        base = calc_base_components(data)
        assert "existing_depreciation" in base
        assert "new_depreciation" in base
        assert base["depreciation"] == base["existing_depreciation"] + base["new_depreciation"]


class TestCalcYearAddedValue:
    """calc_year_added_value のユニットテスト"""

    def _make_base(self):
        data = HearingData()
        data.company.operating_profit_2024 = 2_275_980
        data.company.labor_cost = 6_713_298
        data.company.depreciation = 2_822_935
        data.company.revenue_2024 = 64_199_095
        data.company.employee_count = 2
        data.equipment.total_price = 14_114_675
        return calc_base_components(data)

    def test_year0_equals_base(self):
        """year=0 は基準年度の付加価値額と一致"""
        base = self._make_base()
        assert calc_year_added_value(base, 0) == base["added_value"]

    def test_year5_greater_than_base(self):
        """5年目の付加価値額 > 基準年度"""
        Config.reset_rates()
        base = self._make_base()
        year5 = calc_year_added_value(base, 5)
        assert year5 > base["added_value"]

    def test_monotonic_increase(self):
        """付加価値額は単調増加（成長率>1の場合）"""
        Config.reset_rates()
        base = self._make_base()
        prev = calc_year_added_value(base, 0)
        for yr in range(1, 6):
            current = calc_year_added_value(base, yr)
            assert current > prev, f"year {yr}: {current} <= {prev}"
            prev = current

    def test_growth_rate_positive(self):
        """5年後の付加価値額年率がプラス成長"""
        Config.reset_rates()
        base = self._make_base()
        year5 = calc_year_added_value(base, 5)
        annual_growth = calc_cagr(base["added_value"], year5, 5) * 100
        assert annual_growth > 0, f"年率 {annual_growth:.1f}% <= 0%"

    def test_operating_profit_positive_all_years(self):
        """★リグレッション: 全年度で営業利益がプラス"""
        Config.reset_rates()
        base = self._make_base()
        for yr in range(0, 6):
            projected_op = int(base["op_profit"] * Config.GROWTH_RATE ** yr)
            assert projected_op > 0, f"year {yr}: 営業利益 {projected_op:,} <= 0"


class TestCalcYearSalary:
    """calc_year_salary のユニットテスト"""

    def test_salary_growth_meets_requirement(self):
        """★リグレッション: SALARY_GROWTH_RATE=1.04でCAGR≥3.5%"""
        Config.reset_rates()
        data = HearingData()
        data.company.total_salary = 10_000_000
        data.company.revenue_2024 = 50_000_000
        data.company.employee_count = 2
        data.equipment.total_price = 10_000_000
        base = calc_base_components(data)
        year5 = calc_year_salary(base, 5)
        annual_growth = calc_cagr(base["salary"], year5, 5) * 100
        assert annual_growth >= 3.5, f"給与年率 {annual_growth:.1f}% < 3.5%"


class TestCalcCagr:
    """calc_cagr のユニットテスト"""

    def test_basic_cagr(self):
        """基本的なCAGR計算"""
        # 100 → 110 in 1 year = 10%
        cagr = calc_cagr(100, 110, 1)
        assert abs(cagr - 0.10) < 0.001

    def test_5_year_cagr(self):
        """5年間のCAGR"""
        # 100 * 1.05^5 ≈ 127.63
        cagr = calc_cagr(100, 100 * 1.05**5, 5)
        assert abs(cagr - 0.05) < 0.001

    def test_zero_base_returns_zero(self):
        """基準値がゼロの場合は0を返す"""
        assert calc_cagr(0, 100, 5) == 0.0

    def test_zero_years_returns_zero(self):
        """年数がゼロの場合は0を返す"""
        assert calc_cagr(100, 200, 0) == 0.0


class TestPerCapitaSalaryCagr:
    """calc_per_capita_salary_cagr のユニットテスト"""

    def test_per_capita_matches_salary_growth(self):
        """従業員数一定なら1人当たりCAGR = 給与総額CAGR"""
        Config.reset_rates()
        cagr = calc_per_capita_salary_cagr(10_000_000, 2, 5)
        expected = Config.SALARY_GROWTH_RATE - 1  # 0.04
        assert abs(cagr - expected) < 0.001

    def test_zero_employees_returns_zero(self):
        """従業員0の場合は0を返す"""
        assert calc_per_capita_salary_cagr(10_000_000, 0, 5) == 0.0

    def test_zero_salary_returns_zero(self):
        """給与0の場合は0を返す"""
        assert calc_per_capita_salary_cagr(0, 2, 5) == 0.0


class TestCheckRequirements:
    """check_requirements のユニットテスト"""

    def _make_data(self):
        data = HearingData()
        data.company.operating_profit_2024 = 2_275_980
        data.company.labor_cost = 5_560_971
        data.company.depreciation = 2_499_244
        data.company.revenue_2024 = 64_199_095
        data.company.total_salary = 2_494_000
        data.company.employee_count = 2
        data.equipment.total_price = 11_250_000
        return data

    def test_requirements_met_with_default_rates(self):
        """★リグレッション: デフォルト成長率で1人当たり給与要件充足"""
        Config.reset_rates()
        data = self._make_data()
        result = check_requirements(data)
        # 1人当たり給与CAGR ≥ 3.5%（これが今回の修正の主目的）
        assert result["salary_per_capita_ok"], f"1人当たり給与CAGR {result['salary_per_capita_cagr']*100:.2f}% < 3.5%"

    def test_requirements_met_high_profit_ratio(self):
        """付加価値額CAGR: 営業利益構成比が高い場合に≥3%"""
        Config.reset_rates()
        data = HearingData()
        data.company.operating_profit_2024 = 8_000_000  # 高め
        data.company.labor_cost = 3_000_000
        data.company.depreciation = 1_000_000
        data.company.revenue_2024 = 50_000_000
        data.company.total_salary = 2_494_000
        data.company.employee_count = 2
        data.equipment.total_price = 5_000_000
        result = check_requirements(data)
        assert result["added_value_ok"], f"付加価値額CAGR {result['added_value_cagr']*100:.2f}% < 3%"

    def test_low_salary_rate_fails(self):
        """SALARY_GROWTH_RATE=1.02では1人当たり給与要件を満たさない"""
        Config.reset_rates()
        Config.SALARY_GROWTH_RATE = 1.02
        data = self._make_data()
        result = check_requirements(data)
        assert not result["salary_per_capita_ok"]
        assert any("給与" in w for w in result["warnings"])
        Config.reset_rates()  # 元に戻す


class TestValidateFinancialInputs:
    """validate_financial_inputs のユニットテスト"""

    def test_normal_data_no_warnings(self):
        """正常データでは警告なし"""
        data = HearingData()
        data.company.revenue_2024 = 64_199_095
        data.company.operating_profit_2024 = 2_275_980
        data.company.labor_cost = 5_560_971
        data.company.depreciation = 2_499_244
        data.company.employee_count = 2
        data.equipment.total_price = 11_250_000
        warnings = validate_financial_inputs(data)
        assert len(warnings) == 0

    def test_sga_exceeds_revenue_warning(self):
        """人件費+減価償却費が売上高を超える場合に警告"""
        data = HearingData()
        data.company.revenue_2024 = 10_000_000
        data.company.labor_cost = 8_000_000
        data.company.depreciation = 5_000_000  # 合計13M > 売上10M
        data.company.employee_count = 2
        data.equipment.total_price = 5_000_000
        warnings = validate_financial_inputs(data)
        assert any("売上高" in w for w in warnings)

    def test_negative_profit_warning(self):
        """営業利益マイナスで警告"""
        data = HearingData()
        data.company.revenue_2024 = 50_000_000
        data.company.operating_profit_2024 = -1_000_000
        data.company.employee_count = 2
        data.equipment.total_price = 5_000_000
        warnings = validate_financial_inputs(data)
        assert any("マイナス" in w for w in warnings)

    def test_zero_employees_warning(self):
        """従業員数ゼロで警告"""
        data = HearingData()
        data.company.revenue_2024 = 50_000_000
        data.company.operating_profit_2024 = 5_000_000
        data.company.employee_count = 0
        data.equipment.total_price = 5_000_000
        warnings = validate_financial_inputs(data)
        assert any("従業員" in w for w in warnings)


class TestConfigReset:
    """Config.reset_rates() のテスト"""

    def test_growth_rate_reset(self):
        """reset_rates() でデフォルト値に戻る"""
        Config.GROWTH_RATE = 1.10
        Config.SALARY_GROWTH_RATE = 1.05
        Config.reset_rates()
        assert Config.GROWTH_RATE == 1.05
        assert Config.SALARY_GROWTH_RATE == 1.04  # ★更新: 1.025→1.04

    def test_reset_is_idempotent(self):
        """reset_rates() は何度呼んでも同じ"""
        Config.reset_rates()
        g1 = Config.GROWTH_RATE
        Config.reset_rates()
        g2 = Config.GROWTH_RATE
        assert g1 == g2
