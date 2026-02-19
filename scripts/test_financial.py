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
from financial_utils import calc_base_components, calc_year_added_value, calc_year_salary


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
    ):
        """テスト用のHearingDataを作成"""
        data = HearingData()
        data.company.operating_profit_2024 = op
        data.company.labor_cost = labor
        data.company.depreciation = dep
        data.company.revenue_2024 = revenue
        data.company.total_salary = salary
        data.equipment.total_price = equip_price
        return data

    def test_basic_calculation(self):
        """正常系: 実値がすべて設定されている場合"""
        data = self._make_data()
        base = calc_base_components(data)
        expected = 2_275_980 + 6_713_298 + 2_822_935  # = 11,812,213
        assert base["added_value"] == expected
        assert base["labor_cost"] == 6_713_298
        assert base["depreciation"] == 2_822_935
        assert base["op_profit"] == 2_275_980

    def test_fallback_labor_cost(self):
        """labor_cost=0の場合、revenue*0.35がfallbackされる"""
        data = self._make_data(labor=0)
        base = calc_base_components(data)
        expected_labor = int(64_199_095 * 0.35)
        assert base["labor_cost"] == expected_labor
        assert base["labor_cost"] != 0

    def test_fallback_depreciation(self):
        """depreciation=0の場合、equip_price/5がfallbackされる"""
        data = self._make_data(dep=0)
        base = calc_base_components(data)
        expected_dep = int(14_114_675 / 5)
        assert base["depreciation"] == expected_dep
        assert base["depreciation"] != 0

    def test_fallback_salary(self):
        """total_salary=0の場合、revenue*0.30がfallbackされる"""
        data = self._make_data(salary=0)
        base = calc_base_components(data)
        expected_salary = int(64_199_095 * 0.30)
        assert base["salary"] == expected_salary

    def test_added_value_formula(self):
        """付加価値額 = 営業利益 + 人件費 + 減価償却費"""
        data = self._make_data()
        base = calc_base_components(data)
        assert base["added_value"] == base["op_profit"] + base["labor_cost"] + base["depreciation"]

    def test_consistency_across_calls(self):
        """同じデータで複数回呼んでも同じ結果になる（冪等性）"""
        data = self._make_data()
        base1 = calc_base_components(data)
        base2 = calc_base_components(data)
        assert base1 == base2


class TestCalcYearAddedValue:
    """calc_year_added_value のユニットテスト"""

    def _make_base(self):
        data = HearingData()
        data.company.operating_profit_2024 = 2_275_980
        data.company.labor_cost = 6_713_298
        data.company.depreciation = 2_822_935
        data.company.revenue_2024 = 64_199_095
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
        Config.reset_rates()  # 5%にリセット
        base = self._make_base()
        year5 = calc_year_added_value(base, 5)
        annual_growth = ((year5 / base["added_value"]) ** (1 / 5) - 1) * 100
        # 各成分の成長率が異なるため（OP=5%, LC=2.5%, DEP=0%）
        # 全体の成長率はOP構成比に依存する。プラス成長であればOK。
        assert annual_growth > 0, f"年率 {annual_growth:.1f}% <= 0%"

    def test_high_profit_ratio_meets_4_percent(self):
        """営業利益の構成比が高い場合、年率4%以上を達成"""
        Config.reset_rates()
        data = HearingData()
        data.company.operating_profit_2024 = 10_000_000  # 高い営業利益
        data.company.labor_cost = 2_000_000
        data.company.depreciation = 1_000_000
        data.company.revenue_2024 = 50_000_000
        data.equipment.total_price = 5_000_000
        base = calc_base_components(data)
        year5 = calc_year_added_value(base, 5)
        annual_growth = ((year5 / base["added_value"]) ** (1 / 5) - 1) * 100
        assert annual_growth >= 4.0, f"年率 {annual_growth:.1f}% < 4%"


class TestCalcYearSalary:
    """calc_year_salary のユニットテスト"""

    def test_salary_growth(self):
        """給与支給総額が年率2%以上成長"""
        Config.reset_rates()  # 2.5%にリセット
        data = HearingData()
        data.company.total_salary = 10_000_000
        data.company.revenue_2024 = 50_000_000
        data.equipment.total_price = 10_000_000
        base = calc_base_components(data)
        year5 = calc_year_salary(base, 5)
        annual_growth = ((year5 / base["salary"]) ** (1 / 5) - 1) * 100
        assert annual_growth >= 2.0, f"給与年率 {annual_growth:.1f}% < 2%"


class TestConfigReset:
    """Config.reset_rates() のテスト"""

    def test_growth_rate_reset(self):
        """reset_rates() でデフォルト値に戻る"""
        Config.GROWTH_RATE = 1.10
        Config.SALARY_GROWTH_RATE = 1.05
        Config.reset_rates()
        assert Config.GROWTH_RATE == 1.05
        assert Config.SALARY_GROWTH_RATE == 1.025

    def test_reset_is_idempotent(self):
        """reset_rates() は何度呼んでも同じ"""
        Config.reset_rates()
        g1 = Config.GROWTH_RATE
        Config.reset_rates()
        g2 = Config.GROWTH_RATE
        assert g1 == g2
