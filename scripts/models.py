#!/usr/bin/env python3
"""データクラス定義"""

from dataclasses import dataclass, field
from typing import List


@dataclass
class CompanyInfo:
    """企業基本情報"""
    name: str = ""
    representative: str = ""
    address: str = ""
    prefecture: str = ""
    postal_code: str = ""
    phone: str = ""
    established_date: str = ""
    capital: int = 0
    industry: str = ""
    business_description: str = ""
    employee_count: int = 0
    officer_count: int = 1
    url: str = ""
    # 財務情報
    revenue_2022: int = 0
    revenue_2023: int = 0
    revenue_2024: int = 0
    gross_profit_2022: int = 0
    gross_profit_2023: int = 0
    gross_profit_2024: int = 0
    operating_profit_2022: int = 0
    operating_profit_2023: int = 0
    operating_profit_2024: int = 0
    # 付加価値額算出用（決算書PDFから取得）— 直近期
    labor_cost: int = 0           # 人件費合計
    depreciation: int = 0         # 減価償却費
    total_salary: int = 0         # 給与支給総額（役員報酬除く）
    # 前期（2期分アップロード時に使用）
    labor_cost_prev: int = 0      # 前期 人件費合計
    depreciation_prev: int = 0    # 前期 減価償却費
    total_salary_prev: int = 0    # 前期 給与支給総額
    # 決算期ラベル
    fiscal_period_latest: str = ""   # 直近期の決算期（例: "2024年3月期"）
    fiscal_period_prev: str = ""     # 前期の決算期（例: "2023年3月期"）


@dataclass
class LaborShortageInfo:
    """人手不足情報"""
    shortage_tasks: str = ""
    recruitment_period: str = ""
    applications: int = 0
    hired: int = 0
    overtime_hours: float = 0
    current_workers: int = 0
    desired_workers: int = 0
    job_openings_ratio: float = 0


@dataclass
class LaborSavingInfo:
    """省力化効果情報"""
    target_tasks: str = ""
    current_hours: float = 0
    target_hours: float = 0
    reduction_hours: float = 0
    reduction_rate: float = 0


@dataclass
class EquipmentInfo:
    """導入設備情報"""
    name: str = ""
    category: str = ""
    manufacturer: str = ""
    model: str = ""
    quantity: int = 1
    total_price: int = 0
    vendor: str = ""
    features: str = ""
    catalog_number: str = ""


@dataclass
class FundingInfo:
    """資金調達情報"""
    subsidy_amount: int = 0
    self_funding: int = 0
    total_investment: int = 0
    implementation_manager: str = ""
    implementation_period: str = ""
    bank_name: str = ""


@dataclass
class WorkProcess:
    """作業工程"""
    name: str = ""
    time_minutes: int = 0
    description: str = ""


@dataclass
class OfficerInfo:
    """役員情報"""
    name: str = ""
    position: str = ""
    birth_date: str = ""


@dataclass
class EmployeeInfo:
    """従業員情報"""
    name: str = ""
    birth_date: str = ""
    hire_date: str = ""


@dataclass
class ShareholderInfo:
    """株主情報"""
    name: str = ""
    shares: int = 0


@dataclass
class HearingData:
    """ヒアリングデータ全体"""
    company: CompanyInfo = field(default_factory=CompanyInfo)
    labor_shortage: LaborShortageInfo = field(default_factory=LaborShortageInfo)
    labor_saving: LaborSavingInfo = field(default_factory=LaborSavingInfo)
    equipment: EquipmentInfo = field(default_factory=EquipmentInfo)
    funding: FundingInfo = field(default_factory=FundingInfo)
    officers: List[OfficerInfo] = field(default_factory=list)
    employees: List[EmployeeInfo] = field(default_factory=list)
    shareholders: List[ShareholderInfo] = field(default_factory=list)
    before_processes: List[WorkProcess] = field(default_factory=list)
    after_processes: List[WorkProcess] = field(default_factory=list)
    # Phase 4: 追加フィールド
    motivation_background: str = ""  # なぜ今必要か（シート3）
    time_utilization_plan: str = ""  # 効果の活用計画（シート6）
    wage_increase_rate: float = 0.0  # 賃上げ率（シート7）
    wage_increase_target: str = ""  # 賃上げ対象者（シート7）
    wage_increase_timing: str = ""  # 賃上げ実施時期（シート7）
