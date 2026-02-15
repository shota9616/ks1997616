#!/usr/bin/env python3
"""
事業計画テキスト生成（PREP法）

【編集ガイド】
事業計画書の文章テンプレートを変更したい場合はこのファイルを編集してください。
各メソッドが1つのセクションに対応しています。

利用可能な変数:
  self.c  = CompanyInfo（企業名、業種、従業員数、財務情報など）
  self.s  = LaborShortageInfo（人手不足情報、残業時間など）
  self.l  = LaborSavingInfo（省力化効果、削減時間など）
  self.e  = EquipmentInfo（設備名、価格、特徴など）
  self.f  = FundingInfo（補助金額、自己資金など）
  self.data = HearingData（全データ、工程データ含む）
"""

from models import HearingData
from config import Config


class ContentGenerator:
    """採択レベルの文章を生成するクラス"""

    def __init__(self, data: HearingData):
        self.data = data
        self.c = data.company
        self.s = data.labor_shortage
        self.l = data.labor_saving
        self.e = data.equipment
        self.f = data.funding
        # Phase 2: Config参照 + _get_default_job_ratio メソッド
        self.job_ratio = self.s.job_openings_ratio if self.s.job_openings_ratio > 0 else self._get_default_job_ratio()
        self.manufacturer = self.e.manufacturer if self.e.manufacturer else "オーダーメイド開発"
        self.model = self.e.model if self.e.model else "カスタム仕様"

    def _get_default_job_ratio(self) -> float:
        """業種別デフォルト有効求人倍率を取得（Phase 2）"""
        for keyword, ratio in Config.INDUSTRY_JOB_RATIOS.items():
            if keyword in self.c.industry:
                return ratio
        return Config.DEFAULT_JOB_RATIO

    def _get_industry_philosophy(self) -> str:
        """業種別経営理念テンプレートを取得（Phase 3）"""
        for keyword, template in Config.INDUSTRY_PHILOSOPHY_TEMPLATES.items():
            if keyword in self.c.industry:
                return template
        return Config.DEFAULT_PHILOSOPHY_TEMPLATE.format(industry=self.c.industry)

    def generate_business_overview_table_data(self) -> dict:
        """事業者概要テーブル用のデータを生成"""
        return {
            "事業者名": self.c.name,
            # Phase 3: 業種別経営理念テンプレート
            "経営理念": self._get_industry_philosophy(),
            "経営戦略": f"{self.c.industry}として、{self.c.business_description}を専門に、高品質なサービスで顧客満足を追求。デジタル化・AI活用による業務効率向上で競争力を強化し、限られた人員で最大の成果を創出する戦略を推進。",
            "事業コンセプト": f"対象エリア：{self.c.prefecture}を中心とした地域。ターゲット：{self.c.industry}サービスを必要とする個人・法人顧客。提供サービス：{self.c.business_description}。強み：専門技術と豊富な経験に基づく高品質サービス。",
            "事業内容": f"①{self.c.business_description}の提供\n②差別化ポイント：専門資格者による高品質サービス、地域特性への深い理解\n③顧客価値：専門性の高いサービス提供、迅速な対応、長期的な信頼関係構築",
            "長期的なビジョン": f"5年後：{self.e.name}の活用による業務効率化を完了し、受注能力を1.5倍に拡大。従業員の働き方改革を実現。10年後：{self.c.prefecture}地域でトップクラスの{self.c.industry}事業者を目指し、後継者育成と事業承継の基盤を確立する。",
            "直近実績": {
                "売上金額": [self.c.revenue_2022, self.c.revenue_2023, self.c.revenue_2024],
                "売上総利益": [self.c.gross_profit_2022, self.c.gross_profit_2023, self.c.gross_profit_2024],
                "営業利益": [self.c.operating_profit_2022, self.c.operating_profit_2023, self.c.operating_profit_2024],
                "従業員数": [self.c.employee_count, self.c.employee_count, self.c.employee_count],
            }
        }

    def generate_section_1_1(self) -> str:
        """1-1 現状分析（PREP法、600字以上）"""
        added_value_2024 = self.c.operating_profit_2024 + int(self.c.revenue_2024 * Config.LABOR_COST_RATIO) + self.c.depreciation

        return f"""当社{self.c.name}は、{self.c.established_date}の設立以来、{self.c.prefecture}を拠点として{self.c.industry}を営む企業である。主たる事業内容は{self.c.business_description}であり、現在、役員{self.c.officer_count}名、従業員{self.c.employee_count}名の体制で事業を運営している。

当社の経営を取り巻く環境は、近年大きく変化している。市場環境においては、{self.c.industry}に対する需要は堅調に推移しており、当社の売上高は2022年度{self.c.revenue_2022:,}円、2023年度{self.c.revenue_2023:,}円、2024年度{self.c.revenue_2024:,}円と着実に成長を遂げている。営業利益についても2022年度{self.c.operating_profit_2022:,}円、2023年度{self.c.operating_profit_2023:,}円、2024年度{self.c.operating_profit_2024:,}円と堅調に推移しており、当社の技術力と顧客からの信頼が数字として表れている。

しかしながら、事業成長を支える人材の確保については極めて厳しい状況に直面している。{self.c.industry}における有効求人倍率は{self.job_ratio}倍と高水準で推移しており、必要な人材を確保することが年々困難になっている。当社においても、{self.s.recruitment_period}にわたり継続的に求人活動を実施しているものの、{"応募者が極めて少なく" if self.s.applications == 0 else f"応募者数は{self.s.applications}名にとどまり"}、{"採用に至った人材は皆無であり" if self.s.hired == 0 else f"実際に採用に至ったのは{self.s.hired}名という"}厳しい結果となっている。

このような人手不足の状況下において、当社の競争力の源泉である技術力と品質を維持しながら、増加する顧客ニーズに対応していくためには、業務の省力化・効率化が不可欠な経営課題となっている。"""

    def generate_swot_analysis(self) -> str:
        """SWOT分析を生成"""
        return f"""【SWOT分析】

■強み（Strengths）
当社の最大の強みは、{self.c.established_date}の設立以来培ってきた{self.c.industry}における専門的な技術力とノウハウである。{self.c.business_description}に関する長年の経験に裏打ちされた高品質なサービス提供により、顧客からの厚い信頼を獲得している。また、役員{self.c.officer_count}名、従業員{self.c.employee_count}名という機動力のある組織体制により、顧客ニーズへの迅速な対応が可能である。

■弱み（Weaknesses）
一方で、{self.s.shortage_tasks}における業務効率の低さが課題である。従来型の手作業に依存した業務プロセスでは、1件あたりの作業時間が長く、増加する需要に十分対応できていない。また、慢性的な人手不足により、従業員への負担が過大となっている。

■機会（Opportunities）
省力化投資補助金を活用した{self.e.name}の導入は、当社にとって業務改革を実現する絶好の機会である。AI・デジタル技術の進展により、これまで自動化が困難であった業務も効率化が可能となっている。

■脅威（Threats）
{self.c.industry}における有効求人倍率は{self.job_ratio}倍と高水準で推移しており、人材確保の競争は今後さらに激化すると予想される。また、同業他社もデジタル化・省力化を進めており、対応が遅れれば競争力を失うリスクがある。"""

    def generate_section_1_2(self) -> str:
        """1-2 経営上の課題（PREP法、700字以上）"""
        return f"""当社が直面している最も深刻な経営課題は、慢性的な人手不足とそれに起因する従業員の過重労働である。

現在、{self.s.shortage_tasks}の業務を担当しているのは{self.s.current_workers}名であるが、業務量に対して適正な人員は{self.s.desired_workers}名が必要と考えている。すなわち、現状では{max(0, self.s.desired_workers - self.s.current_workers)}名の人員が不足している状態で業務を遂行せざるを得ない状況にある。

この人員不足を補うため、現場の従業員は月平均{self.s.overtime_hours}時間の残業を余儀なくされている。この数値は、厚生労働省が定める時間外労働の上限規制である月45時間に迫る水準であり、従業員の健康管理の観点からも早急な改善が求められている。長時間労働の常態化は、従業員の疲労蓄積による作業効率の低下を招くだけでなく、ミスや事故のリスクを高め、最悪の場合には貴重な人材の離職につながりかねない。

特に深刻なのは、{self.s.shortage_tasks}における作業負担である。この業務は従来、熟練した従業員の経験と勘に依存しており、1件あたり{self.l.current_hours}時間もの作業時間を要している。案件数の増加に伴い、この作業に費やす時間が増大し、他の重要業務に充てる時間が圧迫されている状況である。

さらに、新規人材の採用が困難な状況が続く中、既存従業員の高齢化も進行しており、技術やノウハウの継承という観点からも、早急に業務プロセスの見直しと省力化を図る必要性が高まっている。このまま対策を講じなければ、当社の事業継続そのものが危ぶまれる事態に陥りかねない。"""

    def generate_section_1_3(self) -> str:
        """1-3 動機・目的（PREP法、400字以上）"""
        # Phase 4: motivation_background を反映
        motivation_text = ""
        if self.data.motivation_background:
            motivation_text = f"\n\n本設備導入を決断した背景として、{self.data.motivation_background}という事情がある。"

        return f"""上記の経営課題を解決するため、当社は{self.e.name}の導入を決断した。{motivation_text}

本設備導入の最大の目的は、{self.s.shortage_tasks}における作業時間を大幅に削減し、従業員の過重労働を解消することにある。具体的には、現在1件あたり{self.l.current_hours}時間を要している作業を、本設備の導入により{self.l.target_hours}時間まで短縮することを目指している。これにより、作業時間を{self.l.reduction_rate:.0f}%削減し、月{self.s.overtime_hours}時間に及ぶ残業時間の大幅な圧縮を実現する。

省力化により創出された時間は、より付加価値の高い業務に充当する計画である。従業員が本来の専門性を発揮できる環境を整備することで、サービス品質の向上と顧客満足度の向上を図り、ひいては売上拡大と利益率の改善につなげていく。また、労働環境の改善は従業員の定着率向上にも寄与し、人材確保の面でもプラスの効果が期待できる。

本補助金を活用することで、当社の経営基盤を強化し、持続可能な成長を実現したい。"""

    def generate_section_2_1(self) -> str:
        """2-1 ビフォーアフター（PREP法、1000字以上）"""
        before_total = sum(p.time_minutes for p in self.data.before_processes)
        after_total = sum(p.time_minutes for p in self.data.after_processes)
        reduction_minutes = before_total - after_total

        text = f"""本事業において導入する{self.e.name}について、導入前後の業務プロセスの変化を詳細に説明する。

【導入前の業務プロセス】
現在、{self.s.shortage_tasks}の業務は、以下のプロセスで実施している。"""

        for p in self.data.before_processes:
            text += f"\n「{p.name}」工程では、{p.description}を行っており、所要時間は{p.time_minutes}分である。"

        text += f"""

これらの工程を合計すると、1サイクルあたり{before_total}分（約{before_total/60:.1f}時間）を要している。この作業を1日に複数回実施するため、{self.s.shortage_tasks}だけで1日あたり{self.l.current_hours}時間もの時間を費やしている状況である。作業の大部分は従業員の手作業に依存しており、膨大な資料との照合作業が必要となり、従業員の負担が極めて大きい。

【導入後の業務プロセス】
{self.e.name}を導入することで、業務プロセスは以下のように変化する。"""

        for p in self.data.after_processes:
            text += f"\n「{p.name}」工程は、{p.description}により{p.time_minutes}分で完了する。"

        # Phase 1: ゼロ除算防止
        reduction_pct = (reduction_minutes / before_total * 100) if before_total > 0 else 0

        text += f"""

導入後の合計所要時間は{after_total}分（約{after_total/60:.1f}時間）となる。導入前と比較して、{reduction_minutes}分（約{reduction_minutes/60:.1f}時間）の短縮、削減率にして{reduction_pct:.0f}%の省力化を実現する。

【工程別の省力化効果】
各工程における具体的な省力化効果は以下のとおりである。"""

        # 工程別の詳細分析
        process_pairs = list(zip(self.data.before_processes, self.data.after_processes))
        for bp, ap in process_pairs:
            saved = bp.time_minutes - ap.time_minutes
            if saved > 0:
                pct = saved / bp.time_minutes * 100 if bp.time_minutes > 0 else 0
                text += f"\n・「{bp.name}」工程：{bp.time_minutes}分→{ap.time_minutes}分（{saved}分削減、{pct:.0f}%減）。従来の{bp.description}を{ap.description}に置き換えることで効率化される。"
            else:
                text += f"\n・「{bp.name}」工程：{bp.time_minutes}分→{ap.time_minutes}分。本工程は人間の判断が必要であり、所要時間に変化はない。"

        # 最も効果の大きい工程を特定
        biggest = max(process_pairs, key=lambda pair: pair[0].time_minutes - pair[1].time_minutes)

        text += f"""

【省力化の仕組み】
{self.e.name}の主要機能として、{self.e.features}が挙げられる。特に「{biggest[0].name}」工程においては、従来{biggest[0].description}に{biggest[0].time_minutes}分を要していたが、本設備の{biggest[1].description}機能により{biggest[1].time_minutes}分まで短縮される。これが本事業における最大の省力化ポイントである。

本設備の導入により、従業員は定型的・反復的な作業から解放され、顧客対応や品質管理といった人間の判断力が求められる高付加価値業務に集中できるようになる。1日あたりの削減時間は{self.l.reduction_hours:.1f}時間となり、月間では約{self.l.reduction_hours * Config.WORKING_DAYS_PER_MONTH:.0f}時間の業務時間を創出できる。"""

        return text

    def generate_section_2_2(self) -> str:
        """2-2 効果（PREP法、600字以上）"""
        # Phase 2: Config参照
        annual_saving = int(self.l.reduction_hours * Config.WORKING_DAYS_PER_MONTH * 12 * Config.HOURLY_WAGE)
        # Phase 4: time_utilization_plan を反映
        utilization_text = ""
        if self.data.time_utilization_plan:
            utilization_text = f"具体的には、{self.data.time_utilization_plan}に充てる計画である。"

        return f"""本事業の実施により期待される効果について、定量的・定性的の両面から説明する。

【定量的効果】
作業時間の削減効果として、1日あたり{self.l.reduction_hours:.1f}時間、月間では約{self.l.reduction_hours * Config.WORKING_DAYS_PER_MONTH:.0f}時間の業務時間を創出できる。この時間を人件費に換算すると、時給{Config.HOURLY_WAGE:,}円として年間約{annual_saving:,}円相当の効果となる。また、残業時間の削減により、割増賃金の支出も抑制される。現状の月{self.s.overtime_hours}時間の残業を半減できれば、年間で相当額の人件費削減が見込まれる。

【定性的効果】
まず、従業員の労働環境が大幅に改善される。長時間労働の解消により、従業員のワークライフバランスが向上し、心身の健康維持に寄与する。これは従業員の定着率向上につながり、採用難が続く現状において極めて重要な効果である。

次に、業務品質の安定化が期待できる。手作業に依存していた工程を自動化することで、ヒューマンエラーのリスクが大幅に低減される。一定の品質を安定して提供できることは、顧客からの信頼向上につながる。

さらに、創出された時間を活用して、より付加価値の高いサービスの提供や、新規顧客の開拓に注力することが可能となる。{utilization_text}これにより、売上の拡大と利益率の向上を実現し、持続的な事業成長の基盤を構築できる。"""

    def generate_section_3_1(self) -> str:
        """3-1 生産性向上（PREP法、700字以上）"""
        # Phase 2: Config参照
        base_added_value = self.c.operating_profit_2024 + int(self.c.revenue_2024 * Config.LABOR_COST_RATIO) + self.c.depreciation
        growth = Config.GROWTH_RATE

        # Phase 4: 賃上げ計画データの反映
        wage_detail = ""
        if self.data.wage_increase_rate > 0:
            wage_detail = f"当社は賃上げ率{self.data.wage_increase_rate}%を計画しており、"
            if self.data.wage_increase_target:
                wage_detail += f"対象は{self.data.wage_increase_target}、"
            if self.data.wage_increase_timing:
                wage_detail += f"{self.data.wage_increase_timing}より実施予定である。"
            else:
                wage_detail += "次年度より実施予定である。"

        growth_pct = (Config.GROWTH_RATE - 1) * 100
        salary_growth_pct = (Config.SALARY_GROWTH_RATE - 1) * 100

        return f"""本事業の実施により、当社は付加価値額の年率{growth_pct:.0f}%以上の向上を目指す。

【付加価値額の向上計画】
当社の付加価値額（営業利益＋人件費＋減価償却費）は、直近の2024年度実績で約{base_added_value:,}円である。本事業により省力化を実現し、業務効率を向上させることで、より多くの案件に対応可能となる。これにより、売上高の拡大を図りながら、付加価値額を毎年{growth_pct:.0f}%以上成長させていく計画である。

5年間の付加価値額推移の計画は以下のとおりである。
基準年度：約{base_added_value:,}円
1年目：約{int(base_added_value * growth):,}円（前年比+{growth_pct:.1f}%）
2年目：約{int(base_added_value * growth ** 2):,}円（前年比+{growth_pct:.1f}%）
3年目：約{int(base_added_value * growth ** 3):,}円（前年比+{growth_pct:.1f}%）
4年目：約{int(base_added_value * growth ** 4):,}円（前年比+{growth_pct:.1f}%）
5年目：約{int(base_added_value * growth ** 5):,}円（前年比+{growth_pct:.1f}%）

【給与支給総額の向上計画】
生産性向上により創出した利益の一部を原資として、従業員への還元を行う。具体的には、1人当たり給与支給総額の年平均成長率{salary_growth_pct:.1f}%以上を達成する計画である。{wage_detail}

【事業場内最低賃金の引上げ】
当社は、事業場内最低賃金について、{self.c.prefecture}の地域別最低賃金を30円以上上回る水準を維持することを表明する。

【投資回収計画】
本設備への投資額{self.f.total_investment:,}円は、省力化による人件費削減効果と売上拡大による利益増加により、約2〜3年で回収できる見込みである。"""
