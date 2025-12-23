"""
호텔 직원 데이터 모델
"""
from dataclasses import dataclass
from typing import Optional


@dataclass
class Employee:
    """호텔 직원 정보"""
    employee_id: str  # 직원 번호
    name: str  # 이름
    department: str  # 부서
    position: str  # 직급
    base_salary: int  # 기본급
    bonus: int = 0  # 상여금
    overtime_pay: int = 0  # 초과근무수당
    meal_allowance: int = 0  # 식대
    transportation_allowance: int = 0  # 교통비
    national_pension: int = 0  # 국민연금
    health_insurance: int = 0  # 건강보험
    employment_insurance: int = 0  # 고용보험
    income_tax: int = 0  # 소득세
    resident_tax: int = 0  # 주민세
    bank_name: str = ""  # 은행명
    account_number: str = ""  # 계좌번호

    def calculate_gross_salary(self) -> int:
        """총 지급액 계산"""
        return (
            self.base_salary
            + self.bonus
            + self.overtime_pay
            + self.meal_allowance
            + self.transportation_allowance
        )

    def calculate_total_deductions(self) -> int:
        """총 공제액 계산"""
        return (
            self.national_pension
            + self.health_insurance
            + self.employment_insurance
            + self.income_tax
            + self.resident_tax
        )

    def calculate_net_salary(self) -> int:
        """실수령액 계산"""
        return self.calculate_gross_salary() - self.calculate_total_deductions()
