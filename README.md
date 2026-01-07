# hotel-pay
호텔 직원들 급여 자동생성 시스템

## 개요
호텔 직원들의 급여 명세서와 급여이체 리스트를 자동으로 생성하는 시스템입니다.

## 주요 기능
- 📄 **급여 명세서 자동 생성**: 각 직원별 상세한 급여 명세서를 엑셀 파일로 생성
- 💰 **급여이체 리스트 자동 생성**: 은행 이체를 위한 전체 직원 급여 리스트 생성
- 📊 **자동 계산**: 지급액, 공제액, 실수령액 자동 계산

## 설치 방법

### 필수 요구사항
- Python 3.7 이상

### 설치
```bash
pip install -r requirements.txt
```

## 사용 방법

### 기본 사용
```bash
python main.py
```

위 명령어를 실행하면 현재 연월 기준으로 다음 파일들이 생성됩니다:
- `급여명세서.xlsx`: 각 직원별 급여 명세서 (직원마다 별도 시트)
- `급여이체리스트.xlsx`: 은행 이체를 위한 전체 직원 급여 리스트

### 특정 연월 지정
```bash
# 2025년 1월 급여 생성
python main.py --year 2025 --month 1

# 2024년 12월 급여 생성
python main.py --year 2024 --month 12
```

### 프로그래밍 방식으로 사용

```python
from datetime import datetime
from employee import Employee
from payroll_statement import PayrollStatementGenerator
from payment_transfer_list import PaymentTransferListGenerator

# 직원 데이터 생성
employees = [
    Employee(
        employee_id="H2024001",
        name="김호텔",
        department="프론트데스크",
        position="매니저",
        base_salary=3500000,
        bonus=500000,
        overtime_pay=200000,
        meal_allowance=150000,
        transportation_allowance=100000,
        national_pension=315000,
        health_insurance=245000,
        employment_insurance=35000,
        income_tax=280000,
        resident_tax=28000,
        bank_name="신한은행",
        account_number="110-123-456789"
    ),
    # 추가 직원...
]

# 급여 명세서 생성
year = datetime.now().year
month = datetime.now().month
statement_gen = PayrollStatementGenerator(year, month)
statement_gen.generate(employees, "급여명세서.xlsx")

# 급여이체 리스트 생성
transfer_gen = PaymentTransferListGenerator(year, month)
transfer_gen.generate(employees, "급여이체리스트.xlsx")
```

## 데이터 구조

### Employee 클래스
직원 정보를 담는 데이터 클래스입니다.

**지급 항목:**
- `base_salary`: 기본급
- `bonus`: 상여금
- `overtime_pay`: 초과근무수당
- `meal_allowance`: 식대
- `transportation_allowance`: 교통비

**공제 항목:**
- `national_pension`: 국민연금
- `health_insurance`: 건강보험
- `employment_insurance`: 고용보험
- `income_tax`: 소득세
- `resident_tax`: 주민세

**기본 정보:**
- `employee_id`: 직원번호
- `name`: 이름
- `department`: 부서
- `position`: 직급
- `bank_name`: 은행명
- `account_number`: 계좌번호

## 출력 예시

### 급여 명세서 (급여명세서.xlsx)
각 직원별로 별도의 시트가 생성되며, 다음 정보가 포함됩니다:
- 직원 기본 정보 (직원번호, 성명, 부서, 직급)
- 지급 항목 및 금액
- 공제 항목 및 금액
- 실수령액

### 급여이체 리스트 (급여이체리스트.xlsx)
은행 이체를 위한 전체 직원 리스트로, 다음 정보가 포함됩니다:
- 직원번호, 성명, 부서, 직급
- 은행명, 계좌번호
- 이체금액 (실수령액)
- 전체 합계

## 커스터마이징

### 샘플 데이터 수정
`sample_data.py` 파일을 수정하여 직원 데이터를 변경할 수 있습니다.

### 급여 계산 로직 수정
`employee.py` 파일의 계산 메서드를 수정하여 급여 계산 로직을 변경할 수 있습니다.

## 라이선스
MIT License
