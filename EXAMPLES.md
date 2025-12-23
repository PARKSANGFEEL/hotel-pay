# 사용 예시 (Examples)

## 기본 실행

```bash
python main.py
```

**출력:**
```
=== 호텔 직원 급여 자동 생성 시스템 ===
처리 기간: 2025년 12월

직원 수: 5명

1. 급여 명세서 생성 중...
   ✓ 급여 명세서 생성 완료: 급여명세서.xlsx

2. 급여이체 리스트 생성 중...
   ✓ 급여이체 리스트 생성 완료: 급여이체리스트.xlsx

=== 처리 결과 요약 ===
총 지급액: 19,130,000원
총 공제액: 3,888,000원
총 실수령액: 15,242,000원

처리가 완료되었습니다!
```

## 특정 연월 지정

```bash
python main.py --year 2024 --month 12
```

## 생성된 파일 설명

### 1. 급여명세서.xlsx

각 직원마다 별도의 시트가 생성되며, 다음 정보가 포함됩니다:

```
2025년 12월 급여 명세서

직원번호: H2024001        성명: 김호텔
부서: 프론트데스크        직급: 매니저

지급 항목          금액
기본급            3,500,000
상여금              500,000
초과근무수당        200,000
식대                150,000
교통비              100,000
지급 합계         4,450,000

공제 항목          금액
국민연금            315,000
건강보험            245,000
고용보험             35,000
소득세              280,000
주민세               28,000
공제 합계           903,000

실수령액          3,547,000
```

### 2. 급여이체리스트.xlsx

전체 직원의 급여 이체 정보가 한 시트에 정리됩니다:

```
2025년 12월 급여이체 리스트

번호  직원번호    성명      부서            직급    은행명      계좌번호              이체금액
1     H2024001   김호텔    프론트데스크    매니저  신한은행    110-123-456789       3,547,000
2     H2024002   이객실    하우스키핑      수석    국민은행    123456-01-123456     2,948,000
3     H2024003   박요리    식음료부        주방장  우리은행    1002-123-456789      4,135,000
4     H2024004   최서비스  프론트데스크    직원    하나은행    123-456789-12345     2,388,000
5     H2024005   정청소    하우스키핑      직원    기업은행    123-456-789012       2,224,000

합계                                                                               15,242,000
```

## 커스텀 직원 데이터 사용

`sample_data.py`의 `get_sample_employees()` 함수를 수정하거나, 새로운 함수를 만들어 사용할 수 있습니다:

```python
from employee import Employee
from payroll_statement import PayrollStatementGenerator
from payment_transfer_list import PaymentTransferListGenerator

# 직원 데이터 직접 생성
employees = [
    Employee(
        employee_id="H2025001",
        name="홍길동",
        department="객실관리부",
        position="팀장",
        base_salary=3800000,
        bonus=400000,
        overtime_pay=250000,
        meal_allowance=150000,
        transportation_allowance=100000,
        national_pension=342000,
        health_insurance=266000,
        employment_insurance=38000,
        income_tax=320000,
        resident_tax=32000,
        bank_name="카카오뱅크",
        account_number="3333-01-1234567"
    ),
]

# 급여 명세서 및 이체 리스트 생성
statement_gen = PayrollStatementGenerator(2025, 1)
statement_gen.generate(employees, "2025년1월_급여명세서.xlsx")

transfer_gen = PaymentTransferListGenerator(2025, 1)
transfer_gen.generate(employees, "2025년1월_급여이체리스트.xlsx")
```

## 테스트 실행

```bash
python test_payroll.py
```

**출력:**
```
=== 호텔 급여 시스템 테스트 시작 ===

Testing employee calculations...
✓ 총 지급액: 3,700,000원
✓ 총 공제액: 752,000원
✓ 실수령액: 2,948,000원
Employee calculations test passed!

Testing payroll statement generation...
✓ 급여 명세서 생성됨: test_급여명세서.xlsx
Payroll statement generation test passed!

Testing payment transfer list generation...
✓ 급여이체 리스트 생성됨: test_급여이체리스트.xlsx
Payment transfer list generation test passed!

=== 모든 테스트 통과! ===
```
