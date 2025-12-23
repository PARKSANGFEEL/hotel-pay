"""
호텔 급여 시스템 테스트
"""
import os
import sys
from employee import Employee
from payroll_statement import PayrollStatementGenerator
from payment_transfer_list import PaymentTransferListGenerator


def test_employee_calculations():
    """직원 급여 계산 테스트"""
    print("Testing employee calculations...")
    
    employee = Employee(
        employee_id="TEST001",
        name="테스트",
        department="테스트부서",
        position="직원",
        base_salary=3000000,
        bonus=300000,
        overtime_pay=150000,
        meal_allowance=150000,
        transportation_allowance=100000,
        national_pension=270000,
        health_insurance=210000,
        employment_insurance=30000,
        income_tax=220000,
        resident_tax=22000,
        bank_name="테스트은행",
        account_number="123-456-789"
    )
    
    # 총 지급액 확인
    gross = employee.calculate_gross_salary()
    expected_gross = 3000000 + 300000 + 150000 + 150000 + 100000
    assert gross == expected_gross, f"총 지급액 오류: {gross} != {expected_gross}"
    print(f"✓ 총 지급액: {gross:,}원")
    
    # 총 공제액 확인
    deductions = employee.calculate_total_deductions()
    expected_deductions = 270000 + 210000 + 30000 + 220000 + 22000
    assert deductions == expected_deductions, f"총 공제액 오류: {deductions} != {expected_deductions}"
    print(f"✓ 총 공제액: {deductions:,}원")
    
    # 실수령액 확인
    net = employee.calculate_net_salary()
    expected_net = expected_gross - expected_deductions
    assert net == expected_net, f"실수령액 오류: {net} != {expected_net}"
    print(f"✓ 실수령액: {net:,}원")
    
    print("Employee calculations test passed!\n")


def test_payroll_statement_generation():
    """급여 명세서 생성 테스트"""
    print("Testing payroll statement generation...")
    
    employees = [
        Employee(
            employee_id="TEST001",
            name="테스트1",
            department="테스트부서",
            position="직원",
            base_salary=3000000,
            bonus=300000,
            overtime_pay=150000,
            meal_allowance=150000,
            transportation_allowance=100000,
            national_pension=270000,
            health_insurance=210000,
            employment_insurance=30000,
            income_tax=220000,
            resident_tax=22000,
            bank_name="테스트은행",
            account_number="123-456-789"
        )
    ]
    
    generator = PayrollStatementGenerator(2025, 12)
    output_path = "test_급여명세서.xlsx"
    
    # 기존 파일이 있으면 삭제
    if os.path.exists(output_path):
        os.remove(output_path)
    
    # 생성
    result_path = generator.generate(employees, output_path)
    
    # 파일 생성 확인
    assert os.path.exists(result_path), f"파일이 생성되지 않음: {result_path}"
    print(f"✓ 급여 명세서 생성됨: {result_path}")
    
    # 정리
    os.remove(result_path)
    
    print("Payroll statement generation test passed!\n")


def test_payment_transfer_list_generation():
    """급여이체 리스트 생성 테스트"""
    print("Testing payment transfer list generation...")
    
    employees = [
        Employee(
            employee_id="TEST001",
            name="테스트1",
            department="테스트부서",
            position="직원",
            base_salary=3000000,
            bonus=300000,
            overtime_pay=150000,
            meal_allowance=150000,
            transportation_allowance=100000,
            national_pension=270000,
            health_insurance=210000,
            employment_insurance=30000,
            income_tax=220000,
            resident_tax=22000,
            bank_name="테스트은행",
            account_number="123-456-789"
        ),
        Employee(
            employee_id="TEST002",
            name="테스트2",
            department="테스트부서",
            position="매니저",
            base_salary=4000000,
            bonus=400000,
            overtime_pay=200000,
            meal_allowance=150000,
            transportation_allowance=100000,
            national_pension=360000,
            health_insurance=280000,
            employment_insurance=40000,
            income_tax=300000,
            resident_tax=30000,
            bank_name="테스트은행",
            account_number="987-654-321"
        )
    ]
    
    generator = PaymentTransferListGenerator(2025, 12)
    output_path = "test_급여이체리스트.xlsx"
    
    # 기존 파일이 있으면 삭제
    if os.path.exists(output_path):
        os.remove(output_path)
    
    # 생성
    result_path = generator.generate(employees, output_path)
    
    # 파일 생성 확인
    assert os.path.exists(result_path), f"파일이 생성되지 않음: {result_path}"
    print(f"✓ 급여이체 리스트 생성됨: {result_path}")
    
    # 정리
    os.remove(result_path)
    
    print("Payment transfer list generation test passed!\n")


def main():
    """모든 테스트 실행"""
    print("=== 호텔 급여 시스템 테스트 시작 ===\n")
    
    try:
        test_employee_calculations()
        test_payroll_statement_generation()
        test_payment_transfer_list_generation()
        
        print("=== 모든 테스트 통과! ===")
        return 0
    except AssertionError as e:
        print(f"\n❌ 테스트 실패: {e}")
        return 1
    except Exception as e:
        print(f"\n❌ 예상치 못한 오류: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
