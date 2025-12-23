"""
호텔 직원 급여 자동 생성 메인 스크립트
"""
from datetime import datetime
from payroll_statement import PayrollStatementGenerator
from payment_transfer_list import PaymentTransferListGenerator
from sample_data import get_sample_employees


def main():
    """메인 함수"""
    # 현재 연월
    now = datetime.now()
    year = now.year
    month = now.month
    
    print(f"=== 호텔 직원 급여 자동 생성 시스템 ===")
    print(f"처리 기간: {year}년 {month}월\n")
    
    # 샘플 직원 데이터 로드
    employees = get_sample_employees()
    print(f"직원 수: {len(employees)}명")
    
    # 급여 명세서 생성
    print("\n1. 급여 명세서 생성 중...")
    statement_generator = PayrollStatementGenerator(year, month)
    statement_path = statement_generator.generate(employees)
    print(f"   ✓ 급여 명세서 생성 완료: {statement_path}")
    
    # 급여이체 리스트 생성
    print("\n2. 급여이체 리스트 생성 중...")
    transfer_generator = PaymentTransferListGenerator(year, month)
    transfer_path = transfer_generator.generate(employees)
    print(f"   ✓ 급여이체 리스트 생성 완료: {transfer_path}")
    
    # 요약 정보 출력
    print("\n=== 처리 결과 요약 ===")
    total_gross = sum(emp.calculate_gross_salary() for emp in employees)
    total_deductions = sum(emp.calculate_total_deductions() for emp in employees)
    total_net = sum(emp.calculate_net_salary() for emp in employees)
    
    print(f"총 지급액: {total_gross:,}원")
    print(f"총 공제액: {total_deductions:,}원")
    print(f"총 실수령액: {total_net:,}원")
    
    print("\n처리가 완료되었습니다!")


if __name__ == "__main__":
    main()
