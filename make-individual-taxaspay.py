import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import shutil

# 현재 년월 설정
current_year = "2025"
current_month = "12"

print("=" * 80)
print("그리드인 직원별 taxaspay 파일 생성")
print("=" * 80)

# 1. 템플릿 파일 읽기
print("\n[1/4] taxaspay.xlsx 템플릿 읽기...")
df_template = pd.read_excel("taxaspay.xlsx", header=None)
print(f"  ✓ 템플릿 파일 읽기 완료 (Shape: {df_template.shape})")

# 2. 그리드인_텍사스_급여발송 데이터 읽기
print("\n[2/4] 그리드인_텍사스_급여발송 데이터 읽기...")
df_gridin_texas = pd.read_excel(f"그리드인_텍사스_급여발송_{current_year}년{current_month}월.xlsx", header=None)

# 직원 데이터는 6행부터 시작 (0-based index)
employee_data_start = 6
employee_list = []

for idx in range(employee_data_start, len(df_gridin_texas)):
    row = df_gridin_texas.iloc[idx]
    if pd.notna(row[0]) and row[0] != '':  # 이름이 있는 행
        employee_list.append({
            '이름': row[0],
            '주민등록번호': row[1] if pd.notna(row[1]) else '',
            '귀속월': row[2] if pd.notna(row[2]) else f'{current_year}{current_month}',
            '기본급': row[5] if pd.notna(row[5]) else 0,
            '상여': row[6] if pd.notna(row[6]) else 0,
            '식대': row[7] if pd.notna(row[7]) else 0,
            '육아수당': row[8] if pd.notna(row[8]) else 0,
            '연차수당': row[9] if pd.notna(row[9]) else 0,
            '성과급': row[10] if pd.notna(row[10]) else 0,
            '소급분': row[11] if pd.notna(row[11]) else 0,
            '경조사비': row[12] if pd.notna(row[12]) else 0,
            '국민연금': row[13] if pd.notna(row[13]) else 0,
            '건강보험': row[14] if pd.notna(row[14]) else 0,
            '장기요양보험': row[15] if pd.notna(row[15]) else 0,
            '고용보험': row[16] if pd.notna(row[16]) else 0,
            '소득세': row[17] if pd.notna(row[17]) else 0,
            '지방소득세': row[18] if pd.notna(row[18]) else 0
        })

print(f"  ✓ {len(employee_list)}명의 직원 데이터 읽기 완료")

# 3. 한 파일에 모든 직원 데이터 추가
print("\n[3/4] 급여명세서 데이터 구성 중...")

# 템플릿의 헤더 부분(0-5행) 유지
df_output = df_template.iloc[:6].copy()  # 헤더 6행까지

# 각 직원별 데이터 행 추가
for idx, emp in enumerate(employee_list, 1):
    # 새 행 생성
    new_row = [None] * len(df_template.columns)
    
    new_row[0] = emp['이름']  # 이름
    new_row[1] = emp['주민등록번호']  # 주민등록번호
    new_row[2] = emp['귀속월']  # 귀속월
    new_row[3] = 22  # 근무일수 (기본값)
    
    # 지급 항목
    new_row[4] = 0  # 근로시간 (기본값)
    new_row[5] = emp['기본급']  # 기본급 금액
    new_row[6] = emp['상여']  # 상여
    new_row[7] = emp['식대']  # 식대
    new_row[8] = 0  # 주휴수당
    new_row[9] = 0  # 연장수당
    new_row[10] = 0  # 야간수당
    new_row[11] = 0  # 미지급분
    new_row[12] = 0  # 휴일기본수당
    new_row[13] = emp['연차수당']  # 연차수당
    new_row[14] = 0  # 휴일근무수당
    
    # 공제 항목
    new_row[15] = emp['국민연금']  # 국민연금
    new_row[16] = emp['건강보험']  # 건강보험
    new_row[17] = emp['장기요양보험']  # 장기요양보험
    new_row[18] = emp['고용보험']  # 고용보험
    new_row[19] = emp['소득세']  # 소득세
    new_row[20] = emp['지방소득세']  # 지방소득세
    
    # 산출식 컬럼들 (21-30)
    new_row[21] = '(근로시간+주휴시간) x 통상시급'
    new_row[22] = '근로계약 및 관계규정에 의함'
    new_row[23] = '근로계약 및 관계규정에 의함'
    
    # 공제 항목 산출식 (31-36)
    new_row[31] = '각공단 부과금액'
    new_row[32] = '각공단 부과금액'
    new_row[33] = '각공단 부과금액'
    new_row[34] = '각공단 부과금액'
    new_row[35] = '근로소득간이세액표 기준으로 공제'
    new_row[36] = '소득세의 10%'
    
    # 데이터프레임에 행 추가
    df_output.loc[len(df_output)] = new_row
    print(f"  [{idx}/{len(employee_list)}] {emp['이름']} 님 데이터 추가 완료")

# 4. 파일 저장
print("\n[4/4] 파일 저장 중...")
output_filename = f"그리드인_급여명세서_{current_year}년{current_month}월.xlsx"

# 엑셀로 저장
df_output.to_excel(output_filename, index=False, header=False, engine='openpyxl')

print("\n" + "=" * 80)
print("✅ 처리 완료!")
print("=" * 80)
print(f"생성된 파일: {output_filename}")
print(f"포함된 직원 수: {len(employee_list)}명")
print("\n포함된 직원 목록:")
for emp in employee_list:
    print(f"  - {emp['이름']}")
print("=" * 80)
