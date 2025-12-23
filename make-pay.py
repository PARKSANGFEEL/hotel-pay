import pandas as pd
import glob
from openpyxl import load_workbook
import calendar

# 현재 년월 설정 (필요시 수정 가능)
current_year = "2025"
current_month = "12"

# 해당 월의 일수 계산
days_in_month = calendar.monthrange(int(current_year), int(current_month))[1]

# 파일 경로 - 그리드인_급여대장으로 시작하는 파일 자동 찾기
gridin_files = glob.glob("그리드인_급여대장*.xlsx")
if not gridin_files:
    print("❌ 오류: '그리드인_급여대장'으로 시작하는 파일을 찾을 수 없습니다.")
    exit(1)

input_file = gridin_files[0]  # 첫 번째 매칭 파일 사용
template_file = "taxaspay_최종본.xlsx"  # 최종본 템플릿 사용
output_file = f"그리드인_텍사스_급여발송_{current_year}년{current_month}월.xlsx"

print("=" * 80)
print("그리드인 급여 데이터 처리 시작")
print("=" * 80)
print(f"\n📂 입력 파일: {input_file}")

# 1. 그리드인_급여대장 읽기
print("\n[1/3] 그리드인_급여대장.xlsx 읽기...")
df_gridin = pd.read_excel(input_file, header=None)

# Worker.xlsx에서 사원 정보 읽기 (주민등록번호)
try:
    df_worker_raw = pd.read_excel("worker.xlsx", header=None)
    
    worker_data = []
    
    for idx in range(len(df_worker_raw)):
        row = df_worker_raw.iloc[idx]
        if pd.notna(row[0]) and row[0] != '':
            name = row[0]
            
            # row[1]과 row[2]를 확인하여 은행/계좌번호 여부 판단
            if pd.notna(row[1]) and not pd.isna(row[1]):
                # 은행 정보 있음
                bank = str(row[1]).strip()
                account = str(row[2]).strip() if pd.notna(row[2]) else "확인필요"
                resident = str(row[3]).strip() if pd.notna(row[3]) else "확인필요"
            else:
                # 은행 정보 없음
                bank = "확인필요"
                account = "확인필요"
                resident = str(row[3]).strip() if pd.notna(row[3]) else "확인필요"
            
            # E열(4): 주당근무일자, F열(5): 일당근무시간
            days_per_week = int(row[4]) if pd.notna(row[4]) else 0
            hours_per_day = int(row[5]) if pd.notna(row[5]) else 0
            
            worker_data.append({
                '이름': name,
                '은행': bank,
                '계좌번호': account,
                '주민등록번호': resident,
                '주당근무일자': days_per_week,
                '일당근무시간': hours_per_day
            })
    
    df_worker = pd.DataFrame(worker_data)
    print(f"  ✓ worker.xlsx 파일에서 {len(df_worker)}명의 정보를 읽었습니다.")
    
except Exception as e:
    print(f"  ⚠ worker.xlsx 파일 읽기 오류: {e}")
    df_worker = None

# 데이터 추출 (2행부터, 합계 행 제외)
data_rows = []
employee_list = []

for idx in range(2, len(df_gridin)):
    row = df_gridin.iloc[idx]
    # 사원코드가 숫자인 행만 처리 (합계 행 제외)
    if pd.notna(row[0]) and str(row[0]).replace('.', '').isdigit():
        employee_name = row[1]  # 사원명
        
        # 이름에서 (재) 제거
        employee_name = employee_name.replace('(재)', '').strip()
        
        # worker.xlsx에서 주민등록번호와 근무 정보 찾기
        resident_number = "확인필요"
        days_per_week = 0
        hours_per_day = 0
        if df_worker is not None:
            worker_info = df_worker[df_worker['이름'] == employee_name]
            if not worker_info.empty:
                resident_number = worker_info.iloc[0]['주민등록번호']
                days_per_week = worker_info.iloc[0]['주당근무일자']
                hours_per_day = worker_info.iloc[0]['일당근무시간']
        
        # 근무일수 계산: 주당 7일 근무면 해당 월 전체 일수, 아니면 비율 계산
        if days_per_week == 7:
            work_days = days_in_month
        else:
            work_days = int((days_in_month / 7) * days_per_week)
        
        # 근로시간 계산: 근무일수 * 일당근무시간
        work_hours = work_days * hours_per_day
        
        # 숫자 변환 함수
        def to_int(value):
            if pd.isna(value) or value == 0 or value == '':
                return 0
            try:
                if isinstance(value, str):
                    value = value.replace(',', '')
                return int(float(value))
            except:
                return 0
        
        # df_gridin.iloc[1]의 컬럼 인덱스:
        # 5:기본급, 6:상여, 7:식대, 8:주휴수당, 9:연장수당, 10:야간수당
        # 11:미지급분, 12:휴일기본수당, 13:연차수당, 14:휴일근로수당
        # 16:국민연금, 17:건강보험, 18:고용보험, 19:장기요양보험료
        # 20:소득세, 21:지방소득세
        
        employee_list.append({
            '이름': employee_name,
            '주민등록번호': resident_number,
            '귀속월': f'{current_year}{current_month}',
            '근무일수': work_days,
            '근로시간': work_hours,
            '기본급': to_int(row[5]),
            '상여': to_int(row[6]),
            '식대': to_int(row[7]),
            '주휴수당': to_int(row[8]),
            '연장수당': to_int(row[9]),
            '야간수당': to_int(row[10]),
            '미지급분': to_int(row[11]),
            '휴일기본수당': to_int(row[12]),
            '연차수당': to_int(row[13]),
            '휴일근무수당': to_int(row[14]),
            '국민연금': to_int(row[16]),
            '건강보험': to_int(row[17]),
            '고용보험': to_int(row[18]),
            '장기요양보험': to_int(row[19]),
            '소득세': to_int(row[20]),
            '지방소득세': to_int(row[21])
        })

print(f"  ✓ {len(employee_list)}명의 급여 데이터를 읽었습니다.")

# 2. 그리드인_직업급여이체계좌리스트 파일 생성
print("\n[2/3] 그리드인_직업급여이체계좌리스트 파일 생성 중...")

account_list_file = f"그리드인_급여이체_{current_year}년{current_month}월.xlsx"

# 계좌 리스트 데이터 준비
account_data = []
for idx, emp in enumerate(employee_list, 1):
    # worker.xlsx에서 은행과 계좌번호 찾기
    bank = "확인필요"
    account = "확인필요"
    total_amount = (emp['기본급'] + emp['상여'] + emp['식대'] + emp['주휴수당'] + 
                   emp['연장수당'] + emp['야간수당'] + emp['미지급분'] + 
                   emp['휴일기본수당'] + emp['연차수당'] + emp['휴일근무수당'] -
                   emp['국민연금'] - emp['건강보험'] - emp['고용보험'] - 
                   emp['장기요양보험'] - emp['소득세'] - emp['지방소득세'])
    
    if df_worker is not None:
        worker_info = df_worker[df_worker['이름'] == emp['이름']]
        if not worker_info.empty:
            bank = worker_info.iloc[0]['은행']
            account = worker_info.iloc[0]['계좌번호']
    
    # 이름을 영문으로 변환 (특정 이름만)
    display_name = emp['이름']
    if display_name == '김알비나':
        display_name = 'KIMALBINA'
    elif display_name == '엄스베틀라나':
        display_name = 'EMSVETLANA'
    elif display_name == '김엘레나':
        display_name = 'KIMELENA'
    
    account_data.append({
        'A_은행명': bank,
        'B_계좌번호': account,
        'C_금액': total_amount,
        'D_이름': display_name,
        'E_급여': '급여',
        'F_급여이체해당월': f'급여이체{current_month}',
        'G_해당월급여': f'{current_month}월급여',
        'H_순번': idx
    })

df_account = pd.DataFrame(account_data)

# 엑셀로 저장
with pd.ExcelWriter(account_list_file, engine='openpyxl') as writer:
    df_account.to_excel(writer, index=False, header=False)
    
    # 워크북과 시트 가져오기
    wb_account = writer.book
    ws_account = writer.sheets['Sheet1']
    
    # C열(금액)에 쉼표 포맷 적용
    for row_idx in range(1, len(account_data) + 1):
        ws_account.cell(row_idx, 3).number_format = '#,##0'

print(f"  ✓ 계좌 리스트 파일 생성 완료: {account_list_file}")
print(f"  ✓ {len(account_data)}개 계좌 정보 포함")

# 3. taxaspay_최종본.xlsx 파일을 직접 열어서 작업
print("\n[3/3] taxaspay_최종본.xlsx 파일에 데이터 추가 중...")

# 템플릿 파일을 직접 로드
wb = load_workbook(template_file)
ws = wb.active

# 템플릿의 7번 행(첫 데이터 행)을 기준으로 복사
template_row_num = 7

# 각 직원별로 행 추가
for idx, emp in enumerate(employee_list, 1):
    # 새 행 번호 (7행부터 시작)
    new_row_num = 6 + idx
    
    # 템플릿 7행의 모든 셀을 새 행으로 복사 (스타일 포함)
    for col in range(1, ws.max_column + 1):
        source_cell = ws.cell(template_row_num, col)
        target_cell = ws.cell(new_row_num, col)
        
        # 값과 스타일 복사
        target_cell.value = source_cell.value
        if source_cell.has_style:
            target_cell.font = source_cell.font.copy()
            target_cell.border = source_cell.border.copy()
            target_cell.fill = source_cell.fill.copy()
            target_cell.number_format = source_cell.number_format
            target_cell.protection = source_cell.protection.copy()
            target_cell.alignment = source_cell.alignment.copy()
    
    # 행 높이 복사
    ws.row_dimensions[new_row_num].height = ws.row_dimensions[template_row_num].height
    
    # 데이터 입력 (컬럼 A=1, B=2, ... 순서)
    ws.cell(new_row_num, 1).value = emp['이름']
    ws.cell(new_row_num, 2).value = emp['주민등록번호']
    ws.cell(new_row_num, 3).value = emp['귀속월']
    ws.cell(new_row_num, 4).value = emp['근무일수']
    ws.cell(new_row_num, 5).value = emp['근로시간']
    ws.cell(new_row_num, 6).value = emp['기본급']
    ws.cell(new_row_num, 7).value = emp['상여']
    ws.cell(new_row_num, 8).value = emp['식대']
    ws.cell(new_row_num, 9).value = emp['주휴수당']
    ws.cell(new_row_num, 10).value = emp['연장수당']
    ws.cell(new_row_num, 11).value = emp['야간수당']
    ws.cell(new_row_num, 12).value = emp['미지급분']
    ws.cell(new_row_num, 13).value = emp['휴일기본수당']
    ws.cell(new_row_num, 14).value = emp['연차수당']
    ws.cell(new_row_num, 15).value = emp['휴일근무수당']
    ws.cell(new_row_num, 16).value = emp['국민연금']
    ws.cell(new_row_num, 17).value = emp['건강보험']
    ws.cell(new_row_num, 18).value = emp['장기요양보험']
    ws.cell(new_row_num, 19).value = emp['고용보험']
    ws.cell(new_row_num, 20).value = emp['소득세']
    ws.cell(new_row_num, 21).value = emp['지방소득세']
    
    # 금액 컬럼에 쉼표 포맷 적용
    for col_idx in range(6, 22):  # F부터 U까지 (6-21)
        ws.cell(new_row_num, col_idx).number_format = '#,##0'
    
    print(f"  [{idx}/{len(employee_list)}] {emp['이름']} 님 데이터 추가 완료")

# 파일을 새 이름으로 저장
wb.save(output_file)

print("\n" + "=" * 80)
print("✅ 처리 완료!")
print("=" * 80)
print(f"생성된 파일:")
print(f"  1. {account_list_file}")
print(f"  2. {output_file}")
print(f"\n포함된 직원 수: {len(employee_list)}명")
print("\n포함된 직원 목록:")
for emp in employee_list:
    print(f"  - {emp['이름']}")
print("=" * 80)

