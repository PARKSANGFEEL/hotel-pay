import pandas as pd
import glob
from openpyxl import load_workbook
import calendar



# 파일 경로 - 그리드인_급여대장으로 시작하는 파일 자동 찾기
gridin_files = []
for ext in ("xlsx", "xls", "xlsm", "xlsb"):
    gridin_files += glob.glob(f"그리드인_급여대장*.{ext}")
if not gridin_files:
    print("❌ 오류: '그리드인_급여대장'으로 시작하는 파일을 찾을 수 없습니다.")
    exit(1)

# 급여대장 1행에서 년월 추출
df_preview = pd.read_excel(gridin_files[0], header=None, nrows=2)
pay_title_row = df_preview.iloc[0]
pay_title_str = ''
for cell in pay_title_row:
    if isinstance(cell, str) and '급여대장' in cell:
        pay_title_str = cell
        break

import re
pay_year = ""
pay_month = ""
match = re.search(r"(20[0-9]{2})년\s*([0-9]{1,2})월", pay_title_str)
if match:
    pay_year = match.group(1)
    pay_month = match.group(2).zfill(2)
else:
    # fallback: 기존 값 사용
    pay_year = "2025"
    pay_month = "12"

current_year = pay_year
current_month = pay_month

# 해당 월의 일수 계산
days_in_month = calendar.monthrange(int(current_year), int(current_month))[1]

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
print("\n[디버그] 급여대장 데이터프레임 상위 20개 행 (인덱스 포함):")
for i in range(min(20, len(df_gridin))):
    print(f"[{i}] {list(df_gridin.iloc[i])}")
df_gridin = pd.read_excel(input_file, header=None)
print("\n[디버그] 급여대장 데이터프레임 상위 20개 행 (인덱스 포함):")
for i in range(min(20, len(df_gridin))):
    print(f"[{i}] {list(df_gridin.iloc[i])}")

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


# 3행씩 한 세트로 처리 (5,6,7행이 한 직원)
def to_int(value):
    if pd.isna(value) or value == 0 or value == '':
        return 0
    try:
        if isinstance(value, str):
            value = value.replace(',', '')
        return int(float(value))
    except:
        return 0

for idx in range(2, len(df_gridin), 3):
    if idx + 2 >= len(df_gridin):
        break  # 남은 행이 3개 미만이면 종료
    row1 = df_gridin.iloc[idx]    # 5행: 기본정보/지급
    row2 = df_gridin.iloc[idx+1]  # 6행: 공제
    row3 = df_gridin.iloc[idx+2]  # 7행: 실지급 등

    # 사원명 추출 (row1의 1번 컬럼)
    employee_name = str(row1[1]).replace('(재)', '').strip() if pd.notna(row1[1]) else "확인필요"
    # 사원번호와 이름이 모두 있는 경우만 직원으로 인정
    has_empno = pd.notna(row1[0]) and str(row1[0]).strip().isdigit()
    has_name = pd.notna(row1[1]) and str(row1[1]).strip() != ''
    is_header = False
    if has_name:
        name_str = str(row1[1]).replace(' ', '').replace('\xa0', '').replace('\x00', '')
        if name_str in ['성명', '성명']:
            is_header = True
    if not (has_empno and has_name) or is_header:
        continue
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
    # 디버그: 각 항목의 원본 값과 변환값 출력
    print(f"[디버그] {employee_name} 기본급: {row1[2]}, 변환: {to_int(row1[2])}")
    print(f"[디버그] {employee_name} 식대: {row1[3]}, 변환: {to_int(row1[3])}")
    print(f"[디버그] {employee_name} 주휴수당: {row1[4]}, 변환: {to_int(row1[4])}")
    print(f"[디버그] {employee_name} 연차수당: {row1[13]}, 변환: {to_int(row1[13])}")
    print(f"[디버그] {employee_name} 국민연금: {row1[9]}, 변환: {to_int(row1[9])}")
    print(f"[디버그] {employee_name} 차인지급액 원본: {row3[14]}, 변환값: {to_int(row3[14])}")
    employee_list.append({
        '이름': employee_name,
        '주민등록번호': resident_number,
        '귀속월': f'{current_year}{current_month}',
        '근무일수': work_days,
        '근로시간': work_hours,
        # 지급항목 (row1)
        '기본급': to_int(row1[2]),
        '식대': to_int(row1[3]),
        '주휴수당': to_int(row1[4]),
        '연장수당': to_int(row1[5]),
        '야간수당': to_int(row1[6]),
        '미지급분': to_int(row1[7]),
        '휴일기본수당': to_int(row1[8]),
        '상여': 0,  # 상여는 별도 컬럼 없음
        '연차수당': 0,
        # 공제항목 (row1)
        '국민연금': to_int(row1[9]),
        '건강보험': to_int(row1[10]),
        '고용보험': to_int(row1[11]),
        '장기요양보험': to_int(row1[12]),  # 실제로는 row1[12]가 장기요양보험료
        '소득세': to_int(row1[13]),
        '지방소득세': to_int(row1[14]),
        # 실지급액(차인지급액)
        '차인지급액': to_int(row3[14])
    })

    # 사원명 추출 (row1의 1번 컬럼)
    employee_name = str(row1[1]).replace('(재)', '').strip() if pd.notna(row1[1]) else "확인필요"

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

    # 사원번호와 이름이 모두 있는 경우만 직원으로 인정
    has_empno = pd.notna(row1[0]) and str(row1[0]).strip() != ''
    has_name = pd.notna(row1[1]) and str(row1[1]).strip() != ''
    # row1[1]이 '성명' 또는 '성  명' 등 헤더 문자열이면 직원으로 보지 않음
    is_header = False
    if has_name:
        name_str = str(row1[1]).replace(' ', '').replace('\xa0', '').replace('\x00', '')
        if name_str in ['성명', '성명']:
            is_header = True
    if not (has_empno and has_name) or is_header:
        continue
    # 디버그: 차인지급액 원본 값과 변환값 출력
    print(f"[디버그] {employee_name} 차인지급액 원본: {row3[14]}, 변환값: {to_int(row3[14])}")
    employee_list.append({
        '이름': employee_name,
        '주민등록번호': resident_number,
        '귀속월': f'{current_year}{current_month}',
        '근무일수': work_days,
        '근로시간': work_hours,
        # 지급항목 (row1)
        '기본급': to_int(row1[2]),
        '상여': to_int(row1[6]),
        '식대': to_int(row1[3]),
        '주휴수당': to_int(row1[4]),
        '연장수당': to_int(row1[5]),
        '야간수당': to_int(row1[10]),
        '미지급분': to_int(row1[11]),
        '휴일기본수당': to_int(row1[12]),
        '연차수당': to_int(row1[13]),
        '휴일근무수당': to_int(row1[14]),
        # 공제항목 (row2)
        '국민연금': to_int(row2[5]),
        '건강보험': to_int(row2[6]),
        '고용보험': to_int(row2[7]),
        '장기요양보험': to_int(row2[8]),
        '소득세': to_int(row2[9]),
        '지방소득세': to_int(row2[10]),
        # 실지급액(차인지급액)
        '차인지급액': to_int(row3[14])
    })


# employee_list 중복 제거 (직원명 기준)
unique_employees = []
seen_names = set()
for emp in employee_list:
    if emp['이름'] not in seen_names:
        unique_employees.append(emp)
        seen_names.add(emp['이름'])
employee_list = unique_employees

print(f"  ✓ {len(employee_list)}명의 급여 데이터를 읽었습니다.")

# 2. 그리드인_직업급여이체계좌리스트 파일 생성
print("\n[2/3] 그리드인_직업급여이체계좌리스트 파일 생성 중...")

account_list_file = f"그리드인_급여이체_{current_year}년{current_month}월.xlsx"

# 계좌 리스트 데이터 준비

# 급여이체 계좌리스트: worker.xlsx에 있는 직원만, 차인지급액을 이체금액으로 사용
account_data = []
idx = 1
for emp in employee_list:
    # worker.xlsx에 있는 직원만 포함
    if df_worker is not None:
        worker_info = df_worker[df_worker['이름'] == emp['이름']]
        if worker_info.empty:
            continue  # worker에 없는 직원은 제외
        bank = worker_info.iloc[0]['은행']
        account = worker_info.iloc[0]['계좌번호']
    else:
        continue

    # 차인지급액: 급여대장 3번째 행(실지급액)에서 추출한 값 사용
    # employee_list 생성 시 '차인지급액' 항목을 추가해야 함
    total_amount = emp.get('차인지급액', 0)

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
    idx += 1

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
    ws.cell(new_row_num, 6).value = emp['기본급']      # row1[2]
    ws.cell(new_row_num, 7).value = emp.get('상여', 0) # 상여는 별도 컬럼 없음(0)
    ws.cell(new_row_num, 8).value = emp['식대']        # row1[3]
    ws.cell(new_row_num, 9).value = emp['주휴수당']    # row1[4]
    ws.cell(new_row_num, 10).value = emp['연장수당']   # row1[5]
    ws.cell(new_row_num, 11).value = emp['야간수당']   # row1[6]
    ws.cell(new_row_num, 12).value = emp['미지급분']   # row1[7]
    ws.cell(new_row_num, 13).value = emp['휴일기본수당'] # row1[8]
    ws.cell(new_row_num, 14).value = 0   # 연차수당 없음(0)
    ws.cell(new_row_num, 15).value = 0                 # 휴일근무수당 없음(0)
    ws.cell(new_row_num, 16).value = emp['국민연금']   # row1[9]
    ws.cell(new_row_num, 17).value = emp['건강보험']   # row1[10]
    ws.cell(new_row_num, 18).value = emp['장기요양보험'] # row1[12]
    ws.cell(new_row_num, 19).value = emp['고용보험']   # row1[11]
    ws.cell(new_row_num, 20).value = emp['소득세']     # row1[13]
    ws.cell(new_row_num, 21).value = emp['지방소득세'] # row1[14]
    
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

