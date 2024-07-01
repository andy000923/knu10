import pandas as pd
import re

# 엑셀 파일 불러오기
input_file = 'jeonnam_sugang_list.xls'  # 원본 엑셀 파일 경로 (xls 파일을 xlsx로 변환)
output_file = 'jeonnam_sugang_sorted.xlsx'  # 결과를 저장할 엑셀 파일 경로

# 엑셀 파일을 데이터프레임으로 읽어오기 (2행을 헤더로 설정)
df = pd.read_excel(input_file, header=0)

# 특정 컬럼으로 그룹화 (예: '과목명' 컬럼)
group_column = '교과목명'  # 필터링 기준 컬럼

# 유효하지 않은 문자를 대체하는 함수
def sanitize_sheet_name(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

# '과목명' 컬럼이 존재하는지 확인
if group_column in df.columns:
    grouped = df.groupby(group_column)

    # 엑셀 라이터 생성
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        for group_name, group_df in grouped:
            # 그룹 이름으로 시트 이름 설정
            sheet_name = sanitize_sheet_name(str(group_name))[:31]  # 시트 이름은 최대 31자 제한
            # 그룹 데이터프레임을 엑셀 시트에 작성
            group_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"데이터를 필터링하여 각각의 시트로 저장했습니다: {output_file}")
else:
    print(f"컬럼 '{group_column}'이(가) 데이터프레임에 존재하지 않습니다.")
