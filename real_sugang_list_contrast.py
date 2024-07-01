import pandas as pd

# 엑셀 파일 경로
file1 = 'jeonnam.xlsx'
file2 = 'jeonnam_sugang_list.xls'
output_file = 'example3.xlsx'

# 엑셀 파일 읽기
df1 = pd.read_excel(file1, header=1)
df2 = pd.read_excel(file2)

# 비교할 컬럼
column1_name = '이름'
column1_subject = '과목명'
column2_name = '성명'
column2_subject = '교과목명'

# 두 컬럼을 결합하여 비교할 수 있는 새로운 컬럼 생성
df1['combined'] = df1[column1_name] + ' ' + df1[column1_subject]
df2['combined'] = df2[column2_name] + ' ' + df2[column2_subject]

# 'combined' 컬럼의 값을 리스트로 추출
combined_in_file2 = df2['combined'].tolist()

# 'combined' 컬럼에서 겹치지 않는 값 필터링
filtered_df1 = df1[~df1['combined'].isin(combined_in_file2)]

# 불필요한 'combined' 컬럼 삭제
filtered_df1 = filtered_df1.drop(columns=['combined'])

# 결과를 엑셀 파일로 저장
filtered_df1.to_excel(output_file, index=False)

print(f"겹치지 않는 항목을 {output_file}에 저장했습니다.")
