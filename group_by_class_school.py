import pandas as pd

def extract_data_from_excel(file_path):
    # 엑셀 파일 읽기 (2번째 행을 컬럼으로 설정)
    df = pd.read_excel(file_path, header=1)

    return df

def find_grade_errors(df):
    # 과목명과 소속학교별로 그룹화하여 최종 성적을 기준으로 정렬
    grouped = df.groupby(['과목명', '소속학교'])

    for (subject, school), group in grouped:
        if school == '전남대학교':
            # 최종 성적 기준으로 정렬
            sorted_group = group.sort_values(by='최종성적', ascending=False).reset_index(drop=True)

            # 상위 50%의 인덱스 계산
            cutoff_index = len(sorted_group) - (len(sorted_group) // 2)

            # 상위 50%에 해당하지 않는 학생들
            lower_half = sorted_group.iloc[cutoff_index:]

            # A 또는 A+인 학생들 중 상위 50%가 아닌 학생들 찾기
            grade_errors = lower_half[(lower_half['등급'] == 'A') | (lower_half['등급'] == 'A+')]

            # 해당 학생들의 이름과 소속학교 출력
            for index, row in grade_errors.iterrows():
                print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 A이상 불가")

            # 상위 80%의 인덱스 계산
            cutoff_index_80 = len(sorted_group) - int(len(sorted_group) * 0.2)

            # 상위 80%에 해당하지 않는 학생들
            lower_80_percent = sorted_group.iloc[cutoff_index_80:]

            # B 또는 B+인 학생들 중 상위 80%에 해당하지 않는 학생들 찾기
            grade_errors_b = lower_80_percent[(lower_80_percent['등급'] == 'B') | (lower_80_percent['등급'] == 'B+')]

            # 해당 학생들의 이름과 소속학교 출력
            for index, row in grade_errors_b.iterrows():
                print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 B이상 불가")

        if school == '제주대학교' or school == '전북대학교':
            # 최종 성적 기준으로 정렬
            sorted_group = group.sort_values(by='최종성적', ascending=False).reset_index(drop=True)

            # 상위 50%의 인덱스 계산
            cutoff_index = len(sorted_group) - int(len(sorted_group) * 0.5)

            # 상위 50%에 해당하지 않는 학생들
            lower_half = sorted_group.iloc[cutoff_index:]

            # A 또는 A+인 학생들 중 상위 50%가 아닌 학생들 찾기
            grade_errors = lower_half[(lower_half['등급'] == 'A') | (lower_half['등급'] == 'A+')]

            # 해당 학생들의 이름과 소속학교 출력
            for index, row in grade_errors.iterrows():
                print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 A이상 불가")

        if school == '충북대학교':
            # 처리할 교과목 리스트
            subjects_to_check = [
                '강원문화사', '여행과문화', '수목보호학', '심리학의이해', '평화사상과통일의길',
                '분쟁의 세계지리', '무역실무', '창업의이해', '심리학 START'
            ]

            if subject in subjects_to_check:
                # 최종 성적 기준으로 정렬
                sorted_group = group.sort_values(by='최종성적', ascending=False).reset_index(drop=True)

                # 상위 33%의 인덱스 계산
                cutoff_index_33 = int(len(sorted_group) * 0.33)

                # 상위 33%에 해당하지 않는 학생들
                lower_two_thirds = sorted_group.iloc[cutoff_index_33:]

                # A 또는 A+인 학생들 중 상위 33%가 아닌 학생들 찾기
                grade_errors_a = lower_two_thirds[(lower_two_thirds['등급'] == 'A') | (lower_two_thirds['등급'] == 'A+')]

                # 해당 학생들의 이름과 소속학교 출력
                for index, row in grade_errors_a.iterrows():
                    print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 A이상 불가")

                # 상위 70%의 인덱스 계산
                cutoff_index_70 = int(len(sorted_group) * 0.7)

                # 상위 70%에 해당하지 않는 학생들
                lower_30_percent = sorted_group.iloc[cutoff_index_70:]

                # B 또는 B+인 학생들 중 상위 70%에 해당하지 않는 학생들 찾기
                grade_errors_b = lower_30_percent[(lower_30_percent['등급'] == 'B') | (lower_30_percent['등급'] == 'B+')]

                # 해당 학생들의 이름과 소속학교 출력
                for index, row in grade_errors_b.iterrows():
                    print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 B이상 불가")

            else:
                # 상위 40%의 인덱스 계산
                cutoff_index_40 = int(len(sorted_group) * 0.4)

                # 상위 40%에 해당하지 않는 학생들
                lower_sixty_percent = sorted_group.iloc[cutoff_index_40:]

                # A 또는 A+인 학생들 중 상위 60%가 아닌 학생들 찾기
                grade_errors_a = lower_sixty_percent[
                    (lower_sixty_percent['등급'] == 'A') | (lower_sixty_percent['등급'] == 'A+')]

                # 해당 학생들의 이름과 소속학교 출력
                for index, row in grade_errors_a.iterrows():
                    print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 A이상 불가")

        if school == '부산대학교':
            # 처리할 교과목 리스트
            subjects_to_check = [
                '강원문화사', '소비문화(부제: 4차 산업혁명 속 일상으로의 소비)',
                '과학과 문화', '수목보호학', '평화사상과 통일의 길', '디지털 시대의 영상과 문화',
                '에듀테크의 이해', '서양고대철학', '중국고대철학', '한옥개론',
                '화훼산업론', '혁신을위한모순해결', '4차 산업혁명과 수학', '트랜스미디어 스토리텔링',
                '한국지형여행'
            ]
            if subject in subjects_to_check:
                # 최종 성적 기준으로 정렬
                sorted_group = group.sort_values(by='최종성적', ascending=False).reset_index(drop=True)

                # 상위 33%의 인덱스 계산
                cutoff_index_33 = len(sorted_group) - int(len(sorted_group) * 0.6)

                # 상위 33%에 해당하지 않는 학생들
                lower_two_thirds = sorted_group.iloc[cutoff_index_33:]

                # A 또는 A+인 학생들 중 상위 33%가 아닌 학생들 찾기
                grade_errors_a = lower_two_thirds[(lower_two_thirds['등급'] == 'A') | (lower_two_thirds['등급'] == 'A+')]

                # 해당 학생들의 이름과 소속학교 출력
                for index, row in grade_errors_a.iterrows():
                    print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 A이상 불가")

                # 상위 70%의 인덱스 계산
                cutoff_index_70 = len(sorted_group) - int(len(sorted_group) * 0.3)

                # 상위 70%에 해당하지 않는 학생들
                lower_30_percent = sorted_group.iloc[cutoff_index_70:]

                # B 또는 B+인 학생들 중 상위 70%에 해당하지 않는 학생들 찾기
                grade_errors_b = lower_30_percent[(lower_30_percent['등급'] == 'B') | (lower_30_percent['등급'] == 'B+')]

                # 해당 학생들의 이름과 소속학교 출력
                for index, row in grade_errors_b.iterrows():
                    print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 B이상 불가")
            else:
                # 상위 40%의 인덱스 계산
                cutoff_index_40 = len(sorted_group) - int(len(sorted_group) * 0.5)

                # 상위 40%에 해당하지 않는 학생들
                lower_sixty_percent = sorted_group.iloc[cutoff_index_40:]

                # A 또는 A+인 학생들 중 상위 60%가 아닌 학생들 찾기
                grade_errors_a = lower_sixty_percent[(lower_sixty_percent['등급'] == 'A') | (lower_sixty_percent['등급'] == 'A+')]

                # 해당 학생들의 이름과 소속학교 출력
                for index, row in grade_errors_a.iterrows():
                    print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 A이상 불가")


        if school == '경상국립대학교':
            # 처리할 교과목 리스트
            subjects_to_check = [
                '여행과 문화', '고사와 성어의 탐구', '산업경영의이해', '평화사상과 통일의 길', '디지털 시대의 영상과 문화',
                '빅데이터로 읽는 한국근대사 1', '리더십', '창업의이해', '심리학 START',
                '불교가 묻고 내가 답한다', '로마 문화와 함께 배우는 교양 라틴어', '우주로의 여행', '한국지형여행'
            ]
            if subject in subjects_to_check:
                # 최종 성적 기준으로 정렬
                sorted_group = group.sort_values(by='최종성적', ascending=False).reset_index(drop=True)

                # 상위 33%의 인덱스 계산
                cutoff_index_33 = len(sorted_group) - int(len(sorted_group) * 0.7)

                # 상위 30%에 해당하지 않는 학생들
                lower_two_thirds = sorted_group.iloc[cutoff_index_33:]

                # A 또는 A+인 학생들 중 상위 33%가 아닌 학생들 찾기
                grade_errors_a = lower_two_thirds[(lower_two_thirds['등급'] == 'A') | (lower_two_thirds['등급'] == 'A+')]

                # 해당 학생들의 이름과 소속학교 출력
                for index, row in grade_errors_a.iterrows():
                    print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 A이상 불가")

                # 상위 70%의 인덱스 계산
                cutoff_index_70 = len(sorted_group) - int(len(sorted_group) * 0.3)

                # 상위 70%에 해당하지 않는 학생들
                lower_30_percent = sorted_group.iloc[cutoff_index_70:]

                # B 또는 B+인 학생들 중 상위 70%에 해당하지 않는 학생들 찾기
                grade_errors_b = lower_30_percent[(lower_30_percent['등급'] == 'B') | (lower_30_percent['등급'] == 'B+')]

                # 해당 학생들의 이름과 소속학교 출력
                for index, row in grade_errors_b.iterrows():
                    print(f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류 B이상 불가")
            else:
                # A 또는 A+ 등급이지만 최종성적이 95 미만인 경우
                grade_errors_ap = group[(group['등급'].isin(['A+'])) & (group['최종성적'] < 95)]

                # 해당 학생들의 정보 출력
                for index, row in grade_errors_ap.iterrows():
                    print(
                        f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (95 이상 아님)")

                # B 또는 B+ 등급이지만 최종성적이 90 미만인 경우
                grade_errors_a = group[(group['등급'].isin(['A'])) & (group['최종성적'] < 90)]

                # 해당 학생들의 정보 출력
                for index, row in grade_errors_a.iterrows():
                    print(
                        f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (90 이상 아님)")
                grade_errors_bp = group[(group['등급'].isin(['B+'])) & (group['최종성적'] < 85)]

                # 해당 학생들의 정보 출력
                for index, row in grade_errors_bp.iterrows():
                    print(
                        f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (85 이상 아님)")

                # B 또는 B+ 등급이지만 최종성적이 90 미만인 경우
                grade_errors_b = group[(group['등급'].isin(['B'])) & (group['최종성적'] < 80)]

                # 해당 학생들의 정보 출력
                for index, row in grade_errors_b.iterrows():
                    print(
                        f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (80 이상 아님)")

                grade_errors_cp = group[(group['등급'].isin(['C+'])) & (group['최종성적'] < 75)]

                # C
                for index, row in grade_errors_cp.iterrows():
                    print(
                        f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (75 이상 아님)")

                # B 또는 B+ 등급이지만 최종성적이 90 미만인 경우
                grade_errors_c = group[(group['등급'].isin(['C'])) & (group['최종성적'] < 70)]

                # 해당 학생들의 정보 출력
                for index, row in grade_errors_c.iterrows():
                    print(
                        f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (70 이상 아님)")
                grade_errors_dp = group[(group['등급'].isin(['D+'])) & (group['최종성적'] < 65)]

                # 해당 학생들의 정보 출력
                for index, row in grade_errors_dp.iterrows():
                    print(
                        f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (65 이상 아님)")

                # B 또는 B+ 등급이지만 최종성적이 90 미만인 경우
                grade_errors_d = group[(group['등급'].isin(['D'])) & (group['최종성적'] < 60)]

                # 해당 학생들의 정보 출력
                for index, row in grade_errors_d.iterrows():
                    print(
                        f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (60 이상 아님)")


        if school == '강원대학교':
            sorted_group = group.sort_values(by='최종성적', ascending=False).reset_index(drop=True)

            # A 또는 A+ 등급이지만 최종성적이 95 미만인 경우
            grade_errors_ap = group[(group['등급'].isin(['A+'])) & (group['최종성적'] < 95)]

            # 해당 학생들의 정보 출력
            for index, row in grade_errors_ap.iterrows():
                print(
                    f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (95 이상 아님)")

            # B 또는 B+ 등급이지만 최종성적이 90 미만인 경우
            grade_errors_a = group[(group['등급'].isin(['A'])) & (group['최종성적'] < 90)]

            # 해당 학생들의 정보 출력
            for index, row in grade_errors_a.iterrows():
                print(
                    f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (90 이상 아님)")
            grade_errors_bp = group[(group['등급'].isin(['B+'])) & (group['최종성적'] < 85)]

            # 해당 학생들의 정보 출력
            for index, row in grade_errors_bp.iterrows():
                print(
                    f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (85 이상 아님)")

            # B 또는 B+ 등급이지만 최종성적이 90 미만인 경우
            grade_errors_b = group[(group['등급'].isin(['B'])) & (group['최종성적'] < 80)]

            # 해당 학생들의 정보 출력
            for index, row in grade_errors_b.iterrows():
                print(
                    f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (80 이상 아님)")

            grade_errors_cp = group[(group['등급'].isin(['C+'])) & (group['최종성적'] < 75)]

            # C
            for index, row in grade_errors_cp.iterrows():
                print(
                    f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (75 이상 아님)")

            # B 또는 B+ 등급이지만 최종성적이 90 미만인 경우
            grade_errors_c = group[(group['등급'].isin(['C'])) & (group['최종성적'] < 70)]

            # 해당 학생들의 정보 출력
            for index, row in grade_errors_c.iterrows():
                print(
                    f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (70 이상 아님)")
            grade_errors_dp = group[(group['등급'].isin(['D+'])) & (group['최종성적'] < 65)]

            # 해당 학생들의 정보 출력
            for index, row in grade_errors_dp.iterrows():
                print(
                    f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (65 이상 아님)")

            # B 또는 B+ 등급이지만 최종성적이 90 미만인 경우
            grade_errors_d = group[(group['등급'].isin(['D'])) & (group['최종성적'] < 60)]

            # 해당 학생들의 정보 출력
            for index, row in grade_errors_d.iterrows():
                print(
                    f"과목명: {row['과목명']}, 이름: {row['이름']}, 소속학교: {row['소속학교']}, 성적오류: 최종성적 {row['최종성적']}, 등급 {row['등급']} (60 이상 아님)")



#
# def jaesugang(file_path):
#     # 엑셀 파일 읽기 (2번째 행을 컬럼으로 설정)
#     df = pd.read_excel(file_path, header=1)
#
#     # 조건에 맞는 행 필터링
#     filtered_df = df[(df['재수강'] == 'Y') & (df['등급'] == 'A+')]
#
#     # 필요한 열만 선택하여 출력
#     result = filtered_df[['과목명', '이름', '소속학교']]
#
#     print(result)
#     return result
#


# 엑셀 파일 경로
file_path = 'new1.xlsx'

# 데이터 추출
df = extract_data_from_excel(file_path)


# 성적 오류 찾기 및 출력
find_grade_errors(df)
#jaesugang(file_path)

