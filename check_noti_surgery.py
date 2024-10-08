# 103, 105
import os
import pandas as pd
from datetime import datetime, timedelta

# 기본 폴더 경로
input_folder = r'C:\Users\USER\Desktop\기획\프로젝트\pdf\output_folder'
output_folder = r'C:\Users\USER\Desktop\기획\프로젝트\pdf\check_noti_output_folder'
os.makedirs(output_folder, exist_ok=True)  # 출력 폴더가 없으면 생성

# 조건에 사용할 키워드 목록
condition_1_keywords = ['진찰료', '입원료', '검사료', '조제료']  # 조건 1에 해당하는 키워드
condition_2_keywords_1 = ['술']  # 조건 2-1에 해당하는 키워드
condition_2_keywords_2 = ['수술'] # 조건 2-2에 해당하는 키워드

# 현재 시간을 이용한 파일명에 사용할 타임스탬프 생성
current_time = datetime.now().strftime('%Y%m%d_%H%M%S')

# 최근 3개월 계산
three_months_ago = datetime.now() - timedelta(days=90)

# 폴더 내 파일명에 'prep_detail'이 포함된 파일만 처리
for filename in os.listdir(input_folder):
    if filename.endswith('.xlsx') and 'prep_detail' in filename:  # 'prep_detail' 확인
        file_path = os.path.join(input_folder, filename)

        # 엑셀 파일 읽기
        df = pd.read_excel(file_path)

        # '진료시작일' 열이 날짜 형식인지 확인 후 변환
        df['진료시작일'] = pd.to_datetime(df['진료시작일'], format='%Y-%m-%d', errors='coerce')

        # '진료내역'과 '코드명' 열에서 공백, 줄바꿈, 개행문자 제거
        df['진료내역'] = df['진료내역'].str.replace(r'[\s\n\t]+', '', regex=True)
        df['코드명'] = df['코드명'].str.replace(r'[\s\n\t]+', '', regex=True)

        # 최근 3개월 이내의 데이터만 필터링
        recent_df = df[df['진료시작일'] >= three_months_ago]

        # 조건 1: '진료내역' 열에 특정 키워드를 포함하고, 진료시작일이 최근 3개월 이내인 경우
        condition_1_df = recent_df[recent_df['진료내역'].str.contains('|'.join(condition_1_keywords), na=False)].copy()
        if not condition_1_df.empty:
            condition_1_df = condition_1_df.assign(상태='치료(103)')  # '상태' 열 추가 및 값 설정
            output_filename = f"{filename.split('_')[0]}_103_{current_time}.xlsx"  # 새로운 파일명 생성
            condition_1_df['진료시작일'] = condition_1_df['진료시작일'].dt.strftime('%Y-%m-%d') # '진료시작일'을 'yyyy-mm-dd' 형식으로 변환
            condition_1_df.to_excel(os.path.join(output_folder, output_filename), index=False)  # 파일 저장
            print(f'{output_filename} 파일이 생성되었습니다.')

        # 조건 2: '코드명' 열에 '술' 키워드를 포함하거나 '진료내역' 열에 '수술'이 포함되고, 진료시작일이 최근 3개월 이내인 경우
        condition_2_df = recent_df[
            recent_df['코드명'].str.contains('|'.join(condition_2_keywords_1), na=False) |
            recent_df['진료내역'].str..contains('|'.join(condition_2_keywords_2),na=False)
        ].copy()

        if not condition_2_df.empty:
            condition_2_df = condition_2_df.assign(상태='수술(105)')  # '상태' 열 추가 및 값 설정
            output_filename = f"{filename.split('_')[0]}_105_{current_time}.xlsx"  # 새로운 파일명 생성
            condition_2_df['진료시작일'] = condition_2_df['진료시작일'].dt.strftime('%Y-%m-%d') # '진료시작일'을 'yyyy-mm-dd' 형식으로 변환
            condition_2_df.to_excel(os.path.join(output_folder, output_filename), index=False)  # 파일 저장
            print(f'{output_filename} 파일이 생성되었습니다.')
