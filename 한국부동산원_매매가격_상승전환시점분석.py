import pandas as pd

# 1. 데이터 로드 및 필터링
file_path = 'H:/파이썬코드모음/시계열다루기/r1_week.csv'
data = pd.read_csv(file_path)
filtered_data = data[(data['area_code2'] == 11) & (data['gubun'] == 1)].copy()

# 2. 매매상승전환시점 찾기
filtered_data.loc[:, 'm_dt_lag'] = filtered_data.groupby('sigun')['m_dt'].shift(1)
turning_points = filtered_data[(filtered_data['m_dt'] > 0) & (filtered_data['m_dt_lag'] <= 0)]
latest_turning_points = turning_points.groupby('sigun')['ymd'].max().reset_index()
latest_turning_points.columns = ['sigun', '매매상승전환시점']

# 3. 가장최근ymd시점 찾기
most_recent_ymd = filtered_data.groupby('sigun')['ymd'].max().reset_index()
most_recent_ymd.columns = ['sigun', '가장최근ymd시점']

# 4. 매매변동률 계산
latest_turning_points = latest_turning_points.merge(
    filtered_data[['sigun', 'ymd', 'm_js']], 
    left_on=['sigun', '매매상승전환시점'], 
    right_on=['sigun', 'ymd'], 
    how='left'
)
latest_turning_points.drop(columns=['ymd'], inplace=True)
latest_turning_points.rename(columns={'m_js': 'm_js_turning'}, inplace=True)

most_recent_ymd = most_recent_ymd.merge(
    filtered_data[['sigun', 'ymd', 'm_js']], 
    left_on=['sigun', '가장최근ymd시점'], 
    right_on=['sigun', 'ymd'], 
    how='left'
)
most_recent_ymd.drop(columns=['ymd'], inplace=True)
most_recent_ymd.rename(columns={'m_js': 'm_js_recent'}, inplace=True)

result = latest_turning_points.merge(most_recent_ymd, on='sigun')
result['매매변동률'] = (result['m_js_recent'] - result['m_js_turning']) / result['m_js_turning'] * 100

# 5. 주수 계산
result['매매상승전환시점'] = pd.to_datetime(result['매매상승전환시점'])
result['가장최근ymd시점'] = pd.to_datetime(result['가장최근ymd시점'])
result['주수'] = (result['가장최근ymd시점'] - result['매매상승전환시점']).dt.days // 7

# 6. 권역구분
gangnam_districts = ['강서구', '양천구', '구로구', '영등포구', '금천구', '동작구', '관악구', '서초구', '강남구', '송파구', '강동구']
result['권역구분'] = result['sigun'].apply(lambda x: '강남권역' if any(district in x for district in gangnam_districts) else '강북권역')

# 정렬 및 포맷팅
result = result.sort_values(by='매매상승전환시점')
result['매매변동률_포맷'] = result['매매변동률'].round(2).astype(str)
formatted_sigun_names = ' → '.join(result.apply(lambda x: f"{x['sigun']} ({x['매매변동률_포맷']})", axis=1))

# 결과 CSV 파일로 저장
output_path = 'H:/네이버블로그자료/한국부동산원_매매상승시점분석.csv'
result.to_csv(output_path, index=False, encoding='utf-8-sig')

# 결과 출력
print(formatted_sigun_names)

# CSV 파일 다운로드 링크 제공
print(f"CSV 파일 저장 경로: {output_path}")
