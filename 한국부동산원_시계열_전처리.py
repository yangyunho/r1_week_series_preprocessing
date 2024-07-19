# download_directory = "C:/Users/Administrator/Downloads"
# root_folder = "H:/파이썬코드모음/시계열다루기"

# 폴더경로 입력받기
download_directory = input("다운로드 디렉터리를 입력하세요: ")
root_folder = input("루트 폴더 경로를 입력하세요: ")
print("")
r1_region_2 = f"{root_folder}/r1_region.json"

print("-----------------------------------------------")
print("R1주간아파트가격시계열통계전처리작업시작")

import pandas as pd
from datetime import datetime
import os
import shutil
import time
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
driver = webdriver.Chrome(service=Service(), options=options)


if download_directory == "default":
    최초다운로드 = os.path.join(os.path.expanduser("~"), "Downloads/")  # 파일들이 들어있는 폴더
else:
    최초다운로드 = f"{download_directory}/"

시계열파일이동할곳 = f"{root_folder}"

###############폴더만들기##############
def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Creating directory. " + directory)

createFolder(시계열파일이동할곳)
###############폴더만들기##############

driver.get(
    "https://www.reb.or.kr/r-one/na/ntt/selectNttList.do?mi=9509&bbsId=1106"
)

# table element 접근. 찾는 속성은 적절하게 고려한다.
table = driver.find_element(By.XPATH, '//*[@id="contView"]/div[4]/table')

# tbody
tbody = table.find_element(By.TAG_NAME, "tbody")

td3_text = []
for tr in tbody.find_elements(By.TAG_NAME, "tr"):
    for td in [tr.find_elements(By.TAG_NAME, "td")[2]]:
        td3_text.append(td.get_attribute("innerText"))


w = td3_text.index("주간아파트가격동향조사 시계열통계표") + 1
m = td3_text.index("전국주택가격동향조사(월간) 시계열 통계표") + 1

driver.find_element(
    By.XPATH, f'//*[@id="contView"]/div[4]/table/tbody/tr[{w}]/td[3]'
).click()
time.sleep(1)

# driver.execute_script("window.scrollTo(0, 700)")

driver.switch_to.frame("raonkuploader_frame_kupload")

driver.find_element(By.XPATH, '//*[@id="file_list"]/li[2]/ul/li[2]').click()
time.sleep(1)

driver.find_element(By.XPATH, '//*[@id="button_download"]').click()
time.sleep(5)

file = 최초다운로드 + "/r1_week.xlsx"

if os.path.exists(file):
    os.remove(file)

time.sleep(5)

# 가장최근파일가져오기
fileEx = r".xlsx"
file_name_and_time_lst = []
# 해당 경로에 있는 파일들의 생성시간을 함께 리스트로 넣어줌.
for f_name in os.listdir(f"{최초다운로드}"):
    if f_name.endswith(fileEx):
        written_time = os.path.getctime(f"{최초다운로드}{f_name}")
        file_name_and_time_lst.append((f_name, written_time))
# 생성시간 역순으로 정렬하고,
sorted_file_lst = sorted(
    file_name_and_time_lst, key=lambda x: x[1], reverse=True
)
# 가장 앞에 있는 것을 넣어준다.
# print(file_name_and_time_lst)
recent_file = sorted_file_lst[0]
recent_file_name = recent_file[0]

# print(recent_file)
print(recent_file_name)

time.sleep(5)

# 파일이름변경
os.rename(f"{최초다운로드}/{recent_file_name}", f"{최초다운로드}/r1_week.xlsx")

###############폴더내파일삭제##############
file = 시계열파일이동할곳 + "/r1_week.xlsx"

if os.path.exists(file):
    os.remove(file)
###############폴더내파일삭제##############


# 예시
shutil.move(f"{최초다운로드}/r1_week.xlsx", 시계열파일이동할곳)
############################한국부동산원주간시계열################################
driver.close()

######################################################################

# 파일이름 나열하기
import os
from datetime import datetime
import pandas as pd
from functools import reduce

###################r1_week가져오기#################################
print("r1_week가져오기")
file_path = f"{root_folder}/r1_week.xlsx"

def r1_week(name):
    df = pd.read_excel(file_path, sheet_name=name)  # for구문으로 csv파일들을 읽어 들인다

    # index가 0인 행만 삭제
    # df = df.drop(index=0, axis=0)

    # index가 0,1,2인 행 삭제
    df = df.drop(index=[0, 1, 2, 3], axis=0)

    #     df
    header = df.iloc[0]
    # print(header)

    header = []
    for i in range(0, 236):
        header.append(i)
    # print(header)
    df = df[df["Unnamed: 7"].notna()]

    df.columns = header
    df

    value_vars = []
    for i in range(0, 236):
        value_vars.append(i)

    df = df.melt(id_vars=0, value_vars=value_vars)

    df = df.astype({"value": "float32"})

    df.rename(
        columns={0: "ymd", "variable": "index", "value": name}, inplace=True
    )

    return df


data_list = [
    r1_week("매매변동률"),
    r1_week("전세변동률"),
    r1_week("매매지수"),
    r1_week("전세지수"),
]

r1_week = reduce(
    lambda x, y: pd.merge(x, y, on=["ymd", "index"], how="outer"), data_list
)

# ymd오름차순!!!!!
r1_week = r1_week.sort_values(by=["ymd", "index"])

r1_week
###################r1_week가져오기#################################

###################지역명_API가져오기#################################
print("지역명_API테이블 가져오기")

# DataFrame 생성 및 출력
region_api = pd.read_json(r1_region_2)
region_api

region_api.columns
columns = list(region_api.columns)
# print(columns)

# region_api.info()

region_api.astype("str")

region_api["area_code"] = region_api["area_code"].fillna(0)

region_api = region_api.astype({"area_code": "str"})

region_api["area_code2"] = region_api["area_code2"].fillna(0)

region_api = region_api.astype({"area_code2": "str"})

region_api["area_code4"] = region_api["area_code4"].fillna(0)

region_api = region_api.astype({"area_code4": "str"})

region_api["area_code5"] = region_api["area_code5"].fillna(0)

region_api = region_api.astype({"area_code5": "str"})

# 176개지역 필터링하기
# region_api = region_api[(region_api['gubun']== 1)|(region_api['gubun']== 13)]

# region_api.info()

region_api
# ###################지역명_API가져오기#################################

###################두테이블조인하기##################################
# print("기존 시세테이블(mama4)과 region_api테이블 병합")
print("시계열&지역명 VLOOKUP작업1")

r1_week = pd.merge(
    left=r1_week,
    right=region_api,
    how="inner",
    left_on=["index"],
    right_on=["index"],
    sort=False,
)

r1_week = r1_week.astype({"gubun": "int64"})

r1_week
###################두테이블조인하기##################################

###################시군구매매증감전세증감순위매기기##################################
series = r1_week.copy()[
    (r1_week.copy()["gubun"] == 3) | (r1_week.copy()["gubun"] == 13)
]
series

series["매매증감순위1"] = (
    series.copy().groupby(["ymd"])["매매변동률"].rank(method="min", ascending=False)
)
series["전세증감순위1"] = (
    series.copy().groupby(["ymd"])["전세변동률"].rank(method="min", ascending=False)
)
series = series[["ymd", "area_code2", "매매증감순위1", "전세증감순위1"]]
###################시군구매매증감전세증감순위매기기##################################

###################두테이블조인하기##################################
# print("기존 시세테이블(mama4)과 region_api테이블 병합")
print("시계열&지역명 VLOOKUP작업2")

r1_week1 = pd.merge(
    left=r1_week,
    right=series,
    how="left",
    left_on=["ymd", "area_code2"],
    right_on=["ymd", "area_code2"],
    sort=False,
)

r1_week1
###################두테이블조인하기##################################

###################시도매매증감전세증감순위매기기##################################
series1 = r1_week.copy()[
    (r1_week.copy()["gubun"] == 1) | (r1_week.copy()["gubun"] == 13)
]
series1

series1["매매증감순위2"] = (
    series1.copy().groupby(["ymd"])["매매변동률"].rank(method="min", ascending=False)
)
series1["전세증감순위2"] = (
    series1.copy().groupby(["ymd"])["전세변동률"].rank(method="min", ascending=False)
)
series1 = series1[["ymd", "index", "매매증감순위2", "전세증감순위2"]]
###################시도매매증감전세증감순위매기기##################################

###################두테이블조인하기##################################
# print("기존 시세테이블(mama4)과 region_api테이블 병합")
print("시계열&지역명 VLOOKUP작업3")

r1_week2 = pd.merge(
    left=r1_week1,
    right=series1,
    how="left",
    left_on=["ymd", "index"],
    right_on=["ymd", "index"],
    sort=False,
)

r1_week2
r1_week2.rename(
    columns={
        "매매변동률": "m_dt",
        "전세변동률": "j_dt",
        "매매지수": "m_js",
        "전세지수": "j_js",
        "매매증감순위1": "m_dt_rank1",
        "전세증감순위1": "j_dt_rank1",
        "매매증감순위2": "m_dt_rank2",
        "전세증감순위2": "j_dt_rank2",
    },
    inplace=True,
)
r1_week2 = r1_week2.sort_index(axis=1)
###################두테이블조인하기##################################
r1_week2.to_csv(
    f"{root_folder}/r1_week.csv", encoding="utf-8-sig", index=False
)
###################두테이블조인하기##################################
print("R1주간아파트가격시계열통계전처리작업종료")
print("-----------------------------------------------")