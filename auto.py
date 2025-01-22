import pandas as pd
import requests
import time



from bs4 import BeautifulSoup
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (kjhtml, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
url = "https://example.com"  # Define the URL here
try:
    response = requests.get(url, headers=headers)
    response.raise_for_status()  # HTTP 에러 확인
    # 엑셀 파일 읽기
    df = pd.read_excel(r'c:\Users\Administrator\Desktop\automation\연구소_정보통신업.xlsx')
    # 예: '제품명' 열에서 검색 키워드 추출
    keywords = df['기업명'].tolist()
except requests.exceptions.RequestException as e:
    print(f"Error: {e}")

def search_info(keyword):
    url = f"https://www.google.com/search?q={keyword}"
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    # 여기서 필요한 정보를 추출하는 로직을 구현합니다
    # 예: 첫 번째 검색 결과의 제목
    result = soup.find('h3')
    if result:
        return result.text.strip()
    else:
        return "No result found"

# 검색 결과를 저장할 리스트
search_results = []

for keyword in keywords[:100]:  # 처음 100개 키워드만 처리
    result = search_info(keyword)
    search_results.append(result)
    time.sleep(2)  # 2초 대기



#for keyword in keywords:
#    result = search_info(keyword)
#    search_results.append(result)

df['검색결과'] = search_results
df.to_excel('output_file.xlsx', index=False)


