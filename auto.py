import pandas as pd
from googleapiclient.discovery import build

# --- API 키 및 검색 엔진 ID 설정 ---
API_KEY = "AIzaSyDwvr7YS3dZ7YRDDK5bTCSWtHlAml2A-aU"  # Google Cloud Console에서 발급받은 API 키
SEARCH_ENGINE_ID = "1049e4fe5dfc54f3b" # Google Custom Search Engine 설정에서 얻은 ID

def google_search(search_term, api_key, search_engine_id, num_results=5):
    """구글 검색을 수행하고 결과를 반환하는 함수"""
    service = build("customsearch", "v1", developerKey=api_key)
    response = service.cse().list(
        q=search_term,
        cx=search_engine_id,
        num=num_results  # 가져올 검색 결과 개수 (최대 10개)
    ).execute()
    return response

def process_excel_and_search(excel_file_path, search_column, output_file_path="output.xlsx"):
    """엑셀 파일을 읽고 분석하여 구글 검색 후 결과를 엑셀 파일로 저장하는 함수"""
    df = pd.read_excel(r"c:\Users\Administrator\Desktop\research_automation\research.xlsx")  # 엑셀 파일 읽기

    # --- 엑셀 데이터 분석 및 검색어 추출 로직 (사용자 정의 필요) ---
    # 예시: 특정 열(search_column)의 내용을 검색어로 사용
    df['검색결과'] = ""  # 검색 결과를 저장할 열 추가

    for index, row in df.iterrows():
        search_term = row["기업명"]  # 검색어 추출 (예시: 특정 열 값)

        if pd.notna(search_term): # 검색어가 NaN 값이 아닌 경우에만 검색 실행
            search_results = google_search(search_term, API_KEY, SEARCH_ENGINE_ID)

            # --- 검색 결과 분석 및 정보 추출 로직 (사용자 정의 필요) ---
            # 예시: 첫 번째 검색 결과의 제목과 URL을 추출하여 문자열로 저장
            if 'items' in search_results:
                first_result = search_results['items'][0]
                result_info = f"제목: {first_result['title']}\nURL: {first_result['link']}"
                df.at[index, '검색결과'] = result_info
            else:
                df.at[index, '검색결과'] = "검색 결과 없음"
        else:
            df.at[index, '검색결과'] = "검색어 없음" # 검색어가 없는 경우 처리

    df.to_excel(output_file_path, index=False)  # 결과를 엑셀 파일로 저장
    print(f"결과 저장 완료: {output_file_path}")

# --- 실행 예시 ---
if __name__ == "__main__":
    excel_file = "input.xlsx"  # 입력 엑셀 파일 경로
    search_column_name = "제품명"  # 검색어를 추출할 열 이름 (예시)
    output_excel = "output_with_google_results.xlsx" # 출력 엑셀 파일 경로

    process_excel_and_search(excel_file, search_column_name, output_excel)