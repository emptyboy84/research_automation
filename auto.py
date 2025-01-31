import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError  # HttpError import 추가

# --- API 키 및 검색 엔진 ID 설정 ---
API_KEY = "AIzaSyDwvr7YS3dZ7YRDDK5bTCSWtHlAml2A-aU"  # Google Cloud Console에서 발급받은 API 키 (주의: 실제 사용 시 안전하게 관리해야 함)
SEARCH_ENGINE_ID = "107384648519829330799" # Google Custom Search Engine 설정에서 얻은 ID

def google_search(search_term, api_key, search_engine_id, num_results=5):
    """구글 검색을 수행하고 결과를 반환하는 함수"""
    service = build("customsearch", "v1", developerKey=api_key)
    try: # try-except 블록으로 감싸서 HttpError 처리
        response = service.cse().list(
            q=search_term,
            cx=search_engine_id,
            num=num_results  # 가져올 검색 결과 개수 (최대 10개)
        ).execute()
        return response
    except HttpError as e: # HttpError 예외 처리
        print(f"Google Search API 에러 발생: {e}") # 에러 메시지 출력
        return None # 에러 발생 시 None 반환


def process_excel_and_search(excel_file_path="research.xlsx", search_column="기업명", output_file_path="output.xlsx"):
    """엑셀 파일을 읽고 분석하여 구글 검색 후 결과를 엑셀 파일로 저장하는 함수

    Args:
        excel_file_path (str): 입력 엑셀 파일 경로
        search_column (str): 검색어를 추출할 열 이름
        output_file_path (str, optional): 출력 엑셀 파일 경로. Defaults to "output.xlsx".
    """
    df = pd.read_excel(excel_file_path)  # 엑셀 파일 읽기 (파라미터 사용)
    print(f"엑셀 파일 읽기 완료: {excel_file_path}") # 파일 읽기 확인 로그 추가

    df['검색결과'] = ""  # 검색 결과를 저장할 열 추가
    print(f"'{search_column}' 열을 기준으로 검색 시작...") # 검색 시작 로그 추가

    for index, row in df.iterrows():
        search_term = row[search_column]  # 검색어 추출 (파라미터 사용)

        if pd.notna(search_term): # 검색어가 NaN 값이 아닌 경우에만 검색 실행
            print(f"검색어: {search_term}") # 검색어 로그 추가 (디버깅용)
            search_results = google_search(search_term, API_KEY, SEARCH_ENGINE_ID)

            if search_results: # 검색 결과가 None 이 아닌 경우에만 처리 (에러 발생 시 None 반환하도록 수정됨)
                # --- 검색 결과 분석 및 정보 추출 로직 (사용자 정의 필요) ---
                # 예시: 첫 번째 검색 결과의 제목과 URL을 추출하여 문자열로 저장
                if 'items' in search_results:
                    first_result = search_results['items'][0]
                    result_info = f"제목: {first_result['title']}\nURL: {first_result['link']}"
                    df.at[index, '검색결과'] = result_info
                    print(f"'{search_term}' 검색 완료, 결과 저장: {first_result['title']}") # 개별 검색 완료 로그 추가
                else:
                    df.at[index, '검색결과'] = "검색 결과 없음"
                    print(f"'{search_term}' 검색 결과 없음") # 검색 실패 로그 추가
            else: # google_search 함수에서 None 반환된 경우 (에러 발생)
                df.at[index, '검색결과'] = "검색 API 에러 발생" # 에러 발생 명시
                print(f"'{search_term}' 검색 API 에러 발생") # 에러 로그 추가

        else:
            df.at[index, '검색결과'] = "검색어 없음"
            print(f"{index+2}행 검색어 없음") # 검색어 없음을 로그로 표시 (엑셀 행 번호로 표시)

    df.to_excel(output_file_path, index=False)  # 결과를 엑셀 파일로 저장 (파라미터 사용)
    print(f"결과 저장 완료: {output_file_path}")

# --- 실행 예시 ---
if __name__ == "__main__":
    excel_file = "research.xlsx"  # 입력 엑셀 파일 경로 (input.xlsx -> 실제 파일 경로로 변경)
    search_column_name = "기업명"  # 검색어를 추출할 열 이름 (예시: "기업명" 열)
    output_excel = "output.xlsx" # 출력 엑셀 파일 경로

    process_excel_and_search(excel_file, search_column_name, output_excel) # 수정된 변수 사용

    print("자동화 작업 완료!") # 전체 작업 완료 로그 추가