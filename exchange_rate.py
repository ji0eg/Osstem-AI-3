# =====================================================
# 환율 수집 프로그램
# 한국은행 ECOS API에서 날짜별 환율을 가져와 엑셀로 저장합니다.
# =====================================================

import requests  # 인터넷에서 데이터를 가져오는 도구
import pandas as pd  # 데이터를 표 형태로 정리하는 도구
from datetime import datetime  # 날짜를 다루는 도구

# -------------------------------------------------------
# [설정] 여기서 원하는 값을 바꾸세요!
# -------------------------------------------------------

# 한국은행 API 키 (https://ecos.bok.or.kr 에서 무료 발급)
API_KEY = "여기에_API_키를_입력하세요"

# 가져올 통화 목록 (통화코드: 화면에 표시될 이름)
CURRENCIES = {
    "USD": "미국 달러",
    "EUR": "유럽 유로",
    "JPY": "일본 엔 (100엔)",
    "CNY": "중국 위안",
    "GBP": "영국 파운드",
}

# 조회 기간 설정
START_DATE = "20240101"  # 시작일 (YYYYMMDD 형식)
END_DATE   = "20241231"  # 종료일 (YYYYMMDD 형식)

# 저장할 엑셀 파일 이름
OUTPUT_FILE = "환율_데이터.xlsx"

# -------------------------------------------------------
# [함수] 한국은행 API에서 특정 통화의 환율을 가져옵니다
# -------------------------------------------------------
def get_exchange_rate(currency_code, start_date, end_date):
    """
    currency_code: 통화 코드 (예: "USD")
    start_date: 시작일 (예: "20240101")
    end_date: 종료일 (예: "20241231")
    반환값: 날짜별 환율이 담긴 딕셔너리(사전)
    """

    # API 주소 만들기
    url = (
        f"https://ecos.bok.or.kr/api/StatisticSearch/"
        f"{API_KEY}/json/kr/1/1000/"  # 최대 1000개 데이터
        f"731Y001/{start_date}/{end_date}/{currency_code}"
    )

    # 인터넷에서 데이터 요청
    response = requests.get(url)

    # 응답이 정상인지 확인
    if response.status_code != 200:
        print(f"⚠️  {currency_code} 데이터를 가져오지 못했습니다. (오류코드: {response.status_code})")
        return {}

    # JSON(데이터 형식) 파싱(읽기)
    data = response.json()

    # 데이터가 없거나 오류인 경우 처리
    if "StatisticSearch" not in data:
        print(f"⚠️  {currency_code}: 데이터가 없습니다. API 키를 확인해주세요.")
        return {}

    # 날짜별 환율을 딕셔너리로 정리
    result = {}
    for item in data["StatisticSearch"]["row"]:
        date = item["TIME"]       # 날짜 (예: "20240103")
        value = item["DATA_VALUE"]  # 환율 값 (예: "1327.5")

        # 날짜 형식을 보기 좋게 변환 (20240103 → 2024-01-03)
        formatted_date = f"{date[:4]}-{date[4:6]}-{date[6:]}"

        # 값이 있을 때만 저장
        if value:
            result[formatted_date] = float(value)

    return result


# -------------------------------------------------------
# [메인] 프로그램 실행 시 이 부분이 동작합니다
# -------------------------------------------------------
def main():
    print("=" * 50)
    print("  환율 데이터 수집 프로그램 시작!")
    print("=" * 50)
    print(f"  조회 기간: {START_DATE} ~ {END_DATE}")
    print(f"  수집 통화: {', '.join(CURRENCIES.keys())}")
    print()

    # 모든 통화의 환율을 담을 빈 표(딕셔너리) 준비
    all_data = {}

    # 각 통화별로 데이터 수집
    for code, name in CURRENCIES.items():
        print(f"  📥 {name} ({code}) 수집 중...")
        rates = get_exchange_rate(code, START_DATE, END_DATE)

        if rates:
            all_data[f"{code} ({name})"] = rates
            print(f"  ✅ {len(rates)}개 데이터 수집 완료!")
        else:
            print(f"  ❌ {code} 데이터 없음")

    # 수집된 데이터가 없으면 종료
    if not all_data:
        print("\n❌ 수집된 데이터가 없습니다. API 키를 확인해주세요.")
        return

    # 데이터를 표(DataFrame) 형태로 변환
    df = pd.DataFrame(all_data)
    df.index.name = "날짜"  # 첫 번째 열 이름 설정
    df = df.sort_index()    # 날짜 순서대로 정렬

    # 엑셀 파일로 저장
    print(f"\n  💾 엑셀 파일 저장 중... ({OUTPUT_FILE})")

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="환율데이터")

        # 엑셀 열 너비 자동 조정 (보기 좋게)
        worksheet = writer.sheets["환율데이터"]
        for col in worksheet.columns:
            max_length = max(len(str(cell.value or "")) for cell in col)
            worksheet.column_dimensions[col[0].column_letter].width = max_length + 4

    print(f"  ✅ 저장 완료! '{OUTPUT_FILE}' 파일을 확인하세요.")
    print()
    print("=" * 50)
    print(f"  총 {len(df)}일 데이터, {len(all_data)}개 통화 저장됨")
    print("=" * 50)


# 프로그램 시작점
if __name__ == "__main__":
    main()
