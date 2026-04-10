# -*- coding: utf-8 -*-
# =====================================================
# SAP Y_OKD_27000037 데이터 조회 및 다운로드
# - 당월 기준으로 자동 조회
# - ALV 결과를 클립보드로 추출 → 엑셀 파일 저장
# =====================================================

import win32com.client          # SAP GUI를 파이썬으로 조작하는 도구
import pyperclip                # 클립보드 복사/붙여넣기 도구
import pandas as pd             # 표 데이터를 엑셀로 저장하는 도구
import time                     # 대기 시간 도구
import os                       # 폴더/파일 경로 도구
from datetime import date       # 오늘 날짜 도구
from dotenv import load_dotenv  # .env 파일에서 설정 읽는 도구

# .env 파일 로드
load_dotenv()

# -------------------------------------------------------
# [설정] 여기서 원하는 값을 바꾸세요!
# -------------------------------------------------------

# 회계연도: 오늘 날짜 기준 자동 설정
TODAY       = date.today()
FISCAL_YEAR = str(TODAY.year)   # 회계연도 (예: "2026")

# 전기종료기간: 항상 12 고정 (전년도 12월)
PRIOR_PERIOD = "12"

# G/L 계정 범위 (비워두면 전체 계정 조회)
ACCOUNT_FROM = ""   # 시작 계정코드 (예: "4100000")
ACCOUNT_TO   = ""   # 종료 계정코드 (예: "4999999")

# 결과 저장 폴더 (파일명은 실행 시 입력받은 종료기간으로 결정)
OUTPUT_DIR = "data/output"

# ALV 클립보드 내보내기 후 대기시간(초) — 컴퓨터가 느리면 늘리세요
CLIPBOARD_WAIT = 3.0

# -------------------------------------------------------
# .env에서 T코드 읽기
# -------------------------------------------------------
SAP_TCODE = os.getenv("SAP_TCODE", "Y_OKD_27000037")


# -------------------------------------------------------
# SAP 연결
# -------------------------------------------------------
def connect_sap():
    """SAP GUI에 연결하고 세션(창) 반환"""
    sap_gui = win32com.client.GetObject("SAPGUI")
    app     = sap_gui.GetScriptingEngine
    conn    = app.Children(0)
    session = conn.Children(0)
    print(f"  [연결 성공] 시스템: {session.Info.SystemName}")
    return session


# -------------------------------------------------------
# T코드 이동
# -------------------------------------------------------
def navigate_to_tcode(session):
    """T코드 화면으로 이동"""
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{SAP_TCODE}"
    session.findById("wnd[0]").sendVKey(0)  # Enter
    time.sleep(1.5)


# -------------------------------------------------------
# 조회 조건 입력
# -------------------------------------------------------
def input_conditions(session, period):
    """
    조회 조건 입력
    - 회계연도: 자동 (올해)
    - 종료기간: 실행 시 입력받은 값
    - 전기종료기간: 12 고정
    """
    # 회계연도 입력 (예: "2026")
    session.findById("wnd[0]/usr/ctxtPAR_01").text = FISCAL_YEAR

    # 종료기간 입력 (실행 시 입력받은 월)
    session.findById("wnd[0]/usr/ctxtPAR_03").text = period

    # 전기종료기간 입력 — 항상 12 고정
    session.findById("wnd[0]/usr/ctxtPAR_02").text = PRIOR_PERIOD

    # G/L 계정 범위 (설정된 경우에만 입력)
    if ACCOUNT_FROM:
        session.findById("wnd[0]/usr/ctxtSD_SAKNR-LOW").text  = ACCOUNT_FROM
    if ACCOUNT_TO:
        session.findById("wnd[0]/usr/ctxtSD_SAKNR-HIGH").text = ACCOUNT_TO

    time.sleep(0.3)
    print(f"  [조건 입력] 회계연도={FISCAL_YEAR}, 종료기간={period}월, 전기종료기간={PRIOR_PERIOD}")


# -------------------------------------------------------
# F8 실행 (조회)
# -------------------------------------------------------
def execute_report(session):
    """F8 키로 조회 실행"""
    session.findById("wnd[0]").sendVKey(8)  # 8 = F8
    time.sleep(3.0)  # 결과 화면 로딩 대기
    print("  [조회 완료] 결과 화면 로드됨")


# -------------------------------------------------------
# 전체 데이터 한 번에 내보내기
# 시스템(Y) → 리스트(I) → 저장(A) → 로컬 파일(I)
# -------------------------------------------------------
def download_all_at_once(session):
    """
    SAP 메뉴 '시스템 → 리스트 → 저장 → 로컬 파일'로
    전체 데이터를 한 번에 파일로 저장합니다.
    형식: 스프레드시트(Excel) → .xls 파일로 저장
    """
    os.makedirs("data/output", exist_ok=True)
    out_dir  = os.path.abspath("data/output")
    filename = "sap_export.xls"   # 스프레드시트 형식은 .xls로 저장됨

    # ① 시스템 → 리스트 → 저장 → 로컬 파일 메뉴 클릭
    # 메뉴 인덱스: 시스템=menu[6], 리스트=menu[5], 저장=menu[2], 로컬파일=menu[2]
    print("  [메뉴] 시스템 → 리스트 → 저장 → 로컬 파일 클릭...")
    session.findById(
        "wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]"
    ).select()
    time.sleep(1.0)

    # ② 형식 선택 팝업: [1,0] 스프레드시트 선택
    #    스프레드시트 = Excel 호환 형식(.xls)으로 저장됨
    try:
        session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150"
            "/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        ).select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(0.8)
        print("  [형식 선택] 스프레드시트 완료")
    except Exception:
        print("  [형식 선택] 팝업 없음 → 건너뜀")

    # ③ 파일명 입력 팝업
    try:
        # 저장 경로 설정
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = out_dir + "\\"
    except Exception:
        pass
    try:
        # 파일명 설정
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
        print(f"  [파일명] {filename}")
    except Exception as e:
        raise RuntimeError(f"파일명 입력 실패: {e}")

    session.findById("wnd[1]/tbar[0]/btn[0]").press()  # 저장(생성) 클릭
    time.sleep(3.0)  # 스프레드시트 생성은 텍스트보다 시간이 더 걸릴 수 있음

    # 덮어쓰기 확인 팝업이 뜰 경우 처리 (파일이 이미 있을 때)
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1.5)
    except Exception:
        pass

    # ④ 파일 탐색 (SAP가 다른 경로에 저장했을 경우 대비)
    search_dirs = [out_dir, os.path.expanduser("~\\Documents"),
                   os.path.expanduser("~"), os.getcwd()]
    temp_path = None
    for d in search_dirs:
        candidate = os.path.join(d, filename)
        if os.path.exists(candidate):
            temp_path = candidate
            break

    if not temp_path:
        raise RuntimeError(
            f"저장 파일을 찾을 수 없습니다. 탐색 경로: {search_dirs}"
        )
    # ── 디버그: 파일 크기 확인 ──────────────────────────────
    file_size = os.path.getsize(temp_path)
    print(f"  [디버그] 파일 경로  : {temp_path}")
    print(f"  [디버그] 파일 크기  : {file_size:,} bytes")

    # ── 디버그: 파일 첫 200바이트 원본 출력 (인코딩 확인용) ──
    with open(temp_path, "rb") as f:
        raw = f.read(200)
    print(f"  [디버그] 첫 200바이트(hex): {raw.hex()}")
    print(f"  [디버그] 첫 200바이트(repr): {repr(raw)}")

    # ── 디버그: 텍스트로 열어서 첫 5줄 출력 ─────────────────
    for enc in ["utf-16", "utf-16-le", "cp949", "utf-8"]:
        try:
            with open(temp_path, "r", encoding=enc) as f:
                preview = [f.readline() for _ in range(5)]
            print(f"  [디버그] 인코딩 {enc} 성공 → 첫 5줄:")
            for i, line in enumerate(preview):
                print(f"    {i+1}: {repr(line)}")
            break
        except Exception as ex:
            print(f"  [디버그] 인코딩 {enc} 실패: {ex}")

    # ⑤ 파일 읽기
    #    SAP '스프레드시트' 형식 = UTF-16 LE 인코딩 탭 구분 텍스트 파일
    #    header=None → 제목/메타데이터 행을 컬럼명으로 처리하지 않고 데이터로 보존
    try:
        df = pd.read_csv(
            temp_path,
            sep="\t",           # 탭으로 열 구분
            encoding="utf-16",  # SAP 스프레드시트 기본 인코딩
            dtype=str,
            header=None,        # SAP 원본 형태 그대로 유지
        )
    except Exception as e:
        raise RuntimeError(f"파일 읽기 실패: {e}")

    # ── 디버그: 파싱 결과 확인 ───────────────────────────────
    print(f"  [디버그] 파싱 후 shape: {df.shape}  (행 × 열)")
    print(f"  [디버그] 컬럼 목록: {list(df.columns)}")
    print(f"  [디버그] 행 0~15 (헤더 포함):")
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 300)
    pd.set_option("display.max_colwidth", 30)
    print(df.head(16).to_string())

    # ⑥ 임시 파일 삭제
    try:
        os.remove(temp_path)
    except Exception:
        pass

    if df.empty:
        print("  [주의] 내보내기 파일이 비어있습니다.")
        return pd.DataFrame()

    # 모든 값을 문자열로 변환 (NaN → 빈 문자열)
    df = df.fillna("").astype(str)

    print(f"  [내보내기 완료] {len(df):,}행 × {len(df.columns)}열")
    return df


# -------------------------------------------------------
# 컬럼 정리 (A열 삭제 + 빈 컬럼 삭제)
# -------------------------------------------------------
def clean_columns(df):
    """
    SAP 스프레드시트 내보내기 시 A열(첫 번째 열)은 항상 비어있으므로 삭제하고,
    값이 하나도 없는 빈 컬럼도 삭제합니다.
    """
    before = len(df.columns)

    # 값이 전혀 없는 컬럼 삭제 (공백·NaN 포함)
    df = df.replace("", pd.NA)
    df = df.dropna(axis=1, how="all")
    df = df.fillna("")
    removed = before - len(df.columns)
    if removed > 0:
        print(f"  [열 정리] 빈 컬럼 {removed}개 삭제")

    print(f"  [열 정리 완료] 남은 컬럼 수: {len(df.columns)}개")
    return df


# -------------------------------------------------------
# 엑셀 파일로 저장 (SAP 원본 형태 유지)
# -------------------------------------------------------
def save_to_excel(df, output_path, sheet_suffix=""):
    """
    DataFrame을 엑셀 파일로 저장합니다.
    SAP 수기 다운로드와 동일한 형태로 저장합니다.
    - 숫자 셀은 실제 숫자로 변환 + 통화 서식 적용 (경고 삼각형 제거)
    - 헤더 없이, 열 너비만 자동 조정
    sheet_suffix: 시트명 뒤에 붙을 연월 (예: "2603" → 시트명 "재무(내부)_2603")
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    sheet_name = f"재무(내부)_{sheet_suffix}" if sheet_suffix else "재무(내부)"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # index=False: 행 번호 숨김 / header=False: 컬럼명 숨김 (SAP 원본 그대로)
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        ws = writer.sheets[sheet_name]

        # 숫자 셀 변환: "137,812,868,632" → 숫자 137812868632 + 통화 서식
        # (Excel 경고 삼각형 = 숫자가 텍스트로 저장된 경우 발생 → 실제 숫자로 바꿔서 해결)
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                cleaned = cell.value.replace(",", "").strip()
                if cleaned == "":
                    continue
                try:
                    # 소수점이 없으면 정수, 있으면 실수로 변환
                    if "." in cleaned:
                        cell.value = float(cleaned)
                    else:
                        cell.value = int(cleaned)
                    # 통화 서식: 1,000 단위 구분 + 소수점 없음 (회계 숫자 표기)
                    cell.number_format = "#,##0"
                except ValueError:
                    pass  # 숫자가 아닌 텍스트는 그대로 유지

        # 열 너비 자동 조정
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    print(f"  [저장 완료] {output_path}")


# -------------------------------------------------------
# 메인 실행
# -------------------------------------------------------
def main():
    # ── 종료기간 입력 받기 ──────────────────────────────
    print("=" * 55)
    print(f"  SAP {SAP_TCODE} - IFRS재무상태표(내부용)")
    print("=" * 55)
    print()
    while True:
        period = input("  조회할 종료기간(월)을 입력하세요 (1~12): ").strip()
        if period.isdigit() and 1 <= int(period) <= 12:
            break
        print("  [오류] 1~12 사이의 숫자를 입력하세요.")

    output_file = os.path.join(OUTPUT_DIR, f"Y_OKD_27000037_{FISCAL_YEAR}년{period}월.xlsx")

    print()
    print(f"  회계연도    : {FISCAL_YEAR}")
    print(f"  종료기간    : {period}월")
    print(f"  전기종료기간: {PRIOR_PERIOD} (고정)")
    print(f"  저장 위치   : {output_file}")
    print()

    # 1단계: SAP 연결
    session = connect_sap()
    print()

    # 2단계: T코드 이동
    print("  T코드로 이동 중...")
    navigate_to_tcode(session)

    # 3단계: 조건 입력
    input_conditions(session, period)
    print()

    # 4단계: 조회 실행 (F8)
    print("  조회 실행 중...")
    execute_report(session)
    print()

    # 5단계: 전체 데이터 한 번에 내보내기
    print("  데이터 추출 중...")
    df = download_all_at_once(session)
    if df.empty:
        print("  [종료] 저장할 데이터가 없습니다.")
        return
    print()

    # 6단계: 컬럼 정리 (B~U열 + 빈 컬럼 삭제)
    print("  컬럼 정리 중...")
    df = clean_columns(df)
    print()

    # 7단계: 엑셀 저장
    print("  엑셀 파일 저장 중...")
    # 시트명: 재무(내부)_YYPP 형식 (예: 2603 = 2026년 3월)
    sheet_suffix = f"{FISCAL_YEAR[2:]}{period.zfill(2)}"
    save_to_excel(df, output_file, sheet_suffix)
    print()

    print("=" * 55)
    print(f"  완료! 총 {len(df):,}건")
    print(f"  저장 위치: {output_file}")
    print("=" * 55)


if __name__ == "__main__":
    main()
