# -*- coding: utf-8 -*-
# =====================================================
# SAP Y_OKD_27000039 데이터 조회 및 다운로드
# - IFRS순익계산서(내부용)
# - 기간 시작/종료기간: 실행 시 입력받아 조회
# - 기간/연도 /$PPF, /$PFFP: 자동 계산 (고정)
# - 스프레드시트([1,0]) XLS 파일 저장 후 엑셀로 변환
# =====================================================

import win32com.client          # SAP GUI를 파이썬으로 조작하는 도구
import pandas as pd             # 표 데이터를 엑셀로 저장하는 도구
import time                     # 대기 시간 도구
import os                       # 폴더/파일 경로 도구
from datetime import date       # 오늘 날짜 도구

# -------------------------------------------------------
# [설정] 여기서 원하는 값을 바꾸세요!
# -------------------------------------------------------

# 회계연도: 오늘 날짜 기준 자동 설정
TODAY       = date.today()
FISCAL_YEAR = str(TODAY.year)        # 당기 회계연도 (예: "2026")
PRIOR_YEAR  = str(TODAY.year - 1)    # 전기 회계연도 (예: "2025")

# 기간/연도 고정값 (SAP 녹화 기준)
# ctxtPAR_07 = /$PPF  = 당기연도 + "001"  (예: "2026001")
# ctxtPAR_06 = /$PFFP = 전기연도 + "001"  (예: "2025001")
PPF_VALUE  = FISCAL_YEAR + "001"
PFFP_VALUE = PRIOR_YEAR  + "001"

# 결과 저장 폴더
OUTPUT_DIR = "data/output"

# T코드
SAP_TCODE = "Y_OKD_27000039"


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
# 조회 조건 입력 (SAP 녹화 스크립트 기준)
# -------------------------------------------------------
def input_conditions(session, period_from, period_to):
    """
    필드 ID 및 입력 순서 — SAP 녹화 스크립트 기준:
    ctxtPAR_04: 기간 시작
    ctxtPAR_05: 전기기간시작  (기간 시작과 동일)
    ctxtPAR_02: 전기종료기간  (종료기간과 동일)
    ctxtPAR_03: 종료기간
    ctxtPAR_07: /$PPF  = 당기연도+001 (고정)
    ctxtPAR_06: /$PFFP = 전기연도+001 (고정) → 마지막 입력 후 바로 F8
    """
    # 기간 시작 — ctxtPAR_04
    session.findById("wnd[0]/usr/ctxtPAR_04").text = period_from
    session.findById("wnd[0]/usr/ctxtPAR_04").setFocus()
    session.findById("wnd[0]/usr/ctxtPAR_04").caretPosition = len(period_from)
    session.findById("wnd[0]").sendVKey(0)

    # 전기기간시작 — ctxtPAR_05 (기간 시작과 동일)
    session.findById("wnd[0]/usr/ctxtPAR_05").text = period_from
    session.findById("wnd[0]/usr/ctxtPAR_05").caretPosition = len(period_from)
    session.findById("wnd[0]").sendVKey(0)

    # 전기종료기간 — ctxtPAR_02 (종료기간과 동일)
    session.findById("wnd[0]/usr/ctxtPAR_02").text = period_to
    session.findById("wnd[0]/usr/ctxtPAR_02").caretPosition = len(period_to)
    session.findById("wnd[0]").sendVKey(0)

    # 종료기간 — ctxtPAR_03
    session.findById("wnd[0]/usr/ctxtPAR_03").text = period_to
    session.findById("wnd[0]/usr/ctxtPAR_03").caretPosition = len(period_to)
    session.findById("wnd[0]").sendVKey(0)

    # 기간/연도 /$PPF — ctxtPAR_07 (고정)
    session.findById("wnd[0]/usr/ctxtPAR_07").text = PPF_VALUE
    session.findById("wnd[0]/usr/ctxtPAR_07").caretPosition = len(PPF_VALUE)
    session.findById("wnd[0]").sendVKey(0)

    # 기간/연도 /$PFFP — ctxtPAR_06 (고정, 마지막 입력)
    session.findById("wnd[0]/usr/ctxtPAR_06").text = PFFP_VALUE
    session.findById("wnd[0]/usr/ctxtPAR_06").caretPosition = len(PFFP_VALUE)
    # Enter 없이 F8로 바로 넘어감 (녹화 기준)

    time.sleep(0.3)
    print(f"  [조건 입력] 기간시작={period_from}월, 종료기간={period_to}월, "
          f"/$PPF={PPF_VALUE}, /$PFFP={PFFP_VALUE}")


# -------------------------------------------------------
# F8 실행 (조회) + 팝업 처리
# -------------------------------------------------------
def execute_report(session):
    """
    F8 키로 조회 실행.
    손익계산서는 F8 후 드릴다운 옵션 팝업이 뜸 → [0,0] 첫 번째 옵션 선택 후 OK
    """
    session.findById("wnd[0]").sendVKey(8)  # F8
    time.sleep(2.0)

    # F8 후 팝업 처리: 드릴다운 결과 선택 (radCEC01-CHOICE[0,0])
    try:
        session.findById(
            "wnd[1]/usr/sub:SAPLKEC1:0110/radCEC01-CHOICE[0,0]"
        ).select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(2.0)
        print("  [팝업] 드릴다운 옵션 선택 완료")
    except Exception:
        time.sleep(1.0)  # 팝업 없으면 그냥 대기

    print("  [조회 완료] 결과 화면 로드됨")


# -------------------------------------------------------
# 전체 데이터 한 번에 내보내기 (SAP 녹화 기준)
# -------------------------------------------------------
def download_all_at_once(session, xls_filename):
    """
    SAP 녹화 스크립트 기준:
    스프레드시트([1,0]) 형식으로 XLS 파일 저장 후 파이썬으로 읽어옵니다.
    SAP 스프레드시트 = UTF-16 LE 인코딩 탭 구분 텍스트 (.XLS 확장자)
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_dir   = os.path.abspath(OUTPUT_DIR)
    temp_path = os.path.join(out_dir, xls_filename)

    # ① 시스템 → 리스트 → 저장 → 로컬 파일 메뉴 클릭
    print("  [메뉴] 시스템 → 리스트 → 저장 → 로컬 파일 클릭...")
    session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select()
    time.sleep(1.0)

    # ② 형식 선택: [1,0] 스프레드시트
    radio = session.findById(
        "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150"
        "/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
    )
    radio.select()
    radio.setFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(0.8)

    # ③ 파일명 입력 후 저장 (녹화 기준: wnd[1]에서 파일명만 입력)
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = out_dir + "\\"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = xls_filename
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(xls_filename)
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(2.0)

    # 덮어쓰기 확인 팝업 처리 (파일이 이미 있을 때)
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1.0)
    except Exception:
        pass

    print(f"  [파일 저장] {temp_path}")

    # ④ XLS 파일 읽기 (UTF-16 LE 탭 구분 텍스트)
    df = pd.read_csv(
        temp_path,
        sep="\t",
        encoding="utf-16",
        dtype=str,
        header=None,    # SAP 원본 형태(제목·메타데이터) 그대로 유지
    )

    # ⑤ 임시 XLS 파일 삭제
    try:
        os.remove(temp_path)
    except Exception:
        pass

    df = df.fillna("").astype(str)
    print(f"  [내보내기 완료] {len(df):,}행 × {len(df.columns)}열")
    return df


# -------------------------------------------------------
# 컬럼 정리 (빈 컬럼 삭제)
# -------------------------------------------------------
def clean_columns(df):
    """값이 하나도 없는 빈 컬럼을 삭제합니다."""
    before = len(df.columns)
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
    SAP 수기 다운로드와 동일한 형태로 Excel 저장.
    숫자 셀은 실제 숫자로 변환 + #,##0 통화 서식 적용.
    sheet_suffix: 시트명 뒤에 붙을 연월 (예: "2603")
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    sheet_name = f"손익(내부)_{sheet_suffix}" if sheet_suffix else "손익(내부)"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        ws = writer.sheets[sheet_name]

        # 숫자 문자열 → 실제 숫자 변환 + 통화 서식 (경고 삼각형 제거)
        for row in ws.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                cleaned = cell.value.replace(",", "").strip()
                if cleaned == "":
                    continue
                try:
                    cell.value = float(cleaned) if "." in cleaned else int(cleaned)
                    cell.number_format = "#,##0"
                except ValueError:
                    pass

        # 열 너비 자동 조정
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    print(f"  [저장 완료] {output_path}")


# -------------------------------------------------------
# 메인 실행
# -------------------------------------------------------
def main():
    print("=" * 55)
    print(f"  SAP {SAP_TCODE} - IFRS순익계산서(내부용)")
    print("=" * 55)
    print()

    # ── 기간 시작/종료 입력 ─────────────────────────────
    while True:
        period_from = input("  기간 시작(월)을 입력하세요 (1~12): ").strip()
        if period_from.isdigit() and 1 <= int(period_from) <= 12:
            break
        print("  [오류] 1~12 사이의 숫자를 입력하세요.")

    while True:
        period_to = input("  종료기간(월)을 입력하세요 (1~12): ").strip()
        if period_to.isdigit() and 1 <= int(period_to) <= 12:
            if int(period_to) >= int(period_from):
                break
            print(f"  [오류] 종료기간은 기간 시작({period_from})보다 크거나 같아야 합니다.")
        else:
            print("  [오류] 1~12 사이의 숫자를 입력하세요.")

    # 임시 XLS 파일명 (녹화 기준 형식: 손익(내부)_YYPP.XLS)
    xls_filename = f"손익(내부)_{FISCAL_YEAR[2:]}{period_to.zfill(2)}.XLS"
    output_file  = os.path.join(
        OUTPUT_DIR,
        f"Y_OKD_27000039_{FISCAL_YEAR}년{period_from}~{period_to}월.xlsx"
    )

    print()
    print(f"  회계연도    : {FISCAL_YEAR}")
    print(f"  기간 시작   : {period_from}월")
    print(f"  종료기간    : {period_to}월")
    print(f"  /$PPF       : {PPF_VALUE} (고정)")
    print(f"  /$PFFP      : {PFFP_VALUE} (고정)")
    print(f"  저장 위치   : {output_file}")
    print()

    # 1단계: SAP 연결
    session = connect_sap()
    print()

    # 2단계: T코드 이동
    print("  T코드로 이동 중...")
    navigate_to_tcode(session)

    # 3단계: 조건 입력
    input_conditions(session, period_from, period_to)
    print()

    # 4단계: 조회 실행 (F8) + 팝업 처리
    print("  조회 실행 중...")
    execute_report(session)
    print()

    # 5단계: 전체 데이터 XLS로 내보내기
    print("  데이터 추출 중...")
    df = download_all_at_once(session, xls_filename)
    if df.empty:
        print("  [종료] 저장할 데이터가 없습니다.")
        return
    print()

    # 6단계: 빈 컬럼 정리
    print("  컬럼 정리 중...")
    df = clean_columns(df)
    print()

    # 7단계: 엑셀 저장
    print("  엑셀 파일 저장 중...")
    sheet_suffix = f"{FISCAL_YEAR[2:]}{period_to.zfill(2)}"
    save_to_excel(df, output_file, sheet_suffix)
    print()

    print("=" * 55)
    print(f"  완료! 총 {len(df):,}건")
    print(f"  저장 위치: {output_file}")
    print("=" * 55)


if __name__ == "__main__":
    main()
