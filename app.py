# -*- coding: utf-8 -*-
# =====================================================
# [오스템 회계팀] Streamlit 웹 앱
# 손익계산서(내부)_download.py 핵심 함수를 재사용하여
# 브라우저에서 SAP 조회 → 결과 확인 → 다운로드 가능
# =====================================================

import streamlit as st
import pandas as pd
import importlib.util   # 파이썬에서 다른 .py 파일을 불러오는 도구
import io               # 메모리 안에서 파일처럼 데이터를 다루는 도구
import os
from datetime import date

# -------------------------------------------------------
# 손익계산서 모듈 불러오기
# (파일명에 괄호가 있어 일반 import 대신 importlib 사용)
# -------------------------------------------------------
_spec = importlib.util.spec_from_file_location(
    "sap_pnl",
    os.path.join(os.path.dirname(__file__), "손익계산서(내부)_download.py")
)
sap = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(sap)


# -------------------------------------------------------
# 샘플 데이터 (SAP 없이 테스트할 때 사용)
# -------------------------------------------------------
def get_sample_data():
    """실제 손익계산서와 유사한 샘플 DataFrame 반환"""
    rows = [
        ["IFRS순익계산서(내부용)", "", "", "", "", "현재 데이터(샘플)"],
        ["", "", "", "", "", ""],
        ["통화", "KRW", "", "대한민국 원", "", ""],
        ["회사 코드", "", "1000", "", "오스템임플란트(주)", ""],
        ["", "", "", "", "", ""],
        ["", "", "", "당기", "", "전기"],
        ["", "", "", "", "", ""],
        ["I. 매출액",        "", "", "150,000,000,000", "", "130,000,000,000"],
        ["  1) 제품매출",    "", "", "120,000,000,000", "", "105,000,000,000"],
        ["  2) 상품매출",    "", "",  "30,000,000,000", "",  "25,000,000,000"],
        ["II. 매출원가",     "", "",  "90,000,000,000", "",  "78,000,000,000"],
        ["III. 매출총이익",  "", "",  "60,000,000,000", "",  "52,000,000,000"],
        ["IV. 판매비와관리비","","",  "25,000,000,000", "",  "22,000,000,000"],
        ["V. 영업이익",      "", "",  "35,000,000,000", "",  "30,000,000,000"],
        ["VI. 영업외수익",   "", "",   "2,000,000,000", "",   "1,500,000,000"],
        ["VII. 영업외비용",  "", "",   "1,000,000,000", "",     "800,000,000"],
        ["VIII. 법인세차감전이익","","","36,000,000,000","","  30,700,000,000"],
        ["IX. 법인세비용",   "", "",   "8,000,000,000", "",   "6,800,000,000"],
        ["X. 당기순이익",    "", "",  "28,000,000,000", "",  "23,900,000,000"],
    ]
    df = pd.DataFrame(rows)
    return df


# -------------------------------------------------------
# SAP 조회 실행 (손익계산서 모듈 함수 재사용)
# -------------------------------------------------------
def run_sap_download(period_from: str, period_to: str):
    """
    손익계산서(내부)_download.py 의 핵심 함수를 순서대로 호출합니다.
    SAP GUI가 열려 있어야 동작합니다.
    """
    fiscal_year = sap.FISCAL_YEAR
    xls_filename = f"손익(내부)_{fiscal_year[2:]}{period_to.zfill(2)}.XLS"

    session = sap.connect_sap()
    sap.navigate_to_tcode(session)
    sap.input_conditions(session, period_from, period_to)
    sap.execute_report(session)
    df = sap.download_all_at_once(session, xls_filename)
    df = sap.clean_columns(df)
    return df


# -------------------------------------------------------
# DataFrame → Excel 메모리 버퍼 변환 (다운로드 버튼용)
# -------------------------------------------------------
def df_to_excel_bytes(df: pd.DataFrame, sheet_suffix: str) -> bytes:
    """DataFrame을 엑셀 파일로 변환해서 bytes로 반환 (파일 저장 없이 브라우저 다운로드)"""
    sheet_name = f"손익(내부)_{sheet_suffix}" if sheet_suffix else "손익(내부)"
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        ws = writer.sheets[sheet_name]
        # 숫자 셀 변환 + 통화 서식
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
    return buf.getvalue()


# -------------------------------------------------------
# 앱 레이아웃
# -------------------------------------------------------
st.set_page_config(page_title="오스템 회계팀", layout="wide")

# ── 상단 제목 ───────────────────────────────────────────
st.title("🏢 [오스템 회계팀]")
st.caption("SAP IFRS 재무 데이터 자동 조회 및 다운로드")
st.divider()

# ── 사이드바 ────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 설정")

    # CSV 파일 업로드 (이전에 저장해둔 파일을 다시 볼 때 사용)
    st.subheader("📂 파일 업로드")
    uploaded_file = st.file_uploader(
        "이전에 저장한 CSV 파일을 업로드하세요",
        type=["csv"],
        help="SAP 없이도 이전 데이터를 바로 확인할 수 있어요."
    )

    st.divider()

    # 조회 파라미터 입력
    st.subheader("📅 조회 기간")
    today = date.today()
    period_from = st.number_input(
        "기간 시작(월)",
        min_value=1, max_value=12,
        value=1,
        help="조회 시작 월을 입력하세요 (1~12)"
    )
    period_to = st.number_input(
        "종료기간(월)",
        min_value=1, max_value=12,
        value=today.month,
        help="조회 종료 월을 입력하세요 (1~12)"
    )

    if period_to < period_from:
        st.warning("⚠️ 종료기간이 기간 시작보다 작습니다.")

    st.divider()
    st.caption(f"회계연도: {sap.FISCAL_YEAR}  |  /$PPF: {sap.PPF_VALUE}  |  /$PFFP: {sap.PFFP_VALUE}")

# ── 메인 화면 ───────────────────────────────────────────
col1, col2, col3 = st.columns([2, 2, 3])

with col1:
    btn_sap = st.button("📊 손익(내부) 조회", type="primary", use_container_width=True)
with col2:
    btn_sample = st.button("🧪 샘플 데이터로 실행", use_container_width=True)
with col3:
    if uploaded_file:
        st.success(f"✅ 업로드된 파일: {uploaded_file.name}")

# ── 결과 표시 영역 ──────────────────────────────────────
if "result_df" not in st.session_state:
    st.session_state.result_df = None  # 조회 결과를 임시 저장하는 공간

# 1) SAP 조회 버튼
if btn_sap:
    if uploaded_file:
        # 업로드된 CSV 파일 읽기
        df = pd.read_csv(uploaded_file, dtype=str).fillna("")
        st.session_state.result_df = df
        st.session_state.result_label = f"📂 업로드 파일 결과 ({uploaded_file.name})"
    else:
        # SAP 직접 조회
        if period_to < period_from:
            st.error("종료기간이 기간 시작보다 작습니다. 다시 입력해 주세요.")
        else:
            with st.spinner("SAP에서 데이터를 조회 중입니다... (SAP GUI가 열려 있어야 합니다)"):
                try:
                    df = run_sap_download(str(period_from), str(period_to))
                    st.session_state.result_df = df
                    st.session_state.result_label = (
                        f"📊 SAP 조회 결과 | {sap.FISCAL_YEAR}년 "
                        f"{period_from}월 ~ {period_to}월"
                    )
                    st.success("✅ 조회 완료!")
                except Exception as e:
                    st.error(f"❌ SAP 연결 오류: {e}")
                    st.info("💡 SAP GUI가 실행 중인지 확인해 주세요. 또는 '샘플 데이터로 실행' 버튼을 눌러보세요.")

# 2) 샘플 데이터 버튼
if btn_sample:
    df = get_sample_data()
    st.session_state.result_df = df
    st.session_state.result_label = "🧪 샘플 데이터 (테스트용)"
    st.info("샘플 데이터로 실행했습니다. 실제 데이터는 SAP 조회를 이용하세요.")

# 3) 결과 출력
if st.session_state.result_df is not None:
    df = st.session_state.result_df
    label = st.session_state.get("result_label", "결과")

    st.subheader(label)
    st.dataframe(df, use_container_width=True, height=500)
    st.caption(f"총 {len(df):,}행 × {len(df.columns)}열")

    # 다운로드 버튼 (CSV / Excel)
    st.divider()
    dl_col1, dl_col2 = st.columns(2)

    suffix = f"{sap.FISCAL_YEAR[2:]}{str(period_to).zfill(2)}"
    filename_base = f"손익계산서_내부_{sap.FISCAL_YEAR}년{period_from}~{period_to}월"

    with dl_col1:
        csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            label="⬇️ CSV 다운로드",
            data=csv_bytes,
            file_name=f"{filename_base}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with dl_col2:
        excel_bytes = df_to_excel_bytes(df, suffix)
        st.download_button(
            label="⬇️ Excel 다운로드",
            data=excel_bytes,
            file_name=f"{filename_base}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
