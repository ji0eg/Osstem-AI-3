# -*- coding: utf-8 -*-
# =====================================================
# SAP ZSDR0026 계약별 장비/재료 잔액조회 자동화
# - 사업장 전체 순회
# - 계약진행상태: 전체 선택
# - ALV → 클립보드 내보내기 → MySQL DB 저장
# =====================================================

import win32com.client   # SAP GUI를 파이썬으로 조작하는 도구
import pyperclip         # 클립보드(복사/붙여넣기) 도구
import pymysql           # MySQL DB 연결 도구
import time              # 시간 지연(대기) 도구
from datetime import datetime, date  # 날짜 처리 도구
from dotenv import load_dotenv       # .env 파일에서 설정 읽는 도구
import os                            # 환경변수 읽는 도구

# .env 파일 로드 (비밀번호, T코드 등 민감한 설정값 불러오기)
load_dotenv()

# -------------------------------------------------------
# [설정] 여기서 원하는 값을 바꾸세요!
# -------------------------------------------------------

# ▼ 기준일자 설정 (빈 문자열이면 오늘 날짜 자동 사용)
# 형식: "YYYYMMDD"  예) "20260315"
BASE_DATE = "20260315"

# 데이터가 없는 사업장은 건너뜀 여부 (True=건너뜀, False=계속)
SKIP_EMPTY = True

# DB에 한 번에 저장하는 행 수 (너무 크면 느려질 수 있음)
BATCH_SIZE = 500

# ALV 클립보드 내보내기 후 대기시간(초) — 컴퓨터가 느리면 늘리세요
CLIPBOARD_WAIT = 3.0

# 금액으로 처리할 컬럼 이름 목록 (쉼표·마이너스 기호 → 숫자 변환)
NUMERIC_COLS = {
    '총계약금액', '총이용금액', '총계약잔액',
    '제품계약금액', '제품이용금액', '제품잔액',
    '장비계약금액', '장비이용금액', '남은장비잔액',
    '재료계약금액', '재료이용금액', '남은재료잔액',
}

# 사업장 코드 목록
PLANT_CODES = [
    '1010','1020','1030','1040','1050','1060','1070','1080','1090',
    '1110','1120','1130','1140','1150','1160','1170','1180','1190',
    '1210','1220','1230','1240','1250','1260','1270','1280','1290',
    '1310','1320','1330','1340','1350','1360','1370','1380','1390',
    '1410','1420','1430','1440','1450','1460','1470','1480','1490',
    '1510','1520','1530','1540','1550','1560','1570','1580','1590',
    '1610','1620','1630','1640','1650','1660','1670','1680','1690',
    '1710','1720','1730','1740','1750','1760','1770','1780','1790',
    '1810','1820','1830','1840','1850','1860','1870','1880','1890',
    '1910','1920','1930','1940','1950','1960','1970','1980','1990',
    '2010','2020','2030','2040','2050','2060','2070','2080','2090',
    '2110','2120','2130','2140','2150','2160','2170','2180','2190',
    '2210','2220','2230','2240','2250','2260','2270','2280','2290',
    '2310','2320','2330','2340','2350','2360','2370','2380','2390',
    '2410','2420','2430','2440','2450','2460','2470','2480','2490',
    '2510','2520','2530','2540','2550','2560','2570','2580','2590',
    '2610','2620','2630','2640','2650','2660','2670','2680','2690',
    '2710','2720','2730','2740','2750','2760','2770','2780','2790',
    '9010','9020',
]

# -------------------------------------------------------
# .env에서 민감한 설정값 읽기
# -------------------------------------------------------
SAP_TCODE = os.getenv("SAP_TCODE", "ZSDR0026")  # T코드
DB_HOST   = os.getenv("DB_HOST")                 # DB 서버 주소
DB_PORT   = int(os.getenv("DB_PORT", "5010"))    # DB 포트 번호
DB_USER   = os.getenv("DB_USER")                 # DB 계정
DB_PASS   = os.getenv("DB_PASS")                 # DB 비밀번호
DB_NAME   = os.getenv("DB_NAME", "ERP_Translate") # DB 이름
DB_TABLE  = "ZSDR0026_잔액조회"                  # 저장할 테이블 이름

# BASE_DATE가 비어있으면 오늘 날짜로 자동 설정
if not BASE_DATE:
    BASE_DATE = date.today().strftime("%Y%m%d")

# SAP 화면 입력용 날짜 형식 변환: "20260315" → "2026.03.15"
SAP_DATE = datetime.strptime(BASE_DATE, "%Y%m%d").strftime("%Y.%m.%d")


# -------------------------------------------------------
# SAP 연결
# -------------------------------------------------------
def connect_sap():
    """SAP GUI에 COM(윈도우 자동화 방식)으로 연결"""
    sap_gui = win32com.client.GetObject("SAPGUI")   # SAP GUI 프로그램 가져오기
    app     = sap_gui.GetScriptingEngine             # 스크립팅 엔진 연결
    conn    = app.Children(0)                         # 첫 번째 연결
    session = conn.Children(0)                        # 첫 번째 세션(창)
    print(f"[SAP] 연결 성공: {session.Info.SystemName}")
    return session


# -------------------------------------------------------
# T-Code 이동
# -------------------------------------------------------
def navigate_to_tcode(session, tcode):
    """SAP 화면 상단 입력창에 T코드를 입력하고 Enter"""
    session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{tcode}"  # /n = 새로 이동
    session.findById("wnd[0]").sendVKey(0)  # Enter
    time.sleep(1.5)  # 화면 로딩 대기


# -------------------------------------------------------
# 선택화면 조건 입력
# -------------------------------------------------------
def input_conditions(session, plant_code):
    """기준일자, 사업장, 계약진행상태(전체) 입력"""
    session.findById("wnd[0]/usr/ctxtP_LFDAT").text        = SAP_DATE    # 기준일자
    session.findById("wnd[0]/usr/ctxtS_VKBUR-LOW").text    = plant_code  # 사업장
    session.findById("wnd[0]/usr/radR2").select()                         # 계약진행상태: 전체
    time.sleep(0.3)


# -------------------------------------------------------
# F8 실행 (조회)
# -------------------------------------------------------
def execute_report(session):
    """F8 키를 눌러 조회 실행"""
    session.findById("wnd[0]").sendVKey(8)  # 8 = F8
    time.sleep(2.5)  # 결과 화면 로딩 대기


# -------------------------------------------------------
# ALV Grid 찾기
# -------------------------------------------------------
def find_alv_grid(session):
    """ZSDR0026 결과 화면의 ALV Grid 객체 반환"""
    try:
        return session.findById("wnd[0]/shellcont/shell")
    except Exception as e:
        raise RuntimeError(f"ALV Grid를 찾지 못했습니다: {e}")


# -------------------------------------------------------
# 클립보드로 내보내기
# -------------------------------------------------------
def extract_alv_to_clipboard(session):
    """ALV 데이터를 클립보드(복사)로 내보내고 텍스트 반환"""
    shell     = find_alv_grid(session)
    row_count = 0

    # 데이터 행 수 확인
    try:
        row_count = shell.RowCount
    except Exception:
        pass

    if row_count == 0:
        print("    ↳ 데이터 없음 (RowCount=0)")
        return None

    print(f"    ↳ ALV RowCount={row_count}")

    # 전체 행 선택
    try:
        shell.setCurrentCell(0, shell.ColumnOrder[0])
        shell.sendVKey(11)   # Ctrl+End (마지막 셀로)
        shell.sendVKey(999)  # Ctrl+Home (첫 셀로)
        shell.selectedRows = f"0-{row_count - 1}"  # 0번~마지막 행 선택
    except Exception:
        pass

    # 클립보드 초기화
    try:
        pyperclip.copy("")
    except Exception:
        pass
    time.sleep(0.2)

    # SAP 내보내기 버튼 → 클립보드(스프레드시트) 선택
    try:
        shell.pressToolbarContextButton("&MB_EXPORT")
        time.sleep(0.8)
        shell.selectContextMenuItem("&PC")
        time.sleep(0.8)
    except Exception as e:
        raise RuntimeError(f"내보내기 버튼 실행 실패: {e}")

    # 형식 선택 팝업창: 클립보드 항목 선택 후 확인
    try:
        radio = session.findById(
            "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150"
            "/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]"
        )
        radio.select()
        radio.setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # 확인 버튼
        time.sleep(CLIPBOARD_WAIT)  # 클립보드 복사 완료 대기
    except Exception as e:
        raise RuntimeError(f"형식 선택 창 처리 실패: {e}")

    # 클립보드에서 텍스트 읽기
    text = pyperclip.paste()
    if not text or len(text) < 10:
        raise RuntimeError("클립보드가 비어있습니다.")
    return text


# -------------------------------------------------------
# 금액 문자열 → 숫자 변환
# -------------------------------------------------------
def _clean_number(s):
    """'280,000-' 같은 SAP 금액 문자열을 float 숫자로 변환"""
    s = s.strip().replace(',', '')  # 쉼표 제거
    if not s or s == '0':
        return 0.0
    negative = s.endswith('-')      # 음수 여부 확인 (SAP는 뒤에 - 붙음)
    s = s.rstrip('-').strip()
    try:
        val = float(s)
        return -val if negative else val
    except ValueError:
        return None


# -------------------------------------------------------
# 클립보드 텍스트 → 레코드 리스트 파싱
# -------------------------------------------------------
def parse_clipboard_text(text, plant_code, run_dt):
    """
    클립보드에서 복사된 파이프(|) 구분 텍스트를 파싱해서
    DB에 저장할 딕셔너리 리스트로 반환
    """
    lines = text.splitlines()

    # 헤더 줄 찾기 ('계약번호'가 있는 줄)
    header_line = None
    header_idx  = None
    for i, line in enumerate(lines):
        if '|' in line and '계약번호' in line:
            header_line = line
            header_idx  = i
            break

    if header_line is None:
        print("    ↳ 헤더를 찾지 못했습니다.")
        return []

    # 헤더 컬럼 이름 파싱
    headers = [h.strip() for h in header_line.split('|')]
    headers = [h for h in headers if h]  # 앞뒤 빈 문자열 제거
    print(f"    ↳ 헤더 컬럼 수: {len(headers)} → {headers}")

    # SAP 화면 헤더명 → DB 컬럼명 매핑표
    HEADER_TO_COL = {
        '부서':         '부서',
        '담당자':       '담당자',
        '거래처':       '거래처',
        '거래처명':     '거래처명',
        '계약번호':     '계약번호',
        '계약구분':     '계약구분',
        '진행상태':     '진행상태_코드',  # 첫 번째 진행상태 = 코드 (1/X/Y)
        '약정/할인':    '약정할인',
        '총계약금액':   '총계약금액',
        '총이용금액':   '총이용금액',
        '총계약잔액':   '총계약잔액',
        '제품계약금액': '제품계약금액',
        '제품이용금액': '제품이용금액',
        '제품잔액':     '제품잔액',
        '장비계약금액': '장비계약금액',
        '장비이용금액': '장비이용금액',
        '남은장비잔액': '남은장비잔액',
        '재료계약금액': '재료계약금액',
        '재료이용금액': '재료이용금액',
        '남은재료잔액': '남은재료잔액',
    }

    rows = []
    for line in lines[header_idx + 1:]:
        stripped = line.strip()
        if not stripped:
            continue  # 빈 줄 스킵
        if all(c in '-|+ ' for c in stripped):
            continue  # 구분선 스킵
        if '|' not in line:
            continue

        # 파이프(|)로 분리, 양끝 빈 요소 제거
        parts = line.split('|')
        parts = parts[1:-1]

        # 컬럼 수가 부족하면 빈 값 채우기
        while len(parts) < len(headers):
            parts.append('')

        # 공백 제거, 빈 값은 None 처리
        vals = [v.strip() if v.strip() else None for v in parts]

        # 계약번호 없으면 스킵 (합계행 등)
        raw = dict(zip(headers, vals[:len(headers)]))
        if not raw.get('계약번호'):
            continue

        # 헤더명 → DB 컬럼명 변환
        # '진행상태'가 두 번 나올 경우: 첫 번째=코드, 두 번째=텍스트
        record = {}
        seen_진행상태 = 0
        for idx_h, h in enumerate(headers):
            val = vals[idx_h] if idx_h < len(vals) else None
            if h == '진행상태':
                if seen_진행상태 == 0:
                    record['진행상태_코드'] = val   # 코드 (1/X/Y)
                else:
                    record['진행상태'] = val         # 텍스트 (진행/완료/취소)
                seen_진행상태 += 1
            else:
                col = HEADER_TO_COL.get(h, h)
                if col in NUMERIC_COLS:
                    record[col] = _clean_number(val) if val else 0.0
                else:
                    record[col] = val

        # 누락된 금액 컬럼은 0으로 채우기
        for col in NUMERIC_COLS:
            if col not in record:
                record[col] = 0.0

        # 공통 필드 추가
        record['기준일자']     = BASE_DATE
        record['사업장']       = plant_code
        record['배치실행일시'] = run_dt
        rows.append(record)

    print(f"    ↳ 파싱 완료: {len(rows):,}건")
    return rows


# -------------------------------------------------------
# DB 연결
# -------------------------------------------------------
def get_db_conn():
    """MySQL DB에 연결 후 연결 객체 반환"""
    return pymysql.connect(
        host=DB_HOST,
        port=DB_PORT,
        user=DB_USER,
        password=DB_PASS,
        database=DB_NAME,
        charset='utf8mb4',
        autocommit=False,  # 수동 커밋 (오류 시 롤백 가능)
    )


# -------------------------------------------------------
# DB 저장 (UPSERT)
# -------------------------------------------------------
def save_to_db(rows, conn):
    """
    레코드를 DB에 저장 (이미 있으면 업데이트, 없으면 삽입)
    회계 장부에서 같은 전표번호가 있으면 수정, 없으면 새로 기입하는 것과 동일
    """
    if not rows:
        return 0

    # DB에 저장할 컬럼 순서
    cols = [
        '사업장', '부서', '담당자', '거래처', '거래처명', '계약번호', '계약구분',
        '진행상태_코드', '약정할인',
        '총계약금액', '총이용금액', '총계약잔액',
        '제품계약금액', '제품이용금액', '제품잔액',
        '장비계약금액', '장비이용금액', '남은장비잔액',
        '재료계약금액', '재료이용금액', '남은재료잔액',
        '진행상태', '기준일자', '배치실행일시',
    ]

    placeholders  = ', '.join(['%s'] * len(cols))
    col_list      = ', '.join([f'`{c}`' for c in cols])
    # 중복 시 업데이트할 컬럼 (기준일자·계약번호는 키이므로 제외)
    update_clause = ', '.join([f'`{c}`=VALUES(`{c}`)' for c in cols if c not in ('기준일자', '계약번호')])

    sql = (
        f"INSERT INTO `{DB_TABLE}` ({col_list}) "
        f"VALUES ({placeholders}) "
        f"ON DUPLICATE KEY UPDATE {update_clause}"
    )

    total = 0
    with conn.cursor() as cur:
        batch = []
        for row in rows:
            batch.append(tuple(row.get(c) for c in cols))
            if len(batch) >= BATCH_SIZE:  # 배치 사이즈에 도달하면 저장
                cur.executemany(sql, batch)
                total += len(batch)
                batch = []
        if batch:  # 남은 데이터 저장
            cur.executemany(sql, batch)
            total += len(batch)
    conn.commit()  # 최종 확정 저장
    return total


# -------------------------------------------------------
# 데이터 없음 확인
# -------------------------------------------------------
def has_no_data(session):
    """조회 결과가 없는 경우 감지 (상태바 메시지 또는 팝업창)"""
    try:
        status   = session.findById("wnd[0]/sbar").text
        keywords = ['데이터 없음', 'No data', '선택된 데이터 없음', '조회 결과가 없습니다']
        if any(k in status for k in keywords):
            return True
    except Exception:
        pass

    # 팝업창이 열려있으면 닫고 True 반환
    try:
        session.findById("wnd[1]")
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            pass
        return True
    except Exception:
        pass

    return False


# -------------------------------------------------------
# 메인 실행
# -------------------------------------------------------
def main():
    print("=" * 60)
    print(f" SAP {SAP_TCODE} 자동화 시작")
    print(f" 기준일자: {BASE_DATE}")
    print(f" 사업장 수: {len(PLANT_CODES)}개")
    print("=" * 60)

    session = connect_sap()
    conn    = get_db_conn()
    run_dt  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # 실행 시각 기록

    total_saved   = 0   # 총 저장 건수
    total_skipped = 0   # 스킵된 사업장 수
    failed_plants = []  # 오류 발생 사업장 목록

    for idx, plant in enumerate(PLANT_CODES, 1):
        print(f"\n[{idx:3d}/{len(PLANT_CODES)}] 사업장: {plant}")
        try:
            navigate_to_tcode(session, SAP_TCODE)   # T코드 이동
            input_conditions(session, plant)          # 조건 입력
            execute_report(session)                   # F8 조회

            # 데이터 없으면 스킵
            if has_no_data(session):
                print("    ↳ 데이터 없음 → 스킵")
                total_skipped += 1
                continue

            # 클립보드로 데이터 추출
            text = extract_alv_to_clipboard(session)
            if not text:
                total_skipped += 1
                continue

            # 텍스트 파싱 → 레코드 리스트
            rows = parse_clipboard_text(text, plant, run_dt)
            if not rows:
                print("    ↳ 파싱 결과 0건 → 스킵")
                total_skipped += 1
                continue

            # DB 저장
            saved        = save_to_db(rows, conn)
            total_saved += saved
            print(f"    ↳ 저장: {saved:,}건  (누계: {total_saved:,}건)")

        except Exception as e:
            print(f"    ↳ 오류: {e}")
            failed_plants.append((plant, str(e)))
            # 오류 발생 시 T코드로 돌아가서 다음 사업장 계속
            try:
                navigate_to_tcode(session, SAP_TCODE)
            except Exception:
                pass
            continue

    conn.close()

    # 최종 결과 출력
    print("\n" + "=" * 60)
    print(" 완료!")
    print(f" 총 저장: {total_saved:,}건")
    print(f" 스킵:    {total_skipped}개 사업장")
    if failed_plants:
        print(f" 실패:    {len(failed_plants)}개 사업장")
        for p, err in failed_plants:
            print(f"   - {p}: {err}")
    print("=" * 60)


# 이 파일을 직접 실행할 때만 main() 동작
if __name__ == "__main__":
    main()
