# -*- coding: utf-8 -*-
# =====================================================
# SAP 필드 ID 조회 스크립트
# y_okd_27000037 선택화면에서 실행하세요.
#
# [사용 방법]
# 1. SAP에서 y_okd_27000037 을 실행해 선택화면을 열어두세요.
# 2. 이 스크립트를 실행하세요.
# 3. 결과가 'y_okd_field_dump.txt' 파일로 저장됩니다.
# 4. 그 파일을 Claude에게 보여주세요.
# =====================================================

import win32com.client  # SAP GUI를 파이썬으로 조작하는 도구

OUTPUT_FILE = "y_okd_field_dump.txt"  # 결과 저장 파일 이름


def dump_tree(obj, depth=0, lines=None):
    """
    SAP GUI 화면의 모든 요소(필드, 버튼 등)를 재귀적으로 탐색합니다.
    회계 장부에서 모든 항목을 순서대로 나열하는 것과 같아요.
    """
    if lines is None:
        lines = []

    indent = "  " * depth  # 들여쓰기 (깊이 표현)

    try:
        obj_type = obj.Type        # 요소 종류 (예: GuiTextField, GuiButton)
        obj_id   = obj.Id          # 요소 고유 경로 (예: wnd[0]/usr/ctxtP_DATE)
        obj_name = getattr(obj, "Name", "")   # 이름
        obj_text = getattr(obj, "Text", "")   # 표시 텍스트 (라벨 등)

        # 타입별 추가 정보 수집
        extra = ""
        if obj_type in ("GuiTextField", "GuiCTextField"):
            extra = f"  ← 입력필드 | 현재값: '{obj_text}'"
        elif obj_type == "GuiRadioButton":
            extra = f"  ← 라디오버튼 | 텍스트: '{obj_text}'"
        elif obj_type == "GuiCheckBox":
            extra = f"  ← 체크박스 | 텍스트: '{obj_text}'"
        elif obj_type == "GuiComboBox":
            extra = f"  ← 드롭다운 | 현재값: '{obj_text}'"
        elif obj_type == "GuiButton":
            extra = f"  ← 버튼 | 텍스트: '{obj_text}'"
        elif obj_type in ("GuiShell", "GuiGridView"):
            row_count = getattr(obj, "RowCount", "?")
            col_count = getattr(obj, "ColumnCount", "?")
            extra = f"  ← ALV Grid | 행={row_count} 열={col_count}"

        line = f"{indent}[{obj_type}] {obj_id}{extra}"
        lines.append(line)

    except Exception:
        return lines

    # 자식 요소들도 재귀 탐색
    try:
        for i in range(obj.Children.Count):
            dump_tree(obj.Children.ElementAt(i), depth + 1, lines)
    except Exception:
        pass

    return lines


def main():
    print("=" * 55)
    print("  SAP 필드 ID 조회 스크립트 시작")
    print("=" * 55)
    print("  SAP에서 y_okd_27000037 선택화면을 열어두세요.")
    print()

    # SAP GUI 연결
    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app     = sap_gui.GetScriptingEngine
        conn    = app.Children(0)
        session = conn.Children(0)
        print(f"  [연결 성공] 시스템: {session.Info.SystemName}")
        print(f"  [현재 화면] T코드: {session.Info.Transaction}")
        print()
    except Exception as e:
        print(f"  [오류] SAP 연결 실패: {e}")
        print("  SAP GUI가 실행 중인지, 스크립팅이 활성화되어 있는지 확인하세요.")
        return

    # 현재 화면 전체 탐색
    print("  화면 요소 탐색 중...")
    try:
        window = session.findById("wnd[0]")
        lines  = dump_tree(window)
    except Exception as e:
        print(f"  [오류] 화면 탐색 실패: {e}")
        return

    # 결과 파일 저장
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(f"=== SAP 필드 덤프 결과 ===\n")
        f.write(f"시스템: {session.Info.SystemName}\n")
        f.write(f"T코드: {session.Info.Transaction}\n")
        f.write(f"총 요소 수: {len(lines)}개\n")
        f.write("=" * 55 + "\n\n")
        f.write("\n".join(lines))

    print(f"  [완료] '{OUTPUT_FILE}' 파일로 저장됐습니다.")
    print()
    print("  다음 단계:")
    print(f"  → '{OUTPUT_FILE}' 파일 내용을 Claude에게 붙여넣어 주세요.")
    print("  → 그러면 다운로드 스크립트를 완성해 드릴게요!")
    print("=" * 55)


if __name__ == "__main__":
    main()
