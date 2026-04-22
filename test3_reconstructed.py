import pygetwindow as gw
from pywinauto import Application
import time
import pyautogui
import pyodbc
import pandas as pd
import os
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import sys


SERVER_OPTIONS = (
    os.environ.get("DENTWEB_SQL_SERVER"),
    "127.0.0.1,1436",
    "192.168.0.245,1436",
)

CONNECTION_OPTIONS = (
    {
        "driver": "ODBC Driver 18 for SQL Server",
        "extra": "Encrypt=yes;TrustServerCertificate=yes;",
    },
    {
        "driver": "ODBC Driver 17 for SQL Server",
        "extra": "TrustServerCertificate=yes;",
    },
    {
        "driver": "SQL Server",
        "extra": "",
    },
)


def get_db_connection():
    errors = []
    server_options = [server for server in SERVER_OPTIONS if server]
    for server in server_options:
        for option in CONNECTION_OPTIONS:
            connection_string = (
                f"DRIVER={{{option['driver']}}};"
                f"SERVER={server};"
                "DATABASE=DentWeb;"
                "UID=sa;"
                "PWD=Q3xzJiwpv2zC;"
                f"{option['extra']}"
            )
            try:
                return pyodbc.connect(connection_string, timeout=5)
            except pyodbc.Error as exc:
                errors.append(f"{server} / {option['driver']}: {exc}")

    joined_errors = "\n".join(errors)
    raise RuntimeError(f"SQL Server 연결에 실패했습니다.\n{joined_errors}")


def get_patient_list():
    '''MSSQL에서 특정 조건을 만족하는 환자 리스트 가져오기'''
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        query = """
            SELECT
                A.n환자ID, A.sz차트번호,
                CASE
                    WHEN A.n보험구분 = 1 THEN '급여제한자'
                    WHEN A.n보험구분 = 2 THEN '건강 보험'
                    WHEN A.n보험구분 = 3 THEN '의료급여 1종'
                    WHEN A.n보험구분 = 4 THEN '의료급여 2종'
                    ELSE '기타'
                END AS 보험구분,
                CASE
                    WHEN A.sz접미사 IS NULL OR A.sz접미사 = '' THEN A.sz이름
                    ELSE A.sz이름 + '' + A.sz접미사
                END AS sz이름,
                A.sz생년월일, A.sz보호자휴대폰, A.sz휴대폰번호, A.sz최종내원일, A.sz주소, A.sz메모,
                DATEDIFF(MONTH, TRY_CONVERT(DATE, A.sz생년월일, 112), GETDATE()) AS 경과_개월
            FROM
                dbo.TB_환자정보 A
            WHERE
                A.sz생년월일 <= FORMAT(DATEADD(YEAR, -65, GETDATE()), 'yyyyMMdd')
                AND A.n환자ID NOT IN (
                    SELECT DISTINCT n환자ID FROM dbo.TB_보험임플환자
                )
                AND A.n보험구분 != 0 AND  A.sz주민번호 != ''
            ORDER BY A.sz이름
        """
        cursor.execute(query)
        columns = [column[0] for column in cursor.description]
        patient_list = [list(row) for row in cursor.fetchall()]
        conn.close()
        return (patient_list, columns)
    except Exception as e:
        print(f'❌ MSSQL 연결 실패. 오류: {e}')
        return ([], [])


def save_to_excel(data, columns):
    '''데이터를 Excel 파일로 저장'''
    try:
        if not data:
            print('❌ 저장할 데이터가 없습니다.')
            return None
        df = pd.DataFrame(data, columns=columns)
        file_name = f"환자조회결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df.to_excel(file_name, index=False, engine='openpyxl')
        print(f"✅ 조회 결과를 '{file_name}' 파일로 저장했습니다.")
        return None
    except Exception as e:
        print(f'❌ Excel 저장 실패. 오류: {e}')
        return None


def find_and_maximize_dentweb():
    """실행 중인 모든 창에서 '덴트웹'이 포함된 창을 찾아 자동으로 최대화"""
    windows = [title for title in gw.getAllTitles() if title.strip()]
    dentweb_title = None
    for title in windows:
        if '▶ 덴트웹' in title:
            dentweb_title = title
            print(f'✅ 덴트웹 창을 찾았습니다: {title}')
            break

    if not dentweb_title:
        print("❌ 실행 중인 창에서 '덴트웹'을 찾을 수 없습니다.")
        return False

    try:
        app = Application(backend='win32').connect(title_re=dentweb_title)
        window = app.top_window()
        print(f'✅ 덴트웹 창 연결 완료: {dentweb_title}')
        minimized_status = window.is_minimized()
        print(f'🔍 창 최소화 상태 확인: {minimized_status}')
        window.set_focus()
        if window.is_minimized():
            print('🔄 창이 최소화 상태입니다. 복원 후 최대화합니다.')
            window.restore()
            time.sleep(1)
        window.maximize()
        print('🚀 덴트웹 창을 전체화면으로 설정했습니다.')
        return (app, window)
    except Exception as e:
        print(f'❌ 덴트웹 창을 최대화하지 못했습니다. 오류: {e}')
        return (None, None)


def find_EditText(window):
    edit_fields = [e for e in window.descendants() if 'EDIT' in e.class_name().upper()]
    if not edit_fields:
        print('❌ 입력창(Edit)을 찾을 수 없습니다.')
        return None
    selected_edit = edit_fields[len(edit_fields) - 1]
    return selected_edit


def input_text_to_field(nameField, text='하창민'):
    '''지정된 입력 필드에 텍스트 입력'''
    if nameField:
        try:
            nameField.click_input()
            time.sleep(1)
            nameField.set_text(text)
            pyautogui.press('enter')
        except RuntimeError as e:
            print(f"마우스 클릭 방식이 막혀 직접 입력 방식으로 전환합니다: {e}")
            try:
                nameField.set_focus()
            except Exception:
                pass
            nameField.set_text(text)
            try:
                nameField.type_keys('{ENTER}')
            except Exception:
                pyautogui.press('enter')
        time.sleep(1)
        print(f"✅ '{text}' 입력 완료!")
        return None
    print('❌ 입력 필드가 존재하지 않습니다.')


def check_for_verification_window(app):
    """'본인확인 방법 선택' 창이 뜨면 취소 버튼 클릭"""
    try:
        time.sleep(1)
        for win in app.windows():
            if '본인확인 방법 선택' not in win.window_text():
                continue
            print(f"✅ '본인확인 방법 선택' 창 감지: {win.window_text()}")
            cancel_buttons = [e for e in win.descendants() if 'BUTTON' in e.class_name().upper()]
            for btn in cancel_buttons:
                if '취소' not in btn.window_text():
                    continue
                print(f"✅ '취소' 버튼 찾음: {btn.window_text()}")
                btn.click_input()
                time.sleep(1)
                print("✅ '본인확인 방법 선택' 창 닫음!")
                return True
            print("❌ '취소' 버튼을 찾을 수 없습니다.")
        print("✅ '본인확인 방법 선택' 창이 나타나지 않음.")
        return False
    except Exception as e:
        print(f"❌ '본인확인 방법 선택' 창 확인 실패. 오류: {e}")
        return None


def click_text_element(window, text='보험자격조회'):
    '''주어진 텍스트를 가진 UI 요소를 클릭'''
    try:
        time.sleep(1)
        elements = window.descendants()
        for element in elements:
            if text not in element.window_text():
                continue
            print(f"✅ '{text}' 버튼을 찾음: {element.window_text()}")
            element.click_input()
            print(f"✅ '{text}' 버튼 클릭 완료!")
            return True
        print(f"❌ '{text}' 버튼을 찾을 수 없습니다.")
        time.sleep(3)
        return False
    except Exception as e:
        print(f"❌ '{text}' 버튼 클릭 실패. 오류: {e}")
        return None


def check_insurance_result_window(app):
    """'ㅇㅇㅇ 님의 조회 결과' 창이 떴는지 확인"""
    try:
        time.sleep(1)

        # '진료 의뢰기관 기호 입력' 창 처리 - '취소' 버튼 클릭
        for win in app.windows():
            if '진료 의뢰기관 기호 입력' not in win.window_text():
                continue
            print(f"✅ '진료 의뢰기관 기호 입력' 창 감지: {win.window_text()}")
            cancel_buttons = [e for e in win.descendants() if 'BUTTON' in e.class_name().upper()]
            for btn in cancel_buttons:
                if '취소' not in btn.window_text():
                    continue
                print(f"✅ '진료 의뢰기관 기호 입력 취소' 버튼 찾음: {btn.window_text()}")
                btn.click_input()
                time.sleep(1)
                print("✅ '진료 의뢰기관 기호 입력' 창 닫음!")
            print("❌ '진료 의뢰기관 기호 입력 취소' 버튼을 찾을 수 없습니다.")

        # '진료 의뢰기관 기호 입력' 창 처리 - '아니오(N)' 버튼 클릭
        for win in app.windows():
            print(f"✅ '진료 의뢰기관 기호 입력' 창 감지: {win.window_text()}")
            cancel_buttons = [e for e in win.descendants() if 'BUTTON' in e.class_name().upper()]
            for btn in cancel_buttons:
                if '아니오(N)' not in btn.window_text():
                    continue
                print(f"✅ '진료 의뢰기관 기호 입력 취소' 버튼 찾음: {btn.window_text()}")
                btn.click_input()
                time.sleep(1)

        # '조회 결과' 창 처리 - '확인' 버튼 클릭
        for win in app.windows():
            if '조회 결과' not in win.window_text():
                continue
            print(f"✅ '조회 결과 선택' 창 감지: {win.window_text()}")
            cancel_buttons = [e for e in win.descendants() if 'BUTTON' in e.class_name().upper()]
            for btn in cancel_buttons:
                if '확인' not in btn.window_text():
                    continue
                print(f"✅ '확인' 버튼 찾음: {btn.window_text()}")
                btn.click_input()
                time.sleep(1)
                print("✅ '조회 결과 선택' 창 닫음!")
                return True
            print("❌ '확인' 버튼을 찾을 수 없습니다.")

        print("✅ '조회 결과 선택' 창이 나타나지 않음.")
        return False
    except Exception as e:
        print(f"❌ '조회 결과' 창 확인 실패. 오류: {e}")
        return None


def show_completion_popup():
    '''작업 완료 팝업 창 띄우기'''
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo('완료', '작업이 성공적으로 완료되었습니다.')
    root.destroy()


if __name__ == '__main__':
    (app, window) = find_and_maximize_dentweb()
    if window:
        nameField = find_EditText(window)
        (patient_list, columns) = get_patient_list()
        if not patient_list:
            print('❌ 조회할 환자가 없습니다.')
        else:
            print(f'✅ 조회할 환자 목록 ({len(patient_list)}명) 가져옴.')
        for patient in patient_list:
            (patient_id, chart_no, patient_type, patient_name, birth_date, guardian_phone, phone, last_visit, address, memo, elapsed_months) = patient
            print(f"\n🔍 '{patient_name}({chart_no})' 조회 시작 ({elapsed_months}개월 경과)")
            search_text = f'{patient_name}({chart_no})'
            input_text_to_field(nameField, search_text)
            check_for_verification_window(app)
            click_text_element(window, '보험자격조회')
            check_for_verification_window(app)
            check_insurance_result_window(app)
        time.sleep(1)
        (latest_data, latest_columns) = get_patient_list()
        save_to_excel(latest_data, latest_columns)
        show_completion_popup()
        exit()
