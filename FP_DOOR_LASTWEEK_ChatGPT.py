# pip install pyautogui
import pyautogui as gui
import time
from pathlib import Path
from datetime import date, timedelta
import keyboard
import sys

# ----- 기본 설정 -----
gui.FAILSAFE = True                 # 마우스를 화면 좌상단으로 이동하면 즉시 중단
gui.PAUSE = 0.2                     # 각 동작 사이 기본 간격
TYPE_INTERVAL = 0.02                # 타이핑 간 간격

# ----- 설정값 -----
CRED_PATH = Path(r"C:\Temp\sap_id.txt")
SAP_EXE   = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
EXPORT_DIR = r"C:\Temp\FP_DOOR_SCHEDULE\\"
OUTLOOK_EXE = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"

# 좌표 (픽셀). 필요시 조정
POS_PROJECT_MENU = (1700, 820)  # 02
POS_ID_FIELD     = (200, 205)   # 06, 07 사번 입력
POS_PW_FIELD     = (200, 228)   # 08, 09 비번 입력
POS_TOPLEFT_CELL = (100, 55)    # 10
POS_PRODUC_SCH   = (41, 205)    # 12 생산계획일 선택
POS_FROM_FIELD     = (285, 205)   # 13 생산계획 시작일
POS_EXECUTE_BTN  = (185, 337)   # 16
POS_WAIT_PIXEL   = (25, 1135)   # 18 (색상 확인)
POS_MENU_CLICK_1 = (366, 130)   # 20
POS_MENU_CLICK_2 = (430, 198)   # 21

TARGET_RGB = (251, 179, 7)      # 19에서 감시할 색상

# 안전/지연 기본값
gui.FAILSAFE = True           # 마우스를 화면 좌상단으로 이동하면 중단
gui.PAUSE = 0.5               # 각 동작 사이 기본 간격
TYPE_INTERVAL = 0.02          # 타이핑 간 간격

def read_credentials(path: Path):
    if not path.exists():
        print(f"자격파일 없음: {path}")
        sys.exit(1)
    lines = [line.rstrip("\r\n") for line in path.read_text(encoding="utf-8").splitlines()]
    if len(lines) < 2:
        print("sap_id.txt 형식 오류: 첫줄 ID, 둘째줄 비밀번호가 있어야 합니다.")
        sys.exit(1)
    return lines[0], lines[1]

def last_week_sat_sun(today: date):
    """
    '지난주'를 ISO 기준(월~일)으로 계산:
    - 이번 주 월요일을 구한 뒤 +7일 = 다음주 월요일
    - 지난주 토요일 = 이번주 월요일 - 8일
    - 지난주 일요일 = 이번주 월요일 - 1일
    """
    this_monday = today - timedelta(days=today.weekday())  # 0=월
    # next_monday = this_monday + timedelta(days=7)
    last_sat = this_monday + timedelta(days=-8)
    last_sun = this_monday + timedelta(days=-2)
    return last_sat, last_sun

def yyyymmdd(d: date) -> str:
    return d.strftime("%Y%m%d")

def yyyy_mm_dd(d: date) -> str:
    return d.strftime("%Y.%m.%d")

def _same_color(px, rgb, tol=0):
    # tol 픽셀 허용오차 (기본 0: 완전 일치)
    return all(abs(int(px[i]) - int(rgb[i])) <= tol for i in range(3))

# SAP 실행중커서 색상 확인
def wait_until_color_gone(x: int, y: int, rgb: tuple, check_interval=1.0, tol=0, timeout=300):
    """
    (x,y) 픽셀이 'rgb' 인 상태가 끝날 때까지 대기.
    rgb가 더 이상 아니면 1초 안정화 대기 후 다음 단계로 진행.
    - check_interval: 확인 주기(초)
    - tol: 색상 허용오차(0이면 완전일치 기준)
    - timeout: 최대 대기(초). 초과 시 경고 후 진행
    """
    start = time.time()
    while True:
        try:
            px = gui.pixel(x, y)
        except Exception:
            # 스크린샷 권한 문제 등이 있으면 진행
            print("경고: 픽셀 읽기 실패 → 다음 단계로 진행합니다.")
            break

        if not _same_color(px, rgb, tol=tol):
            # 요청: 해당 색상이 사라지면 진행 (안정화 1초 대기)
            time.sleep(1.0)
            break

        if time.time() - start > timeout:
            print("경고: 색상 사라짐 대기 타임아웃 → 다음 단계로 진행합니다.")
            break

        time.sleep(check_interval)

# def wait_for_pixel_color(x: int, y: int, rgb: tuple, check_interval=1.0, timeout=300):
#     """
#     주기적으로 (x,y) 픽셀 색상을 확인하여 지정 RGB가 될 때까지 대기.
#     timeout(초) 초과 시 종료(오류 아님, 상황에 맞게 조정).
#     """
#     start = time.time()
#     while True:
#         try:
#             px = gui.pixel(x, y)
#         except Exception as e:
#             # 일부 환경에서 스크린샷 권한 문제 발생 가능
#             px = (-1, -1, -1)
#         if px == rgb:
#             # 요구사항: '색상이 (251,179,7)일 때 대기' → 해당 색상에 도달하면 1초 더 대기 후 진행
#             time.sleep(1.0)
#             break
#         if time.time() - start > timeout:
#             print("경고: 픽셀 색상 감시 타임아웃, 다음 단계로 진행합니다.")
#             break
#         time.sleep(check_interval)

# ── 메인 시퀀스 ───────────────────────────────────────────────────────────────
def main():
    emp_id, emp_pw = read_credentials(CRED_PATH)

    # 안전하게 준비할 시간
    for i in range(3, 0, -1):
        print(f"시작 {i}...")
        time.sleep(1)

    # 00 WIN키 누른 채로 M 입력
    gui.hotkey("win", "m"); time.sleep(1)

    # 01 WIN키 누른 채로 P 입력
    gui.hotkey("win", "p"); time.sleep(2)

    # 02 1700x820 위치 클릭 후 1초 대기
    gui.click(*POS_PROJECT_MENU); time.sleep(2)

    # 00-2 WIN키 누른 채로 M 입력
    gui.hotkey("win", "m"); time.sleep(1)

    # 03 win키 누른 채로 R 입력
    gui.hotkey("win", "r")

    # 04 saplogon.exe 경로 입력 → 1초 대기 → ENTER
    gui.write(SAP_EXE, interval=TYPE_INTERVAL)
    time.sleep(1)
    gui.press("enter")
    time.sleep(5)

    # 05 ALT키 누른 채로 L 입력
    gui.hotkey("alt", "l")
    time.sleep(3)

    # 06 200x205 위치 클릭
    gui.click(*POS_ID_FIELD)
    time.sleep(1)

    # 07 사번 입력(첫줄)
    gui.write(emp_id, interval=TYPE_INTERVAL)
    time.sleep(1)

    # 08 200x228 위치 클릭
    gui.click(*POS_PW_FIELD)
    time.sleep(1)

    # 09 비번 입력(둘째줄)
    gui.write(emp_pw, interval=TYPE_INTERVAL)
    time.sleep(3)

    # 10 100x55 위치 클릭
    gui.click(*POS_TOPLEFT_CELL)
    time.sleep(1)

    # 11 ZKPPR0092 입력 후 1초 대기 후 ENTER
    gui.write("ZKPPR0092", interval=TYPE_INTERVAL)
    time.sleep(1)
    gui.press("enter")
    time.sleep(3)

    # 12 지난주 일요일 날짜 → abcdefgh(YYYYMMDD)
    # 13 지난주 토요일 날짜 → ijklmnop(YYYYMMDD)
    last_sat, last_sun = last_week_sat_sun(date.today())
    abcdefgh = yyyymmdd(last_sat)     # 토요일
    ijklmnop = yyyymmdd(last_sun)     # 일요일

    # 점 표기용
    abcd_dot_ef_dot_gh   = yyyy_mm_dd(last_sat)  # YYYY.MM.DD
    ijkl_dot_mn_dot_op   = yyyy_mm_dd(last_sun)  # YYYY.MM.DD

    # 파일명 일부(YYMMDD) → cdefgh, klmnop
    cdefgh  = abcdefgh[2:]   # 토요일 YYMMDD
    klmnop  = ijklmnop[2:]   # 일요일 YYMMDD

    # 12(재) 41x205 위치 클릭 후 1초 대기
    gui.click(*POS_PRODUC_SCH)
    # gui.press("home")

    # 13(재) 315x205 위치 클릭 후 1초 대기
    gui.click(*POS_FROM_FIELD)
    keyboard.press_and_release("home")

    # 14 abcd.ef.gh 입력 후 1초 대기  (→ YYYY.MM.DD)
    gui.write(abcd_dot_ef_dot_gh, interval=TYPE_INTERVAL); time.sleep(1)
    keyboard.press_and_release("tab")

    # 15 ijkl.mn.op 입력 후 1초 대기  (→ YYYY.MM.DD)
    gui.write(ijkl_dot_mn_dot_op, interval=TYPE_INTERVAL); time.sleep(1)

    # 16 185x337 클릭
    gui.click(*POS_EXECUTE_BTN)

    # 17 F8키 입력
    gui.press("f8")

    # 18 25x1135 위치 색상 인식하여 1초마다 확인
    # 19 색상 RGB값이 (251,179,7)일 때 대기(= 그 색이 될 때까지 대기)
    # wait_for_pixel_color(POS_WAIT_PIXEL[0], POS_WAIT_PIXEL[1], TARGET_RGB, check_interval=1.0)
    # 19 색상 RGB값이 (251,179,7)일 때 사라질 때까지 대기
    wait_until_color_gone(POS_WAIT_PIXEL[0], POS_WAIT_PIXEL[1], TARGET_RGB, check_interval=1.0, tol=3, timeout=300)

    # 20 366x130 위치 클릭 후 1초 대기
    time.sleep(5)
    gui.click(*POS_MENU_CLICK_1); time.sleep(5)

    # # 21 430x198 위치 클릭 후 1초 대기
    # gui.click(*POS_MENU_CLICK_2); time.sleep(2)
    keyboard.press_and_release("L"); time.sleep(1)

    # # 22 아래 화살표 1번 누르기
    keyboard.press_and_release("down"); time.sleep(1)
    # # gui.press("down")

    # # 23 TAB키 입력 후 1초 대기 후 ENTER키 입력
    keyboard.press_and_release("enter"); time.sleep(3)
    # keyboard.press_and_release("tab")
    # time.sleep(1)
    # keyboard.press_and_release("enter")
    # # gui.press("tab"); time.sleep(1); gui.press("enter")

    # # 24 SHIFT키 누른 채로 TAB키 입력(역탭)
    gui.hotkey("shift", "tab"); time.sleep(1)

    # # 25 C:\Temp\FP_DOOR_SCHEDULE\ 입력
    gui.hotkey("Ctrl", "a"); time.sleep(1)
    gui.write(EXPORT_DIR, interval=TYPE_INTERVAL)
    keyboard.press_and_release("backspace"); time.sleep(1)

    # # 26 TAB 키 입력 후 1초 대기
    keyboard.press_and_release("tab")
    time.sleep(1)
    # # gui.press("tab"); time.sleep(1)

    # # 27 "생산계획_cdefgh-klmnop.XLS" 입력 (cdefgh/klmnop = YYMMDD)
    filename = f"Producion-Plan_{cdefgh}-{klmnop}.XLS"
    gui.hotkey("Ctrl", "a"); time.sleep(1)
    gui.write(filename, interval=TYPE_INTERVAL); time.sleep(1)

    # # 28 TAB → 0.5초 대기 → TAB → 1초 대기 → ENTER
    keyboard.press_and_release("tab"); time.sleep(0.5)
    keyboard.press_and_release("tab"); time.sleep(1)
    # gui.press("tab"); time.sleep(0.5)
    # gui.press("tab"); time.sleep(1)
    gui.press("enter")

    # # 29 WIN키 누른 채로 R 입력
    # gui.hotkey("win", "r")

    # # 30 OUTLOOK.EXE 경로 입력 후 1초 대기 후 ENTER
    # gui.write(OUTLOOK_EXE, interval=TYPE_INTERVAL)
    # time.sleep(1)
    # gui.press("enter")

    print("모든 단계 완료")

if __name__ == "__main__":
    main()
