import pyautogui
import openpyxl
import os
import time
import sys
# import win32com.shell.shell as shell
import win32com.client as win32

# # 관리자 권한 부여
# if sys.argv[-1] != 'asadmin':
#     script = os.path.abspath(sys.argv[0])
#     params = ' '.join([script] + sys.argv[1:] + ['asadmin'])
#     shell.ShellExecuteEx(
#         lpVerb='runas', lpFile=sys.executable, lpParameters=params)
#     sys.exit(0)

# BASE = 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'


class ExcelConverter:
    def __init__(self, input_file=''):
        self.input_file = input_file
        self.data = []
        self.results = []

    def open_viewer(self, src):

        # 엑셀 애플리케이션 준비
        # self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel = win32.Dispatch("Excel.Application")
        # 윈도우 화면에 띄울지, 아니면 백그라운드에서 돌릴지 선택
        # 테스트를 위해 True로 해서 직접 엑셀이 돌아가는 걸 눈으로 확인하자
        # self.excel.DisplayAlerts = False
        # self.excel.EnableEvents = False
        self.excel.AskToUpdateLinks = False
        self.excel.ScreenUpdating = False
        self.excel.Visible = True

        try:
            self.excel.Workbooks.Open(src)
        except Exception as e:
            print("excel open error:", e, src)

    def close_viewer(self, wb):
        try:
            self.excel.Quit()
        except Exception as e:
            print("excel close error:", e)

    def read_excel(self):
        wb = openpyxl.load_workbook(self.input_file)
        ws = wb.active

        for data in ws.rows:
            src = data[1].value
            dst = data[2].value

            if not os.path.exists(src):
                print(f'{src} is not exist')
                o.results.append(
                    [data[0], src, dst, 'N', 'src is already exist'])
                continue

            if os.path.exists(dst):
                print(f'{dst} is already exist')
                o.results.append(
                    [data[0], src, dst, 'N', 'dst is already exist'])
                continue

            self.data.append([data[0].value, src, dst])

        wb.close()

    def write_excel(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        for result in self.results:
            ws.append(result)

        wb.save(f'./result_{int(time.time())}.xlsx')
        wb.close()

    def auto_event(self, src, dst):

        # 엑셀 뷰어로 이동 후 프로그램 선택
        x = 936
        y = 208
        pyautogui.moveTo(x, y, 1)
        time.sleep(1)
        # pyautogui.hotkey('alt', 'tab')
        pyautogui.click()

        if self.excel.DisplayAlerts or self.excel.EnableEvents:
            time.sleep(1)
            pyautogui.press('enter')  # enter key를 입력합니다.

        # 대기시간
        time.sleep(2)

        # 인쇄 매뉴 실행
        pyautogui.hotkey('ctrl', 'p')

        # 대기시간
        time.sleep(2)

        # 인쇄 옵션으로 이동 후 매뉴 선택
        x = 549
        y = 435
        pyautogui.moveTo(x, y, 1)
        pyautogui.click()

        # 대기시간
        time.sleep(2)

        # 전체 시트 인쇄 옵션 선택
        x = 549
        y = 526
        pyautogui.moveTo(x, y, 1)
        pyautogui.click()

        # 대기시간
        time.sleep(2)

        # 출력 선택
        x = 435
        y = 216
        pyautogui.moveTo(x, y, 1)
        pyautogui.click()

        # 대기시간
        time.sleep(2)

        # 파일 경로 입력
        pyautogui.write(dst)

        # 대기시간
        time.sleep(1)

        # 저장 버튼 실행
        pyautogui.hotkey('alt', 's')

        # 대기시간
        time.sleep(10)


if __name__ == '__main__':

    # while True:
    #     print(pyautogui.position())

    path = sys.argv[1]
    o = ExcelConverter(path)

    # 데이터 모으기
    o.read_excel()

    if len(o.data) > 0:
        # 존재하는 데이터 분할
        for data in o.data:
            src = data[1]
            dst = data[2]
            try:
                # 엑셀 뷰어 실행
                wb = o.open_viewer(src)

                # 대기시간
                time.sleep(3)

                # pdf 변환 이벤트 실행
                o.auto_event(src, dst)

                # 성공 메시지 전송
                o.results.append([data[0], src, dst, 'Y', ''])
            except Exception as e:
                # 실패 메시지 출력
                print(e)

                # 실패 메시지 전송
                o.results.append([data[0], src, dst, 'N', e])

            # 대기시간
            time.sleep(3)

            # 엑셀 뷰어 닫기
            o.close_viewer(wb)

        else:

            # 결과 추출
            o.write_excel()
