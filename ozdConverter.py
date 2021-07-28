import pyautogui
import openpyxl
import os
import time
import sys
import win32com.shell.shell as shell

# # 관리자 권한 부여
# if sys.argv[-1] != 'asadmin':
#     script = os.path.abspath(sys.argv[0])
#     params = ' '.join([script] + sys.argv[1:] + ['asadmin'])
#     shell.ShellExecuteEx(
#         lpVerb='runas', lpFile=sys.executable, lpParameters=params)
#     sys.exit(0)


class OzdConverter:
    def __init__(self, input_file=''):
        self.input_file = input_file
        self.data = []
        self.results = []

    def open_viewer(self):
        BASE_DIR = 'C:/Program Files (x86)/Forcs/OZ Family/kpicviewer/ozviewer'
        os.chdir(BASE_DIR)
        os.system('ozcviewer.exe')

    def read_excel(self):
        wb = openpyxl.load_workbook(self.input_file)
        ws = wb.active

        for data in ws.rows:
            src = data[1].value
            if not os.path.exists(src):
                print(f'{src} is not exist')
                continue

            dst = data[2].value
            if os.path.exists(dst):
                print(f'{dst} is already exist')
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

        # 오즈 리포트 뷰어로 이동 후 프로그램 선택
        x = 936
        y = 208
        pyautogui.moveTo(x, y, 1)
        time.sleep(1)
        # pyautogui.hotkey('alt', 'tab')
        pyautogui.click()

        # 대기시간
        time.sleep(1)

        # 파일열기 매뉴 실행
        pyautogui.hotkey('ctrl', 'o')
        # x = 683
        # y = 227
        # pyautogui.moveTo(x, y, 1)
        # pyautogui.hotkey('alt', 'tab')

        # 대기시간
        time.sleep(5)

        # 파일 경로 입력
        pyautogui.write(src)
        # x = 1065
        # y = 697
        # pyautogui.moveTo(x, y, 1)

        # 대기시간
        time.sleep(1)

        # 입력 버튼 실행
        pyautogui.hotkey('alt', 'o')
        # pyautogui.hotkey('alt', 'tab')
        # pyautogui.click()

        # 대기시간
        time.sleep(5)

        # 인쇄 매뉴 실행
        pyautogui.hotkey('ctrl', 'p')

        # 대기시간
        time.sleep(5)

        # 엔터 키 입력
        pyautogui.press('enter')

        # 대기시간
        time.sleep(2)

        # 파일 경로 입력
        pyautogui.write(dst)

        # 대기시간
        time.sleep(1)

        # 저장 버튼 실행
        pyautogui.hotkey('alt', 's')

        # 대기시간
        time.sleep(1)


if __name__ == '__main__':

    path = sys.argv[1]
    o = OzdConverter(path)

    o.read_excel()

    if len(o.data) > 0:
        # 오즈 리포트 뷰어 실행
        o.open_viewer()

        # 대기 시간
        time.sleep(5)

        for data in o.data:
            src = data[1]
            if not os.path.exists(src):
                print(f'{src} is not exist')
                o.results.append([data[0], src, dst, 'N', 'src is not exist'])
                continue

            dst = data[2]
            if os.path.exists(dst):
                print(f'{src} is already exist')
                o.results.append(
                    [data[0], src, dst, 'N', 'dst is already exist'])
                continue

            o.auto_event(src, dst)
            o.results.append([data[0], src, dst, 'Y', ''])
        else:
            o.write_excel()
