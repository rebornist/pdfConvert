# 파이썬 매크로 기능을 이용한 PDF 저장

## pyautogui 설치
- pipenv install pyautogui

## 기본 기능 - 마우스

  1. 좌표 객체 얻기 
  ```py
  position = pyautogui.position()
  ```

  1. 화면 전체 크기 확인하기
  ```py
  print(pyautogui.size())
  ```

  3. x, y 좌표
  ```py
  print(position.x)
  print(position.y)
  ```

  4. 마우스 이동 (x 좌표, y 좌표)
  ```py
  pyautogui.moveTo(500, 500)
  ```

  5. 마우스 이동 (x 좌표, y 좌표 2초간)
  ```py
  pyautogui.moveTo(100, 100, 2)  
  ```

  6. 마우스 이동 ( 현재위치에서 )
  ```py
  pyautogui.moveRel(200, 300, 2)
  ```

  7. 마우스 클릭
  ```py
  pyautogui.click()
  ```

  8. 2초 간격으로 2번 클릭
  ```py
  pyautogui.click(clicks= 2, interval=2)
  ```

  9. 더블 클릭
  ```py
  pyautogui.doubleClick()
  ```

  10. 오른쪽 클릭
  ```py
  pyautogui.click(button='right')
  ```

  11. 스크롤하기 
  ```py
  pyautogui.scroll(10)
  ```

  12. 드래그하기
  ```py
  pyautogui.drag(0, 300, 1, button='left')
  ```


## 기본 기능 - 키보드

1. write 함수
```py
pyautogui.write('hello world!') # 괄호 안의 문자를 타이핑 합니다.
pyautogui.write('hello world!', interval=0.25) # 각 문자를 0.25마다 타이핑합니다
```

2. pyperclip 모듈 설치

```py
pipenv install pyperclip

import pyperclip

pyperclip.copy("안녕하세요") # 클립보드에 텍스트를 복사합니다. 

pyautogui.hotkey('ctrl', 'v') # 붙여넣기 (hotkey 설명은 아래에 있습니다.)
```

3. press(), keyDown(), keyUp() 함수

```py
pyautogui.press('shift') # shift 키를 누릅니다.
pyautogui.press('ctrl') # ctrl 키를 누릅니다.

pyautogui.keyDown('ctrl') # ctrl 키를 누른 상태를 유지합니다.
pyautogui.press('c') # c key를 입력합니다. 
pyautogui.keyUp('ctrl') # ctrl 키를 뗍니다.

pyautogui.press(['left', 'left', 'left']) # 왼쪽 방향키를 세번 입력합니다.
pyautogui.press('left', presses=3) # 왼쪽 방향키를 세번 입력합니다. 
pyautogui.press('enter', presses=3, interval=3) # enter 키를 3초에 한번씩 세번 입력합니다. 
```

4. hotkey() 함수

```py
pyautogui.hotkey('ctrl', 'c') # ctrl + c 키를 입력합니다. 
```

5. 키보드 키의 명칭 리스트
  
```py
['\t', '\n', '\r', ' ', '!', '"', '#', '$', '%', '&', "'", '(',
')', '*', '+', ',', '-', '.', '/', '0', '1', '2', '3', '4', '5', '6', '7',
'8', '9', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`',
'a', 'b', 'c', 'd', 'e','f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o',
'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '{', '|', '}', '~',
'accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace',
'browserback', 'browserfavorites', 'browserforward', 'browserhome',
'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear',
'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete',
'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10',
'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20',
'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9',
'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja',
'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail',
'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack',
'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6',
'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn',
'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn',
'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator',
'shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop', 'subtract', 'tab',
'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen',
'command', 'option', 'optionleft', 'optionright']
```