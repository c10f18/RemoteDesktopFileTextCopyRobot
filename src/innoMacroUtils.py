import time
import pyautogui
import clipboard
import os

# 특정 지점 클릭 후 컨트롤c (원격 데스크톱은 핫키조합인 hotkey() func가 적용되지 않는 경우도 있어서 key press 홀딩 방식 사용)
def clickAndCopyAll(x, y):
    # print('clickAndCopyAll()')
    # 마우스 커서 이동
    pyautogui.moveTo(x, y)
    # 클릭
    pyautogui.click()
    # 컨트롤 a
    # pyautogui.hotkey('ctrl', 'a')
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')
    # 컨트롤 c
    # pyautogui.hotkey('ctrl', 'c')
    pyautogui.keyDown('ctrl')
    pyautogui.press('c')
    pyautogui.keyUp('ctrl')
    # print('지점 클릭 후 전체 선택 및 복사')

# 컨트롤c
def CopyAll():
    # print('clickAndCopyAll()')
    # 컨트롤 a
    # pyautogui.hotkey('ctrl', 'a')
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')
    # 컨트롤 c
    # pyautogui.hotkey('ctrl', 'c')
    pyautogui.keyDown('ctrl')
    pyautogui.press('c')
    pyautogui.keyUp('ctrl')
    # print('지점 클릭 후 전체 선택 및 복사')

# 특정 지점 클릭 후 컨트롤a 컨트롤v 컨트롤s (원격 데스크톱은 핫키조합인 hotkey() func가 적용되지 않는 경우도 있어서 key press 홀딩 방식 사용)
# save할 파일이 이미 만들어져있고 전체 덮어쓰기 한다는 가정
def clickAndSelectAllAndPaste(x, y):
    # print('clickAndSelectAllAndPaste()')
    pyautogui.moveTo(x, y)
    pyautogui.click()
    # 컨트롤 a
    pyautogui.keyDown('ctrl')
    pyautogui.press('a')
    pyautogui.keyUp('ctrl')
    # 컨트롤 v
    # pyautogui.hotkey('ctrl', 'v')
    pyautogui.keyDown('ctrl')
    pyautogui.press('v')
    pyautogui.keyUp('ctrl')
    # print('지점 클릭 후 전체 선택 및 붙여넣기')

def clickAndPaste(x, y):
    # print('clickAndPaste()')
    pyautogui.moveTo(x, y)
    pyautogui.click()
    # 컨트롤 v
    # pyautogui.hotkey('ctrl', 'v')
    pyautogui.keyDown('ctrl')
    pyautogui.press('v')
    pyautogui.keyUp('ctrl')
    # print('지점 클릭 후 붙여넣기')

# 큰 창으로 크기 변경 (따로 클릭 안들어감. 현재 포커스된 창 기준)
def maximizeWindow():
    # Alt + space + X
    # pyautogui.hotkey('alt', 'space', 'x')
    pyautogui.keyDown('alt')
    pyautogui.keyDown('space')
    pyautogui.press('x')
    pyautogui.keyUp('alt')
    pyautogui.keyUp('space')
    # print('현재 포커스 된 창 최대화')

def openCmdConsole():
    # window + x > c
    pyautogui.keyDown('win')
    pyautogui.press('x')
    pyautogui.keyUp('win')
    pyautogui.press('c')
    # print('커맨드 창 오픈')

def closeWindow():
    # window + x > c
    pyautogui.keyDown('alt')
    pyautogui.press('f4')
    pyautogui.keyUp('alt')
    # print('현재 포커스 된 창 끄기')

def fileOpenAndSaveClipboard(filepath):
    # 파일 생성, 폴더 없으면 폴더도 생성
    if not os.path.exists(filepath):
        #디렉토리까지 먼저 만든다
        if not os.path.exists(os.path.dirname(filepath)):
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
        if not os.path.isdir(filepath):
            f = open(filepath, 'w', encoding='utf-8')
            f.write(clipboard.paste())
            f.close()
