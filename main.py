import time
import tkinter
from tkinter import *
import tkinter.ttk as ttk
import clipboard
import pyautogui
from src import innoMacroUtils as macro
import os
import configparser
import pickle
import math

# 마우스 포인터 정보
# pyautogui.mouseInfo()
# 대기시간 일괄 조정

##########################################
# config 파일 읽음
# properties = configparser.ConfigParser()
# properties.read('resource/settings.ini')
# config = properties["CONFIG"]
# source_root = config['sourceRootDir']
# target_root = config['targetRootDir']
# sourceX = config['sourceX']
# sourceY = config['sourceY']
# wait_sec = config['wait_sec']
##########################################
# gui
root = Tk()
root.title("Remote Desktop File Copy Robot")
root.geometry("1000x730+2000+50")
##################
frame_xy = LabelFrame(root, text="Remote Desktop Window 좌표")

source_info_lab1 = Label(frame_xy, text = "Remote Desktop Window 내의 좌표를 입력해주세요\n (화면의 좌측 상단이 (0, 0) 입니다.\n해당 좌표를 클릭해서 Remote Desktop Window을 포커스 합니다.)")
source_info_lab2 = Label(frame_xy)
source_info_lab3 = Label(frame_xy, text = "X : ")
source_info_lab4 = Label(frame_xy, text = "Y : ")

source_x_entry = Entry(frame_xy, width=10)
source_x_entry.insert(END, "500")
source_y_entry = Entry(frame_xy, width=10)
source_y_entry.insert(END, "500")
##################
frame_wait_sec = LabelFrame(root, text="명령 프롬프트 실행 대기 시간")

wait_sec_lab1 = Label(frame_wait_sec, text = "명령 프롬프트 실행이 늦는 경우 대기 시간 조정")
wait_sec_lab2 = Label(frame_wait_sec)
wait_sec_lab3 = Label(frame_wait_sec, text = "wait sec : ")

wait_sec_entry = Entry(frame_wait_sec, width=10)
wait_sec_entry.insert(END, "1")
##################
frame_folder_info = LabelFrame(root, text="파일 경로")

folder_info_lab1 = Label(frame_folder_info, text = "파일 경로를 입력하세요")
folder_info_lab2 = Label(frame_folder_info, text = "Remote Desktop 에서의 Root Folder : ")
folder_info_lab3 = Label(frame_folder_info, text = "Local에서 저장할 Root Folder : ")

folder_info_entry1 = Entry(frame_folder_info, width=30)
folder_info_entry1.insert(0, r"C:\Users\vm\Documents")
folder_info_entry2 = Entry(frame_folder_info, width=30)
folder_info_entry2.insert(0, r"D:\macroRoot")
##################
frame_list = LabelFrame(root, text="옮길 파일 리스트 추출은 해당 Remote Desktop의 root 위치에서 명령 프롬프트에 'dir /b /s > {파일명}.txt' 실행 > 텍스트 파일로 리스트 드랍됨 > 복사")
frame_source = LabelFrame(frame_list, text="옮길 파일 리스트")
frame_complete = LabelFrame(frame_list, text="성공")
frame_fail = LabelFrame(frame_list, text="실패")
frame_log = LabelFrame(frame_list, text="Log")
source_list_txt = Text(frame_source, width=30, height=30)
complete_list_txt = Text(frame_complete, width=70, height=13 )
fail_list_txt = Text(frame_fail, width=70, height=13)
log_list_txt = Text(frame_log, width=30, height=30)

##################
p_var = tkinter.DoubleVar()
progressbar = ttk.Progressbar(root, length=500, maximum=100, mode="determinate", variable=p_var)

def progressUpdate(i):
    global p_var
    p_var.set(i)
    progressbar.update()

def stop():
    isStop = True

# pyautogui.mouseInfo()
##########################################
# 파일 원본 리스트 추출은 cmd 창에서 작업하고자 하는 소스 루트 폴더로 접근한 후 dir /b /s > {파일명}.txt 실행하면 추출된다
# 파일 원본 리스트를 한줄씩 조회한다
# f = open("resource/sourceFileList.txt", "r", encoding='utf-8')
logList = []
completeList = []
failList = []
wait_sec = 0.7
isStop = False
def startCopy():
    sourceX = int(source_x_entry.get())
    sourceY = int(source_y_entry.get())
    wait_sec = float(wait_sec_entry.get())
    global isStop
    progressUpdate(0)

    source_root = folder_info_entry1.get()
    target_root = folder_info_entry2.get()
    print(source_root)

    complete_list_txt.delete("1.0", tkinter.END)
    fail_list_txt.delete("1.0", tkinter.END)
    log_list_txt.delete("1.0", tkinter.END)

    souceFileListTextbox = source_list_txt.get("1.0", "end")
    souceFileList = souceFileListTextbox.splitlines()
    # souceFileList = f.readlines()
    total = len(souceFileList)
    # f.close()
    ##########################################
    ## 원격에서 커맨드 창 준비 > 최대화
    pyautogui.click(sourceX, sourceY)
    macro.openCmdConsole()
    time.sleep(wait_sec)
    macro.maximizeWindow()
    ##########################################
    # 파일 리스트 순회
    for i in range(len(souceFileList)):
        if not isStop:
            source_path = souceFileList[i].strip()
            target_path = source_path.replace(source_root, target_root)
            try:
                # 디렉토리가 아니고 파일이면 복사 저장. 원격 환경 판단이라 경로에 .이 포함되었는지로 구분
                if '.' in source_path.replace('.polarion', 'tmp') and len(source_path) > 0:
                    ##########################################
                    # 콘솔창으로 파일 오픈
                    # pyautogui.rightClick()
                    # pyautogui.press('enter')
                    time.sleep(wait_sec)
                    clipboard.copy('notepad.exe "' + source_path + '"')
                    pyautogui.keyDown('ctrl')
                    pyautogui.press('v')
                    pyautogui.keyUp('ctrl')
                    pyautogui.press('enter')
                    # print('콘솔창에 입력 > 파일 열림')
                    # 파일 열림
                    time.sleep(wait_sec)
                    macro.maximizeWindow()
                    macro.CopyAll()
                    # 콘솔창이 아니면 창 닫음. 아이콘 색상으로 콘솔창인지 확인
                    if not pyautogui.pixelMatchesColor(10, 10, (0, 0, 0)):
                        macro.closeWindow()
                    ##########################################
                    # 결과 출력 및 저장
                    progressUpdate(round(i / total * 100, 2))
                    print(str(round(i / total * 100)) + "%, (" + str(i) + " / " + str(
                        total) + "): " + source_path.strip() + " >> " + target_path)
                    logList.append(source_path.strip() + " >> " + target_path)
                    macro.fileOpenAndSaveClipboard(target_path)
                    completeList.append(source_path)
                    complete_list_txt.insert("end", source_path + "\n")
                    log_list_txt.insert("end", source_path.strip() + " >> " + target_path + "\n")
                    # 성공 저장 completeList
                    with open("resource/completeList.txt", "a", encoding='utf-8') as completeFile:
                        completeFile.write(source_path + "\n")
                else:  # 디렉토리 이면 폴더 생성
                    os.makedirs(target_path, exist_ok=True)
                    complete_list_txt.insert("end", source_path + "\n")
            except Exception as e:
                # logList 저장
                progressUpdate(round(i / total * 100, 2))
                errStr = "err!!(" + str(round(i / total * 100)) + "%, " + str(i) + " / " + str(
                    total) + "): " + source_path + " >> " + target_path + "\n "
                print(errStr, e)
                logList.append("err!!: " + source_path + " >> " + target_path)
                log_list_txt.insert("end", "err!!: " + source_path + " >> " + target_path + "\n")
                with open("resource/log.txt", "wb") as logFile:
                    logFile.write("err!!: " + source_path + " >> " + target_path + "\n")
                failList.append(source_path)
                fail_list_txt.insert("end", source_path + "\n")
                # 실패 저장 failList
                with open("resource/failList.txt", "a", encoding='utf-8') as failFile:
                    failFile.write(source_path + "\n")
        else:
            isStop = False
            break
    progressUpdate(100)
    # 마지막에 콘솔창 닫음. 아이콘 색상으로 콘솔창인지 확인
    if pyautogui.pixelMatchesColor(10, 10, (0, 0, 0)):
        macro.closeWindow()

        print("작업 종료")


startBtn = Button(root, text="Start", padx=3, pady=3, bg="yellow", command=startCopy)
stopBtn = Button(root, text="Stop", padx=3, pady=3, bg="yellow", command=stop)

frame_xy.grid(row=0, column=0, sticky="w")
source_info_lab1.grid(row=0, column=1, sticky="w")
source_info_lab2.grid(row=1, column=1, sticky="w")
source_info_lab3.grid(row=1, column=0, sticky="e")
source_x_entry.grid(row=1, column=1, sticky="w")
source_info_lab4.grid(row=2, column=0, sticky="e")
source_y_entry.grid(row=2, column=1, sticky="w")

frame_folder_info.grid(row=1, column=0, sticky="w")
folder_info_lab1.grid(row=0, column=0, sticky="w")
folder_info_lab2.grid(row=1, column=0, sticky="e")
folder_info_entry1.grid(row=1, column=1, sticky="w")
folder_info_lab3.grid(row=2, column=0, sticky="e")
folder_info_entry2.grid(row=2, column=1, sticky="w")

frame_wait_sec.grid(row=2, column=0, sticky="w")
wait_sec_lab1.grid(row=0, column=0, sticky="w")
wait_sec_lab3.grid(row=1, column=0, sticky="e")
wait_sec_entry.grid(row=1, column=1, sticky="w")

frame_list.grid(row=3, column=0, sticky="w")
frame_source.grid(row=0, column=0, rowspan=2, sticky="w")
frame_complete.grid(row=0, column=1, sticky="w")
frame_fail.grid(row=1, column=1, sticky="w")
frame_log.grid(row=0, column=2, rowspan=2, sticky="w")
source_list_txt.pack()
complete_list_txt.pack()
fail_list_txt.pack()
log_list_txt.pack()

progressbar.grid(row=4, column=0)
startBtn.grid(row=4, column=0, sticky="e")
stopBtn.grid(row=4, column=1, sticky="e")

root.mainloop()