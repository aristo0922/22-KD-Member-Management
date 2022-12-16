import pandas as pd
import os

import tkinter as tk
from tkinter import *
from tkinter import filedialog
import tkinter.messagebox as mbox

import pyautogui


root = Tk()
# 파일 열기

def open_pd(addr):
    excel_df=pd.read_excel(addr, engine="openpyxl")
    excel_drop_df= excel_df.drop(['과목명','훈련기관 PK', '훈련기관과정 PK','훈련기관수업PK','이수 누계(초)','접근IP'], axis=1)


    #0.0분 이수 데이터 삭제
    excel_drop_df=excel_drop_df[~excel_drop_df['강의 이수기간(분)'].isin([0.0])]

    if type(excel_drop_df['학습기준일'][0]) == str :
        excel_drop_df.drop([0], axis=0, inplace=True)

    start_date=get_input_date()

    # 2단위기간 이전 데이터 삭제
    target = excel_drop_df['학습기준일'].ge(start_date)
    excel_drop_df=excel_drop_df[target]
    excel_drop_df.columns = excel_drop_df.columns.str.strip()

    excel_check_df = excel_drop_df.groupby(['차시구분','훈련생 성명','학습기준일']).max()
    excel_check_df = excel_check_df.groupby(['학습기준일', '훈련생 성명']).sum()
    
    # 단위기간 최종 출결 
    if excel_check_df.get('진도율(%)') is not None:
        print(excel_check_df.head(10))
        print("<<<<         columns list        >>>")
        print(excel_check_df.columns.tolist())
        print("true")
    else :
        print(excel_check_df.head(10))
        print("<<<<         columns list        >>>")
        print(excel_check_df.columns.tolist())
        print("false")
        root.destroy()



    excel_check_df['이수시간(초)']=(excel_check_df['진도율(%)']*36).round(0)
    excel_check_df['이수시간(초)'].astype(int)
    excel_check_df= excel_check_df.drop(['진도율(%)', '강의 시간(초)','강의 이수기간(분)'], axis=1)
    excel_check_df = excel_check_df.sort_values(by=['훈련생 성명', '학습기준일'])


    # 파일 저장
    process_name = get_input_process()
    current_addr = os.getcwd()
    print(current_addr)
    # excel_check_df.to_excel(addr+process_name+' .xlsx')
    excel_check_df.to_excel(current_addr+'\Downloads\(result) 온라인출석 '+process_name+'.xlsx')

def file_find():
    file = filedialog.askopenfilename(initialdir=r'C:/', title='select file', filetypes=(('excel file','*.xlsx'),('all files','*.*')))
    root.filename=file
    dir_label = tk.Label(root, text=root.filename)
    dir_label.place(x=0, y=0)


def file_upload():
    if root.filename=='':
        mbox.showinfo('warning', 'Please selecct file')
        return
    else:
        open_pd(root.filename)
        root.destroy()
        return

# 단위기간 시작일 받기
def get_input_date() -> int:
    # start_date = tk.simpledialog.askinteger(title= "단위기간 시작일", message="단위기간 시작일(점 빼고 입력, ex/ 20221026 )", parent = root)
    start_date = pyautogui.prompt(text='단위기간 시작일(점 빼고 입력, ex/ 20221026 )', title="단위기간 시작일", default="")

    if start_date is not None:
        return int(start_date)
    else :
        a = pyautogui.alert(text='단위기간 시작일을 입력하지 않았습니다.', title='error', button='OK')
        print(a)

def get_input_process():
    process = pyautogui.prompt(text='과정명을 입력해주세요, ex/ 4기 AI', title="과정 명", default="")
    return process

root.title("출결 데이터 정제 프로그램")
root.geometry("500x200+200+200")
root.resizable(True, True)



try:   
    btn_find = tk.Button(root, text="불러오기", width=10, command=file_find)
    btn_find.pack(side="right", padx=1, pady=1)
    btn_upload = tk.Button(root, text="산출하기", width=10, command=file_upload)
    btn_upload.pack(side="right", padx=1, pady=1)


except:
    pyautogui.alert(text='에러가 발생했습니다. 프로그램을 재실행해주세요.', title='error', button='OK')
    root.destroy()


root.mainloop()