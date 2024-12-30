import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def select_excel_file():
    # tkinter 기본 설정 (GUI 숨기기)
    root = tk.Tk()
    root.withdraw()

    # 파일 선택 대화 상자 열기 (오직 .xlsx 파일만 표시)
    file_path = filedialog.askopenfilename(
        title="엑셀 (.xlsx) 파일을 선택하세요",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if file_path:
        # 파일 확장자 추가 검증
        _, file_extension = os.path.splitext(file_path)
        if file_extension.lower() != '.xlsx':
            messagebox.showerror("확장자 오류", "올바른 형식의 확장자 명이 아닙니다.")
            return None

        try:
            # 엑셀 파일 읽기
            df = pd.read_excel(file_path)
            print("선택한 파일:", file_path)
            print(df)

            # 데이터 변수에 저장
            data = df.values.tolist()
            print(data)
            return data
        except Exception as e:
            messagebox.showerror("읽기 오류", f"파일을 읽는 중 오류가 발생했습니다:\n{e}")
            return None
    else:
        messagebox.showinfo("알림", "파일을 선택하지 않았습니다.")
        return None

if __name__ == "__main__":
    select_excel_file()
