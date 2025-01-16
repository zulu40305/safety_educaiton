import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def select_excel_files():
    # 파일 선택 대화 상자 열기 (여러 파일 선택 가능, 오직 .xlsx 파일만 표시)
    file_paths = filedialog.askopenfilenames(
        title="안전교육 산정",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not file_paths:
        return

    # 파일 개수 검증
    if len(file_paths) != 2:
        messagebox.showerror("선택 오류", "2개의 파일을 선택해야 합니다. 다시 시도하세요.")
        return

    # 두 개의 엑셀 파일 읽기
    try:
        df1 = pd.read_excel(file_paths[0])
        df2 = pd.read_excel(file_paths[1])
        print("첫 번째 파일 데이터:\n", df1)
        print("두 번째 파일 데이터:\n", df2)
    except Exception as e:
        messagebox.showerror("읽기 오류", f"선택한 파일의 확장자가 올바르지 않습니다:\n{e}")
        return

    # 두 데이터프레임 합치기 (예: 위아래로 추가)
    combined_df = pd.concat([df1, df2], ignore_index=True)
    print("결합된 데이터:\n", combined_df)

    # 새로운 엑셀 파일로 저장
    save_path = filedialog.asksaveasfilename(
        title="파일 저장",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if save_path:
        try:
            combined_df.to_excel(save_path, index=False)
            messagebox.showinfo("저장 완료", f"선택한 위치에 성공적으로 파일을 저장하였습니다:\n{save_path}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다:\n{e}")
    else:
        messagebox.showinfo("알림", "파일 저장이 취소되었습니다.")

def main():
    # Tkinter GUI 창 생성
    root = tk.Tk()
    root.title("안전교육 산정 프로그램")

    # 창 크기 설정
    root.geometry("350x150")

    # 파일 선택 및 병합 버튼
    select_button = tk.Button(root, text="파일 선택", command=select_excel_files, font=("Arial", 11))
    select_button.pack(pady=20)

    # 종료 버튼
    quit_button = tk.Button(root, text="종료", command=root.destroy, font=("Arial", 11))
    quit_button.pack(pady=15)

    # Tkinter 이벤트 루프 실행
    root.mainloop()

if __name__ == "__main__":
    main()
