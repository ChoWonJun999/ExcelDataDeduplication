import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


class main:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 데이터 중복 제거")

        self.file_path = None
        self.file_name = tk.StringVar()

        self.df = None

        self.listbox1 = None
        self.listbox2 = None
        self.file_title_list1 = []
        self.file_title_list2 = []

        self.start_button = None
        self.start_state = tk.StringVar()

        self.print_button = None
        self.print_state = tk.StringVar()

        self.create_gui()

    def create_gui(self):
        # 첨부파일 입력 칸 및 키워드 리스트
        file_button = tk.Button(root, text="엑셀 파일 선택", command=self.select_excel, font=("Helvetica", 12))
        file_button.grid(row=0, column=0, padx=20, pady=20, sticky=tk.W)
        file_label = tk.Label(root, textvariable=self.file_name, width=100, font=("Helvetica", 12))
        file_label.grid(row=0, column=1, padx=20, pady=20, sticky=tk.W)

        # 리스트 박스와 스크롤바를 담을 프레임 생성
        frame = tk.Frame(root)
        frame.grid(row=1, column=0, columnspan=2, padx=20, pady=20, sticky=tk.EW)

        # 리스트 박스 위에 라벨 추가
        label1 = tk.Label(frame, text="중복 제거 필드 순위", font=("Helvetica", 12))
        label1.grid(row=0, column=0, padx=10, pady=10)

        label2 = tk.Label(frame, text="필드", font=("Helvetica", 12))
        label2.grid(row=0, column=2, padx=10, pady=10)

        # 리스트 박스 생성 및 항목 추가 (높이를 고정)
        self.listbox1 = tk.Listbox(frame, width=20, height=6, font=("Helvetica", 12))
        self.listbox1.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

        self.listbox2 = tk.Listbox(frame, width=20, height=6, font=("Helvetica", 12))
        self.listbox2.grid(row=1, column=2, sticky="nsew", padx=10, pady=10)

        # 버튼을 담을 프레임 생성 및 리스트 박스 사이에 배치
        button_frame = tk.Frame(frame)
        button_frame.grid(row=1, column=1, sticky="ns")

        # 버튼 생성 및 버튼 프레임에 추가
        button1 = tk.Button(button_frame, text="->", command=self.move_to_right, font=("Helvetica", 12))
        button1.pack(side="top", pady=20, expand=True)

        button2 = tk.Button(button_frame, text="<-", command=self.move_to_left, font=("Helvetica", 12))
        button2.pack(side="top", pady=20, expand=True)

        # 컬럼과 행의 가중치 설정, 윈도우 크기 조절 시 위젯이 적절히 확장되도록 함
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(2, weight=1)
        frame.rowconfigure(1, weight=1)

        # 중복 제거
        self.start_button = tk.Button(root, text="중복 제거", command=self.remove_duple, font=("Helvetica", 12))
        self.start_button.grid(row=2, column=0, padx=20, pady=20, sticky=tk.W)
        self.start_button.config(state=tk.DISABLED)
        start_label = tk.Label(root, textvariable=self.start_state, font=("Helvetica", 12))
        start_label.grid(row=2, column=1, padx=20, pady=20, sticky=tk.W)

        # 엑셀 파일 다운로드
        self.print_button = tk.Button(root, text="엑셀 파일 다운로드", command=self.print_excel, font=("Helvetica", 12))
        self.print_button.grid(row=3, column=0, padx=20, pady=20, sticky=tk.W)
        self.print_button.config(state=tk.DISABLED)
        print_label = tk.Label(root, textvariable=self.print_state, font=("Helvetica", 12))
        print_label.grid(row=3, column=1, padx=20, pady=20, sticky=tk.W)

    # 오른쪽 리스트 박스로 항목을 이동하는 함수
    def move_to_right(self):
        selected_indices = self.listbox1.curselection()
        for i in selected_indices:
            self.listbox2.insert(tk.END, self.listbox1.get(i))
        for i in reversed(selected_indices):
            self.listbox1.delete(i)

    # 왼쪽 리스트 박스로 항목을 이동하는 함수
    def move_to_left(self):
        selected_indices = self.listbox2.curselection()
        for i in selected_indices:
            self.listbox1.insert(tk.END, self.listbox2.get(i))
        for i in reversed(selected_indices):
            self.listbox2.delete(i)

    def select_excel(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", " *.xlsx;*.xls")])

        if self.file_path:
            try:
                df = pd.read_excel(self.file_path, engine='openpyxl', nrows=1, header=None)
                desired_data = df.iloc[0].tolist()

                a1_chk = "휴대폰번호" in desired_data
                b1_chk = "필수사항1" in desired_data
                if a1_chk and b1_chk:
                    self.listbox1.delete(0)
                    self.file_title_list1 = ["휴대폰번호", "필수사항1"]
                    for file_title in self.file_title_list1:
                        self.listbox1.insert(tk.END, file_title)

                    desired_data.remove("휴대폰번호")
                    desired_data.remove("필수사항1")
                    self.file_title_list2 = desired_data
                    self.listbox2.delete(0)
                    for file_title in self.file_title_list2:
                        self.listbox2.insert(tk.END, file_title)
                    pass
                else:
                    tk.messagebox.showwarning("오류", "엑셀 파일의 형식이 맞지 않습니다.")
                    self.file_path = None
                    return
            except Exception as e:
                tk.messagebox.showerror("오류", f"엑셀 파일을 열던 중 오류가 발생했습니다: {str(e)}")
                return

        self.file_name.set(self.file_path)
        self.start_state.set("")
        self.print_state.set("")
        if self.file_path:
            self.start_button.config(state=tk.NORMAL)
            self.print_button.config(state=tk.DISABLED)
        else:
            self.start_button.config(state=tk.DISABLED)
            self.print_button.config(state=tk.DISABLED)

    def remove_duple(self):
        self.df = None

        if self.file_path:
            try:
                # 여러 열을 한 번에 중복 제거
                dedup_fields = self.listbox1.get(0, self.listbox1.size())

                # dtype을 지정하여 데이터를 읽습니다. (예시: 모든 열을 문자열로 읽음)
                dtype_dict = {field: str for field in dedup_fields}

                self.df = pd.read_excel(self.file_path, engine='openpyxl', header=0, dtype=dtype_dict)

                self.df.drop_duplicates(subset=dedup_fields, inplace=True)

                self.start_state.set("중복 제거가 완료됐습니다.")
                self.print_button.config(state=tk.NORMAL)
            except Exception as e:
                tk.messagebox.showerror("오류", f"중복 제거 중 오류가 발생했습니다: {str(e)}")

    def print_excel(self):
        try:
            filename = filedialog.asksaveasfilename(defaultextension=".xlsx", title="Select file",
                                                    filetypes=(("xlsx", "*.xlsx"), ("All Files", "*.*")))
            self.df.to_excel(filename, index=False, engine='openpyxl')
        except Exception as e:
            print('예외가 발생했습니다.', e)
            return
        self.print_state.set("엑셀 파일 다운이 완료됐습니다.")


if __name__ == "__main__":
    root = tk.Tk()

    # 창을 화면 중앙에 배치
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (800 / 2))
    y_coordinate = int((screen_height / 2) - (600 / 2))
    root.geometry(f"1120x600+{x_coordinate}+{y_coordinate}")

    app = main(root)
    root.mainloop()
