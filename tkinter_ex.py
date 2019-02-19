from tkinter import *
from tkinter import messagebox, filedialog
from tkinter.ttk import *
from datetime import datetime, timedelta
import openpyxl


class car_exDate_cal(Frame):
    ## 프로그램 GUI 기본틀 제작 ##
    def __init__(self, master):
        Frame.__init__(self, master)

        self.master = master
        self.master.title("등록차량 만료일 근접차량 추출기")
        self.pack(fill=BOTH, expand=True)

        frame = Frame(self)
        frame.pack(fill=X)
        lbl = Label(frame, text="엑셀 파일을 불러오기")
        lbl.grid(row=0, column=0)
        btn = Button(frame, text="불러오기", command=self.onClickFile)
        btn.grid(row=1, column=0)

    ## 파일을 불러오면 바로 만료일 계산 및 경고 메시지박스 출력 ##
    def onClickFile(self):
        file_path = filedialog.askopenfilename(initialdir="C:/", title="파일 선택")
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active

        for r in ws.rows:
            registered_carcode = r[0].value
            registered_date = r[1].value
            expire_date = registered_date + timedelta(days=30)
            time_left = expire_date - datetime.now()

            if time_left < timedelta(days=20):
                messagebox.showinfo("만료 경고!", "차량번호: " + registered_carcode + "\n"
                                    "남은기간: " + str(time_left))
            else:
                pass



def main():
    root = Tk()
    root.geometry("150x50+100+100")
    app = car_exDate_cal(root)
    root.mainloop()

if __name__ == '__main__':
    main()
