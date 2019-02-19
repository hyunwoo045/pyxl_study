from tkinter import *
from tkinter import messagebox, filedialog
from tkinter.ttk import *
from datetime import datetime, timedelta
import openpyxl


class car_exDate_cal(Frame):
    ## 프로그램 GUI 기본틀 제작 ##
    def __init__(self, master):
        fields = '차량 번호', '등록일', '만료일'
        ents = self.makeform(master, fields)
        print(ents)

        row = Frame(master)
        lab = Label(row, width=15, text="파일: ", anchor='w')
        ent = Entry(row)
        btn = Button(row, text="browser", command=self.onClickFile)
        row.pack(side=TOP, fill=X, padx=5, pady=5)
        lab.pack(side=LEFT)
        ent.pack(side=LEFT, expand=YES, fill=X)
        btn.pack(side=RIGHT, padx=5, pady=5)
        master.bind("<Return>", lambda eff: self.onSaveFile(ents[0][1], ents[1][1], ents[2][1]))
        b1 = Button(master, text='Save',
                    command=(lambda e=ents: self.onSaveFile(ents[0][1], ents[1][1], ents[2][1])))
        b1.pack(side=LEFT, padx=5, pady=5)
        b2 = Button(master, text='Quit', command=master.quit)
        b2.pack(side=LEFT, padx=5, pady=5)



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

    def onSaveFile(self, a, b, c):
        message = a.get()
        message2 = b.get()
        message3 = c.get()
        messagebox.showinfo(message, message2 + " " + message3)

        # wb = openpyxl.load_workbook("C:/example.xlsx")

    def makeform(self, root, fields):
        entries = []
        for field in fields:
            row = Frame(root)
            lab = Label(row, width=15, text=field, anchor='w')
            ent = Entry(row)
            row.pack(side=TOP, fill=X, padx=5, pady=5)
            lab.pack(side=LEFT)
            ent.pack(side=RIGHT, expand=YES, fill=X)
            entries.append((field, ent))
        return entries

def main():
    root = Tk()
    app = car_exDate_cal(root)
    root.mainloop()

if __name__ == '__main__':
    main()
