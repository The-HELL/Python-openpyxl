import _tkinter
import os
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from Append_xlsx import get_name_value, append_xlsx

class Append(object):

    def __init__(self, root):
        self.root = root

        self.root.title("排班")
        # 禁止调整窗口大小
        self.root.resizable(FALSE, FALSE)
        # Frame1:mainframe
        mainframe = ttk.Frame(self.root, padding="5 5 5 5")
        mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)


        # Frame2:daysframe
        daysframe = ttk.Frame(mainframe)
        daysframe.grid(column=3, row=3, sticky=(N, W, E, S))


        # StringVar
        self.model_filename_s = StringVar()
        self.filename_s = StringVar()
        self.begin_date_s = StringVar()
        self.days_s = IntVar()


        # entry
        model_filename_entry = ttk.Entry(mainframe, width=55, textvariable=self.model_filename_s)
        filename_entry = ttk.Entry(mainframe, width=55, textvariable=self.filename_s)
        begin_date_entry = ttk.Entry(mainframe, width=17, textvariable=self.begin_date_s)
        days_entry = ttk.Entry(daysframe, width=2, textvariable=self.days_s)

        # Label
        model_filename = ttk.Label(mainframe, text="模板文件：")
        filename_label = ttk.Label(mainframe, text="文件名：")

        begin_date_label = ttk.Label(mainframe, text="开始日期（xxxx.xx.xx）：")
        days_label = ttk.Label(daysframe, text="添加天数：")


        # button
        get_filename = ttk.Button(mainframe, text="排班文件", command=self.GetPath)
        get_model_filename = ttk.Button(mainframe, text="模板文件", command=self.GetModelFilename)
        Finsh = ttk.Button(root, text="添加", command=self.check_run)

        # grid
        model_filename.grid(column=1, row=1, sticky=E)
        filename_label.grid(column=1, row=2, sticky=E)
        begin_date_label.grid(column=1, row=3, sticky=E)
        days_label.grid(column=1, row=1, sticky=E)

        get_filename.grid(column=3, row=2, sticky=W)
        get_model_filename.grid(column=3, row=1, sticky=W)
        model_filename_entry.grid(column=2, row=1, sticky=W)
        filename_entry.grid(column=2, row=2, sticky=W)
        begin_date_entry.grid(column=2, row=3, sticky=W)

        days_entry.grid(column=2, row=1, sticky=E)

        Finsh.grid(column=0, row=1, sticky=(W, E))

        child_all = []
        child_all.append(self.root.winfo_children())
        child_all.append(mainframe.winfo_children())
        child_all.append(daysframe.winfo_children())

        for parent in child_all:
            for child in parent:
                child.grid_configure(padx=5, pady=5)

        model_filename_entry.focus()
        self.root.bind("<Return>", self.run)
        self.center_window(700, 200)

    def check_filename(self, filename:str):
        suffix = os.path.splitext(filename)[-1]
        if suffix != ".xlsx":
            return False
        else:
            return True

    def check_run(self):
        try:
            model_filename = self.model_filename_s.get()
            filename = self.filename_s.get()
            begin_date = self.begin_date_s.get()
            days = int(self.days_s.get())

        except _tkinter.TclError:
            messagebox.showerror(title="类型错误", message="请检查填写内容的类型是否正确！\n\n添加天数应为int类型")

        else:

            if not model_filename:
                self.Print_Message(message="请添加模板文件")
            elif not filename:
                self.Print_Message(message="请添加排班文件")

            # 检查文件格式是否正确
            elif not self.check_filename(model_filename) or not self.check_filename(filename):
                message = "只接受“xlsx”格式的文件\n请检查文件是否正确！！！"
                self.Print_Message(message=message)
            elif not begin_date:
                self.Print_Message(message="请填写开始日期")
            elif not days:
                self.Print_Message(message="请填写天数")
            else:
                self.run(model_filename=model_filename, filename=filename, begin_date=begin_date, days=days)


    def GetPath(self):
        # 选择文件path_接收文件地址
        path_ = filedialog.askopenfilename()
        self.filename_s.set(path_)

    def GetModelFilename(self):
        # 选择模板文件地址
        path_ = filedialog.askopenfilename()
        self.model_filename_s.set(path_)

    def Print_Message(self, message):
        messagebox.showinfo(message=message)

    def center_window(self, w, h):
        # 获取屏幕 宽、高
        ws = self.root.winfo_screenwidth()
        hs = self.root.winfo_screenheight()

        # 计算x, y位置
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2)

        self.root.geometry("%dx%d+%d+%d" % (w, h ,x, y))

    def run(self, **kwargs):
        try:
            model_filename = kwargs.get("model_filename")
            filename = kwargs.get("filename")

            begin_date = kwargs.get("begin_date")
            days = kwargs.get("days")

            name_value = get_name_value(model_filename=model_filename)

            if append_xlsx(name_value=name_value, filename=filename, begin_date=begin_date, days=days):
                message = "添加成功！"
            else:
                message = "添加失败！！！\n\n\t\t请联系管理员检查！！！"

            self.Print_Message(message=message)

            if message == "添加成功！":
                self.root.destroy()

        except ValueError:
            pass


if __name__ == "__main__":
    root = Tk()
    Append(root)
    root.mainloop()