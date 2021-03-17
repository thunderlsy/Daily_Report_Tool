import time
import tkinter as tk
import os
import threading
from tkinter import filedialog
from open_mul_file import OpenMulFile
from generate_test_log import GenerateTestLog


# ======================================================
# @Author : Chris
# @Time : 5/11/2020 7:00 PM
# @Desc : 图形用户界面
# ======================================================


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)

        self.blue_list_now = 0
        self.blue_list_len = 3

        self.root = master
        self.root.geometry('600x350+600+300')
        self.root.title('Rel_Daily_Report_Tool')
        # self.root.bind("<Motion>", self.call_back)
        self.frm1 = tk.Frame(self.root)
        self.frm2 = tk.Frame(self.root)
        self.frm3 = tk.Frame(self.root)

        self.var = tk.StringVar()
        self.var.set("Run")

        # 文件选择框frm1
        self.frm1.config(bg='#4682B4', height=255, width=160)
        # tk.Label(self.frm1, text='运行框', fg='black').place(anchor=tk.NW)
        self.frm1.place(x=20, y=5)
        tk.Label(self.frm1, text='1.在此选择细项', bg='#4682B4',
                 fg='black', font='Verdana 12 bold').place(x=2, y=5)
        tk.Label(self.frm1, text='2.在此选择Daily_Report', bg='#4682B4',
                 fg='black', font='Verdana 12 bold').place(x=2, y=75)
        tk.Label(self.frm1, text='3.开始运行', bg='#4682B4',
                 fg='black', font='Verdana 12 bold').place(x=2, y=145)
        tk.Button(self.frm1, text='请选若干文件', height=1, width=10, command=self._get_detail_files).place(x=2, y=30)
        tk.Button(self.frm1, text='请选单个文件', height=1, width=10, command=self._get_Daily_report_files).place(x=2, y=100)
        self.run_bt = tk.Button(self.frm1, textvariable=self.var, height=1, width=10, command=self._run_entrance)
        self.run_bt.place(x=2, y=170)

        # 进度条frm2
        self.frm2.config(bg='#4682B4', height=60, width=160)
        self.frm2.place(x=20, y=265)
        # 创建一个背景色为白色的矩形
        self.canvas = tk.Canvas(self.frm2, width=150, height=20, bg="#4682B4")
        self.canvas.place(x=2, y=15)
        # 创建一个矩形外边框（距离左边，距离顶部，矩形宽度，矩形高度），线型宽度，颜色
        self.out_line = self.canvas.create_rectangle(2, 2, 180, 27, width=1, outline="black")

        # 信息框frm3
        self.frm3.config(height=320, width=400)
        self.frm3.place(x=190, y=5)
        self.notes_tx = tk.Text(self.frm3, bd=0, width=70, height=24, bg='#4682B4', font=('Arial', 12))
        self.notes_tx.place(x=0, y=0)

        self.pack()

    # 打开文件并显示路径
    def _get_detail_files(self):
        # self.get_detail_bt.config(state="disable")
        default_dir = r"细项文件路径"
        self.details_files_path = tk.filedialog.askopenfilenames(title=u'选择文件',
                                                                 initialdir=(os.path.expanduser(default_dir)))
        print(self.details_files_path)

        self.notes_tx.insert('end', "正在选择细项表\n")

        for i in self.details_files_path:
            self.notes_tx.insert('end', "已选中" + i[i.rfind("/") + 1:] + "\n")

    def _get_Daily_report_files(self):
        default_dir = r"Daily_Report文件路径"
        self.daily_file_path = tk.filedialog.askopenfilename(title=u'选择文件',
                                                             initialdir=(os.path.expanduser(default_dir)))

        # self.notes_tx.insert('end', "正在选择Daily_report\n" + self.daily_file_path[self.daily_file_path.rfind("/") +
        # 1:] + "\n")

        self.notes_tx.insert('end', "正在选择Daily_report\n")
        if self.daily_file_path:
            self.notes_tx.insert('end', "已选中" + self.daily_file_path[self.daily_file_path.rfind("/") + 1:] + "\n")

    @staticmethod
    def thread_it(func, *args):
        """将函数打包进线程"""

        # 创建
        t = threading.Thread(target=func, args=args)
        # 守护
        t.setDaemon(True)
        # 启动
        t.start()
        # 阻塞--卡死界面！
        # t.join()

    def task_thread_1(self):
        self.notes_tx.insert('end', "-" * 90 + "\n正在读取工作簿......\n")
        self.notes_tx.update()
        tem_obj_open = OpenMulFile(self.details_files_path, self.daily_file_path)
        save_path = tem_obj_open.give_value()
        self.blue_list_now += 1

        self.notes_tx.insert('end', "-" * 90 + "\n正在处理工作簿......\n")
        self.notes_tx.update()
        tem_obj_generate = GenerateTestLog(save_path)
        # self.thread_it(self.task_thread_2, tem_obj_generate)
        tem_obj_generate.run()
        self.blue_list_now += 1

        self.notes_tx.insert('end', "-" * 90 + "\n正在保存工作簿......\n")
        tem_obj_generate.save_book()
        self.blue_list_now += 1

        self.notes_tx.insert('end', "-" * 90 + "\n新的Daily_Report已经生成在原Daily_Report文件夹中，请查看\n")
        self.notes_tx.update()
        pass

    def task_thread_2(self, tem_obj_generate):
        while True:
            self.blue_list_len = tem_obj_generate.blue_list_len
            self.blue_list_now = tem_obj_generate.blue_list_now
            time.sleep(0.01)
            if self.blue_list_len == self.blue_list_now:
                break
        pass

    def _run_entrance(self):
        self.run_bt.config(state="disable")  # 设置按钮只允许点击一次

        self.thread_it(self.task_thread_1)

        """任务进度条"""
        fill_line = self.canvas.create_rectangle(2, 2, 0, 27, width=0, fill="blue")

        while True:
            try:
                time.sleep(2)
                read_percentage = self.blue_list_now / self.blue_list_len
                n = 150 * read_percentage
                # 以矩形的长度作为变量值更新
                self.canvas.coords(fill_line, (0, 0, n, 30))
                self.var.set(str(self.blue_list_now) + "/" + str(self.blue_list_len))
                print(self.blue_list_now, self.blue_list_len)
                if read_percentage == 1:
                    break
                self.update()
            except:
                pass


if __name__ == '__main__':
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()

    single_file_path = ""
