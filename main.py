import random
import tkinter as tk
from tkinter import *
from tkinter import messagebox,filedialog
import os
from pathlib import Path
import shutil
import csv
import webbrowser

win = tk.Tk()  # 用win代指窗口window，这句是定义窗口
win.title('随机抽号 v3.3')
win.geometry('720x450')  # 设置窗口大小为宽度*高度，不设置即为最合适大小
win.iconbitmap(Path(__file__).parent / r'res\icon.ico')

Class = []
Lstbox1 = Class
ini_list_csv = 'ini_list.csv'
ini = Path(__file__).parent / r'res\initialization'

# 检查并处理 ini_list.csv 文件
if not os.path.exists(ini_list_csv):
    # 如果文件不存在，则创建初始文件
    shutil.copyfile(ini, 'ini_list.csv')
# 文件已存在，读取文件内容
with open(ini_list_csv, 'r', newline='') as initial_list:
    ini_reader = csv.reader(initial_list)
    for row in ini_reader:
        Class.append(row[0])

# ============================================================ 函数定义
# 凡在定义函数时出现参数x的均为占位用的无效参数，是为防止快捷键调用时出现bug而设定，初始化x=0


def Close(x=0):  # 关闭
    ifQ = tk.messagebox.askquestion(title='提示', message='您确定要退出吗？')
    if ifQ == 'yes':
        win.destroy()
    else:
        return 0


def Void():
    return  # 这是一个空指令


def Begin(x=0):  # 抽人
    try:
        A = int(random.randint(0, len(Class)-1))
    except ValueError:
        tk.messagebox.showwarning(title='警告', message='名单为空！')
    else:
        T.insert('end', Class[A])
        T.insert('end', '\t')


def Clear(x=0):
    T.delete('1.0', 'end')  # 全选删除旧的内容


def ini(x=0):  # 设置-刷新
    Lstbox1.delete(0, END)
    list_items = Class
    for item in list_items:
        Lstbox1.insert(END, item)


def clr(x=0):  # 设置-清空列表
    ifClr = messagebox.askyesno(title='清空', message='您确定要清空名单内所有项吗？\n注意：再次打开软件即可恢复为预置名单！')
    if ifClr:
        Lstbox1.delete(0, END)
        Class.clear()
        messagebox.showinfo(title='清空', message='您已清空名单内所有项！')


def ins(x=0):  # 设置-添加名字
    if entry.get() != '':
        if Lstbox1.curselection() == ():
            Lstbox1.insert(Lstbox1.size(), entry.get())
        else:
            Lstbox1.insert(Lstbox1.curselection())
        Class.append(entry.get())
        messagebox.showinfo(title='添加提示', message='您已成功添加“'+entry.get()+'”！')
    else:
        messagebox.showinfo(title='添加提示', message='请在文本框中输入要添加的人名。')


def mdf(x=0):  # 设置-修改
    num = Lstbox1.curselection()
    try:
        int(str(num).replace('(', '').replace(',', '').replace(')', ''))
    except ValueError:
        messagebox.showinfo(title='修改提示',message='请在左侧选择即将修改的人名。')
    else:
        ex_name = Class[int(str(num).replace('(', '').replace(',', '').replace(')', ''))]  # 修改前名字
    name = entry.get()
    if name == '':
        messagebox.showinfo(title='修改提示',message='请在文本框中输入新的人名。')
    if name != '' and num != ():
        selected = Lstbox1.curselection()[0]
        Lstbox1.delete(selected)
        Lstbox1.insert(selected,name)
        Class[selected] = name
        tk.messagebox.showinfo(title='修改提示', message='您已成功将“'+str(ex_name)+'”修改为“'+name+'”！')


def delt(x=0):  # 设置-删除名字
    num = Lstbox1.curselection()
    try:
        int(str(num).replace('(', '').replace(',', '').replace(')', ''))
    except ValueError:
        messagebox.showinfo(title='删除提示', message='请在左侧选择要去除的人名。')
    else:
        name = Class[int(str(num).replace('(', '').replace(',', '').replace(')', ''))]
    if num != ():
        Lstbox1.delete(num)
        Class.remove(name)
        tk.messagebox.showinfo(title='删除提示', message='您已成功删除“'+name+'”！')


def imp(x=0):  # 设置-导入列表
    ifcon = messagebox.askyesno(title='注意',
                                message='1. 导入的CSV文件必须是一列数据。\n2. 导入新列表，将会在原有列表的基础上添加，\n    您确定要继续吗？')
    if ifcon:
        file_path = filedialog.askopenfilename(filetypes=[("CSV文件", "*.csv")])
        if file_path:
            with open(file_path, 'r', newline='') as file:
                reader = csv.reader(file)
                for row in reader:
                    try:
                        Lstbox1.insert(END, row[0])
                        Class.append(row[0])
                    except Exception:
                        pass


def Set2(x=0):  # “设置”面板
    global entry
    global Lstbox1
    root = Toplevel(win)
    root.title('设置')
    root.geometry('320x380')
    root.transient(win)  # msgbox式窗口，即设为窗口只有关闭键
    root.protocol('WM_DELETE_WINDOW', Void)  # 点设置窗口的关闭会没反应因为执行了定义的空函数Void
    tLabel = Label(root, text='抽点列表设置', height=2, font=('黑体', 14))
    tLabel.pack()
    frame1 = Frame(root,relief=tk.GROOVE)
    frame2 = Frame(root,relief=tk.GROOVE)
    frame1.place(relx=0.02, rely=0.12)
    frame2.place(relx=0.51, rely=0.12)
    Lstbox1 = Listbox(frame1, font=('宋体', 12), width=18, height=13)
    tLstbox1 = Label(frame1, text='名单', font=('黑体', 13))
    entry = Entry(frame2)
    tentry = Label(frame2, text='文本框', font=('黑体', 13))
    tB4btn = Label(frame2, text=' ')  # text before button，即按钮前的空行
    tLstbox1.pack()
    Lstbox1.pack()
    tentry.pack()
    entry.pack()
    tB4btn.pack()
    btn1 = Button(frame2, text='刷新 (F5)', fg='#FF50FF', cursor='hand2', width=16, command=ini())  # 加括号即打开窗口自动运行一次
    btn2 = Button(frame2, text='添加 (F1)', fg='green', cursor='hand2', width=16, command=ins)  # cursor鼠标形状
    btn3 = Button(frame2, text='修改 (F2)', fg='blue', cursor='hand2', width=16, command=mdf)
    btn4 = Button(frame2, text='删除 (F10)', fg='red', cursor='hand2', width=16, command=delt)
    btn5 = Button(frame2, text='清空', fg='purple', cursor='hand2', width=16, command=clr)
    btn6 = Button(frame2, text='导入列表', fg='#FFAA00', cursor='hand2', width=16, command=imp)
    root.bind("<F1>", ins)    # 添加  # 设置界面中的快捷键
    root.bind("<F10>", delt)  # 删除
    root.bind("<F2>", mdf)    # 修改
    root.bind("<F5>", ini)    # 刷新
    btn1.pack()
    btn2.pack()
    btn3.pack()
    btn4.pack()
    btn5.pack()
    btn6.pack()
    okSet = Button(root, text='完成设置', cursor='hand2', bd=3, height=2, width=10,  # bd为按钮边框粗细
                   font=('黑体', 12), fg='blue', command=lambda: root.destroy())  # lambda为匿名函数，可在command中执行冒号后的内容
    okSet.place(relx=0.35, rely=0.8)
    root.mainloop()


file_path = ''
file_text = ''  # 写入文件内的文本


def Newfile(x=0):  # 新建文件（菜单）
    global file_path
    global file_text
    messagebox.showinfo(title='新建',message='温馨提示：\n新建文件时，请输入文件名及其扩展名，\n若未输入则默认为*.txt')
    file_path = filedialog.asksaveasfilename(title=u'新建文件',
                                             filetypes=[('文本文档（推荐）', '*.txt'),('新版Word文档', '*.docx'),
                                                        ('新版Excel文档', '*.xlsx'),('全部文件', '*')],
                                             defaultextension='*.txt')  # 保存面板标题、保存类型定义、默认扩展名
    file_text = ''
    if file_path is not None:
        try:
            with open(file=file_path, mode='a+', encoding='utf-8') as file:
                file.write(file_text)
        except FileNotFoundError:
            messagebox.showerror(title='新建', message='新建失败！')
        else:
            messagebox.showinfo(title='新建', message='新建成功！')


def Savefile(x=0):  # 保存文件（菜单）
    global file_path
    global file_text
    file_path = filedialog.asksaveasfilename(title=u'保存文件',
                                             filetypes=[('文本文档（推荐）', '*.txt'), ('新版Word文档', '*.docx'),
                                                        ('新版Excel文档', '*.xlsx'), ('全部文件', '*')])  # 保存类型定义
    file_text = T.get(1.0, 'end')
    if file_path is not None:
        try:
            with open(file=file_path, mode='a+', encoding='utf-8') as file:
                file.write(file_text)
        except FileNotFoundError:
            messagebox.showerror(title='保存', message='取消保存！')
        else:
            messagebox.showinfo(title='保存', message='输出内容已成功保存！')


def Showlist(x=0):  # 显示列表
    showClass = Toplevel(win)
    showClass.title('名单')
    showClass.transient(win)  # msgbox式窗口，即设为窗口只有关闭键
    ClassList = Listbox(showClass, font=('宋体', 13), width=30, height=30)
    for item in Class:
        ClassList.insert('end', item)
    slLabel = Label(showClass, text='抽取名单', font=('黑体', 14), height=3, justify="center", anchor='center')
    slTell = Label(showClass, text='\n    此列表仅供查看抽取名单，若要修改请前往“设置”。\n编辑预置名单入口：软件所在目录的“ini_list.csv”文件\n',
                   font=('宋体', 9), justify='left', anchor='center')
    slLabel.pack()
    ClassList.pack()
    slTell.pack()


def update_log():  # 更新日志
    popup_log = Toplevel(win)
    popup_log.title("更新日志")
    popup_log.transient(win)
    scrollbar = tk.Scrollbar(popup_log)
    scrollbar.pack(side='right', fill='y')
    textbox = tk.Text(popup_log, height=30, width=100, font=('宋体', 13), yscrollcommand=scrollbar.set)
    textbox.insert('end', """2022.01.23 22:39 | Lottery v1.0.220123 beta:\n推出简易版小程序，具有老虎机一般的简单抽号功能\n
2022.03.22 23:43 | Lottery - name v1.2.1:\n采用GUI交互式设计，添加了名单修改工具，增设了常用功能的快捷键\n
2022.05.29 12:27 | Lottery - name v2.0:\n完善了名单修改工具，增设了新建文件、保存文件、显示抽取列表等功能\n
2022.05.29 19:05 | Lottery - name v2.1:\n完善了上个版本添加的功能，增加了帮助栏，继续增设了更多快捷键\n
2022.10.02 0:16 | 随机点名小程序 v2.2:\n完善了既有功能，优化诸多场景下的用户体验，继续增设了更多快捷键\n
2024.07.04 10:30 | 随机抽号小程序 v3.0:\n增设批量导入名单的功能，重设了初始名单的读取逻辑，更新了版权信息中的联系方式\n
2024.07.31 23:58 | 随机抽号小程序 v3.1:\n将初始名单设为自动识别模式，若计算机上没有初始名单，则在程序目录下新建初始名单\n
2024.08.01 11:40 | 随机抽号小程序 v3.2:\n增设自动初始化功能，修复了系统Temp文件夹内名单缺失等问题，加入了更新日志\n
2025.09.19 22:20 | 随机抽号小程序 v3.3:\n修复了读取CSV文件时的bug，解决了文字不符合UTF-8编码时读取失败的问题""")
    textbox.pack(side='left', fill='both')
    scrollbar.config(command=textbox.yview)


def cprt():  # 版权信息
    popup = Toplevel(win)
    popup.title("版权信息与联系方式")
    popup.transient(win)  # msgbox式窗口，即设为窗口只有关闭键
    cprt_label = tk.Label(popup, text='\n\n开发者\n\nH.Jupiter.Lyr\n\n\n版权所有者\n\nH.Jupiter.Lyr\n\n\n邮箱',
                          font=('Times New Roman', 14))
    email_label = tk.Label(popup, text='jupiterlyr@foxmail.com', fg='blue', cursor='hand2',
                           font=('Times New Roman', 14))
    close_button = tk.Button(popup, text='关闭', font=('黑体', 14), cursor='hand2', width=8, command=lambda: popup.destroy())
    cprt_label.pack(padx=30)
    email_label.pack(padx=30, pady=15)
    email_label.bind("<Button-1>", lambda: webbrowser.open(f"mailto:'jupiterlyr@foxmail.com'"))
    close_button.pack(pady=40)


def Help(x=0):  # 帮助
    h = Toplevel(win, bg='white')
    h.title('帮助')
    h.transient(win)  # msgbox式窗口，即设为窗口只有关闭键
    hTitle = Label(h, text='帮助手册', font=('黑体', 15), bg='white')
    hQuit = Button(h, text='我知道了', width=10, height=1, bg='white', cursor='hand2', command=h.destroy)
    hText = Label(h, text="""
一、快捷键大全
    说明：尖括号<>括起部分为对应热键，冒号后为描述性文字。
    1. 主页面 - 抽取 <=>：退格键（Backspace）旁的等于号。
    2. 主页面 - 清空 <Delete>：台式机为方向键上方的Del键。
    3. 退出 <Alt+F4>：按住Alt键按F4即可退出。
    4. 新建 <Ctrl+N>：按住Ctrl键按N即可新建一个文件。
    5. 保存 <Ctrl+S>：按住Ctrl键按S即可保存抽取的内容。
    6. 复制 <Ctrl+C>：文本显示窗可自由选择文本并复制。
    7. 设置 <F8>。
    8. 设置 - 刷新 <F5>；添加 <F1>；修改 <F2>；删除 <F10>。
    9. 查看抽取列表 <F3>。

二、功能介绍
    1. 抽取相关
           抽取时，每点一次“抽签”按钮或按一次“=”只会抽取一个人。每抽取
       一次不会覆盖历史记录，仅在点击“清空”或按“Delete”时才会清空记录。
    2. 文本显示窗的可复制性
           显示窗内容可以复制，也可以使用保存功能，自动写入显示窗内的内容。
    3. 抽取列表的更改
           您可以通过“选项-设置”来更改抽取列表。在文本框中输入文字，点击
       “添加”即可加入列表；选择列表中某一人名，并在文本框内输入新的名字，
       点击“修改”即可改名；选择选择列表中某一人名，点击“删除”即可移除该
       人名。
           “清空”将清除所有项，仅在重新启动程序后才能恢复默认列表。
           “读取列表”允许您读取计算机上的一个*.csv文件，实现批量导入。
           在完成更改后，请点击下方“完成设置”退出设置面板。
    4. 新建与保存
           使用“新建”帮助您选择合适的存放文件类型。此处新建文件与在文件夹
       中手动新建并无本质区别，配合“保存”功能，可用于存放抽取结果。
           “保存”时仍需选择要写入的文件，当弹出是否要覆盖原文件时，请选择
       “是”；实际上保存仅会在对应的文件末续写，并不会覆盖原文本。

三、必要说明
    1. 若要更改默认的初始化名单，请前往软件所在根目录，修改“ini_list.csv”
       中的内容
    2. 如产品有问题需反馈，请在“选项-版权信息”中联系作者。制作不易，敬请
       理解，谢谢！
""", font=('宋体', 13), bg='white', width=80, justify="left", anchor='center')
    hTitle.pack()
    hText.pack()
    hQuit.pack()
    h.resizable(False, False)
    h.mainloop()


# ============================================================ 主页面
Title = Label(win, text='随机抽号小程序', font=('黑体', 15), height=2, justify='center', anchor='center')
Lt = Label(win, text='''
按下<>内的快捷键实现对应功能：
    <=> 抽取      <Delete> 清空      <F12> 帮助（更多快捷键见此）

操作说明：
 1、点击“抽签”按钮即可开始抽签，每点击一次抽一个人。
 2、显示窗内容可复制，每次抽签都不会自动清理上一次内容；点“清空”即可
    清空显示窗中的所有内容。
 3、查看抽取列表、更改名单等，请在选项中查看或设置；更多说明详见帮助。
''', font=('宋体', 13), height=10, justify="left", anchor='center')
T = Text(win, height=10, font=('宋体', 12))
begin = Button(win, text='抽签', fg='green', width=6, height=1, cursor='hand2', command=Begin)  # “抽签”按钮，鼠标为手指
clear = Button(win, text='清空', fg='red', width=6, height=1, cursor='hand2', command=Clear)  # “清空”按钮
line1 = Label(win, text='')
line2 = Label(win, text='')
Title.pack()
Lt.pack()
line1.pack()
begin.place(relx=0.4, rely=0.48)
clear.place(relx=0.53, rely=0.48)
line2.pack()
T.pack()

win.bind('<=>', Begin)  # 快捷键
win.bind('<Delete>', Clear)

# ============================================================ 菜单栏
menubar = tk.Menu(win)  # 定义菜单栏menubar
fileM = Menu(menubar, tearoff=0)  # tearoff=0下拉菜单，=1下拉可独立显示菜单
optionM = Menu(menubar, tearoff=0)
helpM = Menu(menubar, tearoff=0)
menubar.add_cascade(label='文件', menu=fileM)
menubar.add_cascade(label='选项', menu=optionM)
menubar.add_cascade(label='帮助', menu=helpM)
fileM.add_command(label='新建', command=Newfile, accelerator='Ctrl+N')
fileM.add_command(label='保存', command=Savefile, accelerator='Ctrl+S')
fileM.add_separator()  # 分割线
fileM.add_command(label='退出', command=Close, accelerator='Alt+F4')
optionM.add_command(label='设置', command=Set2, accelerator='F8')
optionM.add_command(label='显示名单', command=Showlist, accelerator='F3')
optionM.add_command(label='更新日志', command=update_log)
optionM.add_command(label='版权信息', command=cprt)
helpM.add_command(label='使用说明', command=Help, accelerator='F12')

win.config(menu=menubar)  # 设置菜单栏

win.bind('<F8>', Set2)
win.bind('<F3>', Showlist)
win.bind('<F12>', Help)
win.bind('<Control-n>', Newfile)
win.bind('<Control-s>', Savefile)
win.bind('<Alt-F4>', Close)

# ============================================================ 其他
win.protocol('WM_DELETE_WINDOW', Close)  # 点右上角的退出(×)会执行Close函数
win.resizable(False, False)  # 禁用窗口最大化；resizable(横向是否可调,纵向是否可调)
win.mainloop()  # 运行弹窗

