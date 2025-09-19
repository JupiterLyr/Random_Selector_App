import random
import tkinter as tk
from tkinter import messagebox, filedialog, Toplevel, Listbox, Scrollbar, Text, Menu, Label, Button, Entry, Frame, END
import os
from pathlib import Path
import shutil
import csv
import webbrowser

# ============================================================ 全局常量
WIN_WIDTH = 720
WIN_HEIGHT = 450
VERSION = "4.0"
INI_LIST_CSV_NAME = 'ini_list.csv'
RES_DIR = Path(__file__).parent / 'res'
ICON_PATH = RES_DIR / 'icon.ico'
INITIAL_LIST_PATH = RES_DIR / 'initialization'


class RandomSelectorApp:
    def __init__(self, master):
        self.master = master
        self.master.title(f'随机抽号 v{VERSION}')
        self.master.geometry(f'{WIN_WIDTH}x{WIN_HEIGHT}')
        self.master.iconbitmap(ICON_PATH)
        self.master.resizable(False, False)
        self.master.protocol('WM_DELETE_WINDOW', self._on_closing)

        # ============================================================ 实例变量
        self.class_list = []
        self.listbox_widget = None  # Will be assigned in the settings panel
        self.entry_widget = None  # Will be assigned in the settings panel
        self.file_path_var = ''
        self.file_text_var = ''

        self._load_initial_list()
        self._create_widgets()
        self._setup_key_bindings()

    def _load_initial_list(self):
        """检查并处理 ini_list.csv 文件，加载初始名单。"""
        if not os.path.exists(INI_LIST_CSV_NAME):
            shutil.copyfile(INITIAL_LIST_PATH, INI_LIST_CSV_NAME)

        try:
            with open(INI_LIST_CSV_NAME, 'r', newline='') as initial_list:
                ini_reader = csv.reader(initial_list)
                for row in ini_reader:
                    if row and row[0].strip():
                        self.class_list.append(row[0].strip())
        except UnicodeDecodeError:
            messagebox.showerror("错误", "无法读取初始名单文件，请确保它是正确编码的CSV文件。")
        except Exception as e:
            messagebox.showerror("错误", f"加载初始名单时发生错误: {e}")

    def _create_widgets(self):
        """创建主界面的所有Tkinter控件。"""
        # ============================================================ 主页面
        title_label = Label(self.master, text='随机抽号小程序', font=('黑体', 15), height=2, justify='center',
                            anchor='center')
        instruction_label = Label(self.master, text='''
按下<>内的快捷键实现对应功能：
    <=> 抽取      <Delete> 清空      <F12> 帮助（更多快捷键见此）

操作说明：
 1、点击“抽签”按钮即可开始抽签，每点击一次抽一个人。
 2、显示窗内容可复制，每次抽签都不会自动清理上一次内容；点“清空”即可
    清空显示窗中的所有内容。
 3、查看抽取列表、更改名单等，请在选项中查看或设置；更多说明详见帮助。
        ''', font=('宋体', 13), height=10, justify="left", anchor='center')

        self.output_text_area = Text(self.master, height=10, font=('宋体', 12))

        begin_button = Button(self.master, text='抽签', fg='green', width=6, height=1, cursor='hand2',
                              command=self._begin_draw)
        clear_button = Button(self.master, text='清空', fg='red', width=6, height=1, cursor='hand2',
                              command=self._clear_text_area)

        line1 = Label(self.master, text='')
        line2 = Label(self.master, text='')

        title_label.pack()
        instruction_label.pack()
        line1.pack()
        begin_button.place(relx=0.4, rely=0.48)
        clear_button.place(relx=0.53, rely=0.48)
        line2.pack()
        self.output_text_area.pack()

        self._create_menu_bar()

    def _create_menu_bar(self):
        """创建应用程序的菜单栏。"""
        menubar = tk.Menu(self.master)

        file_menu = Menu(menubar, tearoff=0)
        option_menu = Menu(menubar, tearoff=0)
        help_menu = Menu(menubar, tearoff=0)

        menubar.add_cascade(label='文件', menu=file_menu)
        menubar.add_cascade(label='选项', menu=option_menu)
        menubar.add_cascade(label='帮助', menu=help_menu)

        file_menu.add_command(label='新建', command=self._new_file, accelerator='Ctrl+N')
        file_menu.add_command(label='保存', command=self._save_file, accelerator='Ctrl+S')
        file_menu.add_separator()
        file_menu.add_command(label='退出', command=self._on_closing, accelerator='Alt+F4')

        option_menu.add_command(label='设置', command=self._open_settings_panel, accelerator='F8')
        option_menu.add_command(label='显示名单', command=self._show_list_popup, accelerator='F3')
        option_menu.add_command(label='更新日志', command=self._show_update_log)
        option_menu.add_command(label='版权信息', command=self._show_copyright_info)

        help_menu.add_command(label='使用说明', command=self._show_help_manual, accelerator='F12')

        self.master.config(menu=menubar)

    def _setup_key_bindings(self):
        """设置主窗口的快捷键绑定。"""
        self.master.bind('<equal>', self._begin_draw)
        self.master.bind('<Delete>', self._clear_text_area)
        self.master.bind('<F8>', self._open_settings_panel)
        self.master.bind('<F3>', self._show_list_popup)
        self.master.bind('<F12>', self._show_help_manual)
        self.master.bind('<Control-n>', self._new_file)
        self.master.bind('<Control-s>', self._save_file)
        self.master.bind('<Alt-F4>', self._on_closing)

    def _on_closing(self, x=0):
        """处理窗口关闭事件，询问用户是否确认退出。"""
        if_q = tk.messagebox.askquestion(title='提示', message='您确定要退出吗？')
        if if_q == 'yes':
            self.master.destroy()
        else:
            return 0

    def _void_function(self):
        """一个空指令，用于阻止窗口关闭事件。"""
        pass

    def _begin_draw(self, x=0):
        """从名单中随机抽取一个人名并显示在文本框中。"""
        try:
            if not self.class_list:
                raise ValueError("名单为空！")
            index = random.randint(0, len(self.class_list) - 1)
        except ValueError:
            tk.messagebox.showwarning(title='警告', message='名单为空！')
        else:
            self.output_text_area.insert('end', self.class_list[index])
            self.output_text_area.insert('end', '\t')

    def _clear_text_area(self, x=0):
        """清空文本显示区域的所有内容。"""
        self.output_text_area.delete('1.0', 'end')

    def _initialize_settings_listbox(self, x=0):
        """刷新设置面板中的名单列表。"""
        if self.listbox_widget:
            self.listbox_widget.delete(0, END)
            for item in self.class_list:
                self.listbox_widget.insert(END, item)

    def _clear_class_list_action(self, x=0):
        """清空名单列表中的所有项，并提供恢复提示。"""
        if_clr = messagebox.askyesno(title='清空',
                                     message='您确定要清空名单内所有项吗？\n注意：再次打开软件即可恢复为预置名单！')
        if if_clr:
            if self.listbox_widget:
                self.listbox_widget.delete(0, END)
            self.class_list.clear()
            messagebox.showinfo(title='清空', message='您已清空名单内所有项！')

    def _insert_name_action(self, x=0):
        """在名单中添加一个人名。"""
        if not self.entry_widget: return
        name_to_add = self.entry_widget.get().strip()
        if name_to_add:
            if self.listbox_widget:
                if self.listbox_widget.curselection() == ():
                    self.listbox_widget.insert(self.listbox_widget.size(), name_to_add)
                else:
                    self.listbox_widget.insert(self.listbox_widget.curselection(), name_to_add)
            self.class_list.append(name_to_add)
            messagebox.showinfo(title='添加提示', message=f'您已成功添加“{name_to_add}”！')
        else:
            messagebox.showinfo(title='添加提示', message='请在文本框中输入要添加的人名。')

    def _modify_name_action(self, x=0):
        """修改名单中选定的人名。"""
        if not self.listbox_widget or not self.entry_widget: return
        selected_indices = self.listbox_widget.curselection()
        if not selected_indices:
            messagebox.showinfo(title='修改提示', message='请在左侧选择即将修改的人名。')
            return

        selected_index = selected_indices[0]
        new_name = self.entry_widget.get().strip()

        if not new_name:
            messagebox.showinfo(title='修改提示', message='请在文本框中输入新的人名。')
            return

        existing_name = self.class_list[selected_index]
        self.listbox_widget.delete(selected_index)
        self.listbox_widget.insert(selected_index, new_name)
        self.class_list[selected_index] = new_name
        tk.messagebox.showinfo(title='修改提示', message=f'您已成功将“{existing_name}”修改为“{new_name}”！')

    def _delete_name_action(self, x=0):
        """从名单中删除选定的人名。"""
        if not self.listbox_widget: return
        selected_indices = self.listbox_widget.curselection()
        if not selected_indices:
            messagebox.showinfo(title='删除提示', message='请在左侧选择要去除的人名。')
            return

        selected_index = selected_indices[0]
        name_to_delete = self.class_list[selected_index]

        self.listbox_widget.delete(selected_index)
        self.class_list.pop(selected_index)
        tk.messagebox.showinfo(title='删除提示', message=f'您已成功删除“{name_to_delete}”！')

    def _import_list_action(self, x=0):
        """从CSV文件导入名单。"""
        if not self.listbox_widget: return
        if_con = messagebox.askyesno(title='注意',
                                     message='1. 导入的CSV文件必须是一列数据。\n2. 导入新列表，将会在原有列表的基础上添加，\n    您确定要继续吗？')
        if if_con:
            file_path = filedialog.askopenfilename(filetypes=[("CSV文件", "*.csv")])
            if file_path:
                try:
                    with open(file_path, 'r', newline='') as file:
                        reader = csv.reader(file)
                        for row in reader:
                            if row and row[0].strip():
                                name = row[0].strip()
                                self.listbox_widget.insert(END, name)
                                self.class_list.append(name)
                    messagebox.showinfo(title='导入成功', message='名单已成功导入！')
                except UnicodeDecodeError:
                    messagebox.showerror(title='导入失败', message='文件编码不正确，请确保是正确编码的CSV文件。')
                except Exception as e:
                    messagebox.showerror(title='导入失败', message=f'导入过程中发生错误: {e}')

    def _open_settings_panel(self, x=0):
        """打开设置面板，用于管理名单。"""
        settings_root = Toplevel(self.master)
        settings_root.title('设置')
        settings_root.geometry('320x380')
        settings_root.transient(self.master)
        settings_root.protocol('WM_DELETE_WINDOW', self._void_function)

        t_label = Label(settings_root, text='抽点列表设置', height=2, font=('黑体', 14))
        t_label.pack()

        frame1 = Frame(settings_root, relief=tk.GROOVE)
        frame2 = Frame(settings_root, relief=tk.GROOVE)
        frame1.place(relx=0.02, rely=0.12)
        frame2.place(relx=0.51, rely=0.12)

        t_listbox1_label = Label(frame1, text='名单', font=('黑体', 13))
        t_listbox1_label.pack()
        self.listbox_widget = Listbox(frame1, font=('宋体', 12), width=18, height=13)
        self.listbox_widget.pack()
        self._initialize_settings_listbox()

        t_entry_label = Label(frame2, text='文本框', font=('黑体', 13))
        t_entry_label.pack()
        self.entry_widget = Entry(frame2)
        self.entry_widget.pack()

        t_b4_btn = Label(frame2, text=' ')
        t_b4_btn.pack()

        btn_refresh = Button(frame2, text='刷新 (F5)', fg='#FF50FF', cursor='hand2', width=16,
                             command=self._initialize_settings_listbox)
        btn_add = Button(frame2, text='添加 (F1)', fg='green', cursor='hand2', width=16,
                         command=self._insert_name_action)
        btn_modify = Button(frame2, text='修改 (F2)', fg='blue', cursor='hand2', width=16,
                            command=self._modify_name_action)
        btn_delete = Button(frame2, text='删除 (F10)', fg='red', cursor='hand2', width=16,
                            command=self._delete_name_action)
        btn_clear_list = Button(frame2, text='清空', fg='purple', cursor='hand2', width=16,
                                command=self._clear_class_list_action)
        btn_import = Button(frame2, text='导入列表', fg='#FFAA00', cursor='hand2', width=16,
                            command=self._import_list_action)

        btn_refresh.pack()
        btn_add.pack()
        btn_modify.pack()
        btn_delete.pack()
        btn_clear_list.pack()
        btn_import.pack()

        ok_set = Button(settings_root, text='完成设置', cursor='hand2', bd=3, height=2, width=10,
                        font=('黑体', 12), fg='blue', command=lambda: settings_root.destroy())
        ok_set.place(relx=0.35, rely=0.8)

        # 绑定快捷键到设置面板
        settings_root.bind("<F1>", self._insert_name_action)
        settings_root.bind("<F10>", self._delete_name_action)
        settings_root.bind("<F2>", self._modify_name_action)
        settings_root.bind("<F5>", self._initialize_settings_listbox)

        settings_root.mainloop()

    def _new_file(self, x=0):
        """创建一个新文件，用户可以选择文件名和扩展名。"""
        messagebox.showinfo(title='新建',
                            message='温馨提示：\n新建文件时，请输入文件名及其扩展名，\n若未输入则默认为*.txt')
        self.file_path_var = filedialog.asksaveasfilename(title='新建文件',
                                                          filetypes=[('文本文档（推荐）', '*.txt'),
                                                                     ('新版Word文档', '*.docx'),
                                                                     ('新版Excel文档', '*.xlsx'), ('全部文件', '*')],
                                                          defaultextension='*.txt')
        self.file_text_var = ''
        if self.file_path_var:
            try:
                with open(file=self.file_path_var, mode='w') as file:
                    file.write(self.file_text_var)
            except Exception:
                messagebox.showerror(title='新建', message='新建失败！')
            else:
                messagebox.showinfo(title='新建', message='新建成功！')

    def _save_file(self, x=0):
        """保存文本显示区域的内容到文件。"""
        self.file_path_var = filedialog.asksaveasfilename(title='保存文件',
                                                          filetypes=[('文本文档（推荐）', '*.txt'),
                                                                     ('新版Word文档', '*.docx'),
                                                                     ('新版Excel文档', '*.xlsx'), ('全部文件', '*')])
        self.file_text_var = self.output_text_area.get('1.0', 'end-1c')
        if self.file_path_var:
            try:
                with open(file=self.file_path_var, mode='a+') as file:
                    file.write(self.file_text_var)
            except Exception:
                messagebox.showerror(title='保存', message='取消保存或保存失败！')
            else:
                messagebox.showinfo(title='保存', message='输出内容已成功保存！')

    def _show_list_popup(self, x=0):
        """显示一个弹出窗口，展示当前所有名单。"""
        show_class_window = Toplevel(self.master)
        show_class_window.title('名单')
        show_class_window.transient(self.master)

        class_listbox = Listbox(show_class_window, font=('宋体', 13), width=30, height=30)
        for item in self.class_list:
            class_listbox.insert('end', item)

        sl_label = Label(show_class_window, text='抽取名单', font=('黑体', 14), height=3, justify="center",
                         anchor='center')
        sl_tell = Label(show_class_window,
                        text='\n    此列表仅供查看抽取名单，若要修改请前往“设置”。\n编辑预置名单入口：软件所在目录的“ini_list.csv”文件\n',
                        font=('宋体', 9), justify='left', anchor='center')

        sl_label.pack()
        class_listbox.pack()
        sl_tell.pack()

    def _show_update_log(self):
        """显示更新日志。"""
        popup_log = Toplevel(self.master)
        popup_log.title("更新日志")
        popup_log.transient(self.master)

        scrollbar = tk.Scrollbar(popup_log)
        scrollbar.pack(side='right', fill='y')

        textbox = tk.Text(popup_log, height=30, width=100, font=('宋体', 13), yscrollcommand=scrollbar.set)
        textbox.insert('end', """2022.01.23 22:39 | Lottery v1.0.220123 beta:
推出简易版小程序，具有老虎机一般的简单抽号功能

2022.03.22 23:43 | Lottery - name v1.2.1:
采用GUI交互式设计，添加了名单修改工具，增设了常用功能的快捷键

2022.05.29 12:27 | Lottery - name v2.0:
完善了名单修改工具，增设了新建文件、保存文件、显示抽取列表等功能

2022.05.29 19:05 | Lottery - name v2.1:
完善了上个版本添加的功能，增加了帮助栏，继续增设了更多快捷键

2022.10.02 0:16 | 随机点名小程序 v2.2:
完善了既有功能，优化诸多场景下的用户体验，继续增设了更多快捷键

2024.07.04 10:30 | 随机抽号小程序 v3.0:
增设批量导入名单的功能，重设了初始名单的读取逻辑，更新了版权信息中的联系方式

2024.07.31 23:58 | 随机抽号小程序 v3.1:
将初始名单设为自动识别模式，若计算机上没有初始名单，则在程序目录下新建初始名单

2024.08.01 11:40 | 随机抽号小程序 v3.2:
增设自动初始化功能，修复了系统Temp文件夹内名单缺失等问题，加入了更新日志

2025.09.19 22:40 | 随机抽号小程序 v4.0:
采用全新架构（见main_new.py）并修复了读取CSV文件时的bug，解决了文字不符合UTF-8编码时读取失败的问题""")
        textbox.config(state='disabled')
        textbox.pack(side='left', fill='both')
        scrollbar.config(command=textbox.yview)

    def _show_copyright_info(self):
        """显示版权信息和联系方式。"""
        popup = Toplevel(self.master)
        popup.title("版权信息与联系方式")
        popup.transient(self.master)

        cprt_label = tk.Label(popup, text='\n\n开发者\n\nH.Jupiter.Lyr\n\n\n版权所有者\n\nH.Jupiter.Lyr\n\n\n邮箱',
                              font=('Times New Roman', 14))
        email_label = tk.Label(popup, text='jupiterlyr@foxmail.com', fg='blue', cursor='hand2',
                               font=('Times New Roman', 14))
        close_button = tk.Button(popup, text='关闭', font=('黑体', 14), cursor='hand2', width=8,
                                 command=lambda: popup.destroy())

        cprt_label.pack(padx=30)
        email_label.pack(padx=30, pady=15)
        email_label.bind("<Button-1>", lambda event: webbrowser.open(f"mailto:jupiterlyr@foxmail.com"))
        close_button.pack(pady=40)

    def _show_help_manual(self, x=0):
        """显示帮助手册，包含快捷键和功能介绍。"""
        h_window = Toplevel(self.master, bg='white')
        h_window.title('帮助')
        h_window.transient(self.master)

        h_title = Label(h_window, text='帮助手册', font=('黑体', 15), bg='white')
        h_quit = Button(h_window, text='我知道了', width=10, height=1, bg='white', cursor='hand2',
                        command=h_window.destroy)
        h_text = Label(h_window, text="""
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

        h_title.pack()
        h_text.pack()
        h_quit.pack()
        h_window.resizable(False, False)
        h_window.mainloop()

    def start(self):
        """启动Tkinter主循环。"""
        self.master.mainloop()


if __name__ == "__main__":
    root = tk.Tk()
    app = RandomSelectorApp(root)
    app.start()
