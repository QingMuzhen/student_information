import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import sqlite3
import openpyxl
import hashlib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# 确保中文显示正常
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False  # 正确显示负号

# 颜色和字体常量 - 采用更现代的配色方案
BG_COLOR = 'white'
SIDEBAR_COLOR = 'white'
BTN_BG_COLOR = 'white'
BTN_HOVER_COLOR = 'white'
BTN_FG_COLOR = 'black'
FONT = ('Microsoft YaHei', 10)
HEADER_FONT = ('Microsoft YaHei', 12, 'bold')

# 加密密钥的哈希值（原密钥qmzyyds的哈希）
SECRET_KEY_HASH = hashlib.sha256(b'qmzyyds').hexdigest()


class StudentSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("学生信息管理系统")
        self.center_window(self.root, 500, 400)  # 初始窗口大小调整
        self.root.resizable(False, False)
        self.root.configure(bg=BG_COLOR)

        # 设置样式
        self.style = ttk.Style()
        self._setup_styles()

        self.conn = sqlite3.connect('students_encrypted.db')
        self.cursor = self.conn.cursor()
        self._create_tables()

        self.current_user = None
        self.dynamic_fields = {}
        self.dynamic_field_entries = []
        self.exam_names = []
        self._load_exam_names()

        # 新增变量用于存储当前查询表格和滚动条
        self.current_tree = None
        self.current_scrollbar = None
        self.canvas = None

        self.create_login_page()

    def _setup_styles(self):
        """设置ttk样式"""
        self.style.configure('TButton', font=FONT, background=BTN_BG_COLOR, foreground=BTN_FG_COLOR)
        self.style.map('TButton', background=[('active', BTN_HOVER_COLOR)])
        self.style.configure('TLabel', font=FONT, background=BG_COLOR)
        self.style.configure('TEntry', font=FONT, padding=5)
        self.style.configure('TCombobox', font=FONT, padding=5)
        self.style.configure('Treeview', font=FONT, rowheight=25)
        self.style.configure('Treeview.Heading', font=HEADER_FONT, background='#E0E0E0')

    def center_window(self, window, width, height):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    def _create_tables(self):
        """创建数据库表"""
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY, username TEXT UNIQUE, password TEXT)''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY, name TEXT, chinese TEXT, math TEXT, english TEXT, exam_name TEXT)''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS student_fields (
            id INTEGER PRIMARY KEY, student_id INTEGER, field_name TEXT, field_value TEXT,
            FOREIGN KEY (student_id) REFERENCES students(id))''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS operations (
            id INTEGER PRIMARY KEY, username TEXT, operation_type TEXT, operation_time DATETIME)''')
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS exams (
            id INTEGER PRIMARY KEY, exam_name TEXT UNIQUE)''')
        self.conn.commit()

    def _load_exam_names(self):
        self.cursor.execute("SELECT exam_name FROM exams")
        self.exam_names = [row[0] for row in self.cursor.fetchall()]

    def create_login_page(self):
        """创建登录页面"""
        self.clear_window()

        # 创建登录卡片效果
        card_frame = tk.Frame(self.root, bg='white', bd=1, relief=tk.SOLID)
        card_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER, width=350, height=320)

        title_label = tk.Label(card_frame, text="学生信息管理系统", font=('Microsoft YaHei', 16, 'bold'),
                               bg='white')
        title_label.pack(pady=20)

        frame = tk.Frame(card_frame, bg='white')
        frame.pack(fill=tk.BOTH, expand=True, padx=30)

        tk.Label(frame, text="用户名:", font=FONT, bg='white').pack(anchor=tk.W, pady=5)
        username_entry = tk.Entry(frame, font=FONT, bd=1, relief=tk.SOLID, highlightthickness=1)
        username_entry.pack(fill=tk.X, pady=5)
        username_entry.focus()

        tk.Label(frame, text="密码:", font=FONT, bg='white').pack(anchor=tk.W, pady=5)
        password_entry = tk.Entry(frame, font=FONT, show='*', bd=1, relief=tk.SOLID, highlightthickness=1)
        password_entry.pack(fill=tk.X, pady=5)

        btn_frame = tk.Frame(card_frame, bg='white')
        btn_frame.pack(pady=20, fill=tk.X, padx=30)

        tk.Button(btn_frame, text="注册", command=self.create_register_page,
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT,
                  padx=5).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(btn_frame, text="登录",
                  command=lambda: self._handle_login(username_entry, password_entry),
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT,
                  padx=5).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(btn_frame, text="退出", command=self.root.destroy,
                  bg='#E74C3C', fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT,
                  padx=5).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 账号管理按钮
        tk.Button(card_frame, text="账号管理", command=self.create_account_management_page,
                  bg='#F39C12', fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.BOTTOM, pady=10, fill=tk.X, padx=30)

        # 绑定回车键登录
        self.root.bind('<Return>', lambda event: self._handle_login(username_entry, password_entry))

    def create_account_management_page(self):
        """创建账号管理页面"""
        self.clear_window()

        self.root.geometry("600x500")
        self.center_window(self.root, 600, 500)

        title_label = tk.Label(self.root, text="账号管理", font=('Microsoft YaHei', 16, 'bold'), bg=BG_COLOR)
        title_label.pack(pady=20)

        btn_frame = tk.Frame(self.root, bg=BG_COLOR)
        btn_frame.pack(pady=10, fill=tk.X, padx=20)

        tk.Button(btn_frame, text="返回登录", command=self.create_login_page,
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        tk.Button(btn_frame, text="刷新", command=self.create_account_management_page,
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # 创建表格框架
        table_frame = tk.Frame(self.root, bg=BG_COLOR)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 创建表格
        columns = ['username', 'operation']
        tree = ttk.Treeview(table_frame, show='headings', columns=columns)

        for col in columns:
            width = 300 if col == 'username' else 150
            tree.column(col, width=width, anchor=tk.CENTER)
            tree.heading(col, text={'username': '用户名', 'operation': '操作'}[col])

        # 添加滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 加载账号数据
        self.cursor.execute("SELECT id, username FROM users")
        users = self.cursor.fetchall()

        for user_id, username in users:
            tree.insert("", tk.END, values=(username, "修改 | 删除"), iid=user_id)

        tree.bind("<Button-1>", lambda event: self._handle_account_tree_click(event, tree))

    def _handle_account_tree_click(self, event, tree):
        """处理账号表格点击事件"""
        region = tree.identify_region(event.x, event.y)
        if region == "cell":
            row_id = tree.identify_row(event.y)
            column = tree.identify_column(event.x)

            if column == "#2":  # 操作列
                x, y, width, height = tree.bbox(row_id, column)
                click_offset = event.x - x
                if click_offset < width / 2:
                    # 修改账号
                    self._modify_account(row_id)
                else:
                    # 删除账号
                    self._delete_account(row_id)

    def _verify_original_password_and_key(self, user_id):
        """验证原密码和密钥"""
        result = False
        self.cursor.execute("SELECT username FROM users WHERE id=?", (user_id,))
        username = self.cursor.fetchone()[0]

        password_window = tk.Toplevel(self.root)
        password_window.title("验证原密码和密钥")
        self.center_window(password_window, 400, 300)
        password_window.resizable(False, False)
        password_window.grab_set()
        password_window.configure(bg=BG_COLOR)

        frame = tk.Frame(password_window, bg=BG_COLOR)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(frame, text=f"验证 {username} 的密码:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        password_entry = tk.Entry(frame, font=FONT, show='*', bd=1, relief=tk.SOLID)
        password_entry.pack(fill=tk.X, pady=5)

        tk.Label(frame, text="密钥:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        key_entry = tk.Entry(frame, font=FONT, show='*', bd=1, relief=tk.SOLID)
        key_entry.pack(fill=tk.X, pady=5)

        def check_password_and_key():
            nonlocal result
            password = password_entry.get().strip()
            key = key_entry.get().strip()

            # 验证密钥
            key_hash = hashlib.sha256(key.encode()).hexdigest()
            if key_hash != SECRET_KEY_HASH:
                messagebox.showerror("错误", "密钥错误，请重新输入。")
                return

            self.cursor.execute("SELECT * FROM users WHERE id=? AND password=?",
                                (user_id, password))
            if self.cursor.fetchone():
                result = True
                password_window.destroy()
            else:
                messagebox.showerror("错误", "密码错误，请重新输入。")

        btn_frame = tk.Frame(frame, bg=BG_COLOR)
        btn_frame.pack(pady=15, fill=tk.X)

        tk.Button(btn_frame, text="确定", command=check_password_and_key,
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        tk.Button(btn_frame, text="取消", command=password_window.destroy,
                  bg='#E74C3C', fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.RIGHT, padx=10, fill=tk.X, expand=True)

        password_window.bind('<Return>', lambda e: check_password_and_key())
        password_window.wait_window()

        return result

    def _modify_account(self, user_id):
        """修改账号"""
        if not self._verify_original_password_and_key(user_id):
            return

        self.cursor.execute("SELECT username FROM users WHERE id=?", (user_id,))
        username = self.cursor.fetchone()[0]

        modify_window = tk.Toplevel(self.root)
        modify_window.title("修改账号")
        self.center_window(modify_window, 400, 350)
        modify_window.resizable(False, False)
        modify_window.grab_set()
        modify_window.configure(bg=BG_COLOR)

        frame = tk.Frame(modify_window, bg=BG_COLOR)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(frame, text="用户名:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        username_entry = tk.Entry(frame, font=FONT, bd=1, relief=tk.SOLID)
        username_entry.insert(0, username)
        username_entry.pack(fill=tk.X, pady=5)

        tk.Label(frame, text="新密码:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        password_entry = tk.Entry(frame, font=FONT, show='*', bd=1, relief=tk.SOLID)
        password_entry.pack(fill=tk.X, pady=5)

        tk.Label(frame, text="确认新密码:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        confirm_entry = tk.Entry(frame, font=FONT, show='*', bd=1, relief=tk.SOLID)
        confirm_entry.pack(fill=tk.X, pady=5)

        def update_account():
            new_username = username_entry.get().strip()
            password = password_entry.get().strip()
            confirm = confirm_entry.get().strip()

            if not new_username:
                messagebox.showerror("错误", "用户名不能为空，请输入有效的用户名。")
                return

            if password and password != confirm:
                messagebox.showerror("错误", "两次输入的密码不一致，请重新输入。")
                return

            try:
                if password:
                    self.cursor.execute("UPDATE users SET username=?, password=? WHERE id=?",
                                        (new_username, password, user_id))
                else:
                    self.cursor.execute("UPDATE users SET username=? WHERE id=?",
                                        (new_username, user_id))
                self.conn.commit()
                messagebox.showinfo("成功", "账号信息已更新。")
                modify_window.destroy()
                self.create_account_management_page()
            except sqlite3.IntegrityError:
                messagebox.showerror("错误", "用户名已存在，请选择其他用户名。")

        btn_frame = tk.Frame(frame, bg=BG_COLOR)
        btn_frame.pack(pady=20, fill=tk.X)

        tk.Button(btn_frame, text="更新", command=update_account,
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        tk.Button(btn_frame, text="取消", command=modify_window.destroy,
                  bg='#E74C3C', fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.RIGHT, padx=10, fill=tk.X, expand=True)

    def _delete_account(self, user_id):
        """删除账号"""
        if not self._verify_original_password_and_key(user_id):
            return

        self.cursor.execute("SELECT username FROM users WHERE id=?", (user_id,))
        username = self.cursor.fetchone()[0]

        if messagebox.askyesno("确认", f"确定要删除账号 {username} 吗？此操作不可恢复。"):
            try:
                self.cursor.execute("DELETE FROM users WHERE id=?", (user_id,))
                self.conn.commit()
                messagebox.showinfo("成功", "账号已删除。")
                self.create_account_management_page()
            except Exception as e:
                messagebox.showerror("错误", f"删除失败: {str(e)}，请稍后再试。")

    def create_register_page(self):
        """创建注册页面"""
        self.clear_window()

        self.root.geometry("500x500")
        self.center_window(self.root, 500, 500)

        title_label = tk.Label(self.root, text="用户注册", font=('Microsoft YaHei', 16, 'bold'), bg=BG_COLOR)
        title_label.pack(pady=20)

        frame = tk.Frame(self.root, bg=BG_COLOR)
        frame.pack(fill=tk.BOTH, expand=True, padx=50)

        tk.Label(frame, text="用户名:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        username_entry = tk.Entry(frame, font=FONT, bd=1, relief=tk.SOLID)
        username_entry.pack(fill=tk.X, pady=5)
        username_entry.focus()

        tk.Label(frame, text="密码:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        password_entry = tk.Entry(frame, font=FONT, show='*', bd=1, relief=tk.SOLID)
        password_entry.pack(fill=tk.X, pady=5)

        tk.Label(frame, text="确认密码:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        confirm_entry = tk.Entry(frame, font=FONT, show='*', bd=1, relief=tk.SOLID)
        confirm_entry.pack(fill=tk.X, pady=5)

        tk.Label(frame, text="密钥:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        key_entry = tk.Entry(frame, font=FONT, show='*', bd=1, relief=tk.SOLID)
        key_entry.pack(fill=tk.X, pady=5)

        btn_frame = tk.Frame(self.root, bg=BG_COLOR)
        btn_frame.pack(pady=30, fill=tk.X, padx=50)

        tk.Button(btn_frame, text="返回登录", command=self.create_login_page,
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        tk.Button(btn_frame, text="注册",
                  command=lambda: self._handle_register(username_entry, password_entry, confirm_entry,
                                                        key_entry),
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # 绑定回车键注册
        self.root.bind('<Return>',
                       lambda event: self._handle_register(username_entry, password_entry, confirm_entry, key_entry))

    def _handle_login(self, username_entry, password_entry):
        """处理登录逻辑"""
        username = username_entry.get().strip()
        password = password_entry.get().strip()

        if not username or not password:
            messagebox.showerror("错误", "用户名和密码不能为空，请输入有效的信息。")
            return

        self.cursor.execute("SELECT * FROM users WHERE username=? AND password=?", (username, password))
        if self.cursor.fetchone():
            self.current_user = username
            self._create_welcome_page()
            self.root.attributes('-fullscreen', True)  # 全屏显示
        else:
            messagebox.showerror("错误", "用户名或密码错误，请重新输入。")

    def _handle_register(self, username_entry, password_entry, confirm_entry, key_entry):
        """处理注册逻辑"""
        username = username_entry.get().strip()
        password = password_entry.get().strip()
        confirm = confirm_entry.get().strip()
        key = key_entry.get().strip()

        if not username or not password:
            messagebox.showerror("错误", "用户名和密码不能为空，请输入有效的信息。")
            return

        if password != confirm:
            messagebox.showerror("错误", "两次输入的密码不一致，请重新输入。")
            return

        # 验证密钥（使用哈希值比较）
        key_hash = hashlib.sha256(key.encode()).hexdigest()
        if key_hash != SECRET_KEY_HASH:
            messagebox.showerror("错误", "密钥输入错误，请重新输入。")
            return

        try:
            self.cursor.execute("INSERT INTO users (username, password) VALUES (?, ?)", (username, password))
            self.conn.commit()
            messagebox.showinfo("成功", "注册成功，请登录。")
            self.create_login_page()
        except sqlite3.IntegrityError:
            messagebox.showerror("错误", "用户名已存在，请选择其他用户名。")

    def _create_welcome_page(self):
        """创建欢迎页面"""
        self.clear_window()
        self.root.attributes('-fullscreen', True)  # 全屏显示
        self.root.resizable(True, True)
        self.root.configure(bg=BG_COLOR)

        # 顶部状态栏
        status_frame = tk.Frame(self.root, bg='#f0f0f0', height=40)
        status_frame.pack(fill=tk.X)
        status_frame.pack_propagate(False)

        tk.Label(status_frame, text=f"当前用户: {self.current_user}", font=FONT, bg='#f0f0f0').pack(side=tk.LEFT,
                                                                                                    padx=20, pady=10)

        # 添加退出全屏按钮
        tk.Button(status_frame, text="退出全屏", command=lambda: self.root.attributes('-fullscreen', False),
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT, padx=10).pack(side=tk.RIGHT, padx=10)

        tk.Button(status_frame, text="退出登录", command=self._logout,
                  bg='#E74C3C', fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT, padx=10).pack(
            side=tk.RIGHT, padx=10)

        # 左侧导航
        nav_frame = tk.Frame(self.root, width=180, bg=SIDEBAR_COLOR)
        nav_frame.pack(side=tk.LEFT, fill=tk.Y)
        nav_frame.pack_propagate(False)

        # 导航标题
        tk.Label(nav_frame, text="功能导航", font=('Microsoft YaHei', 12, 'bold'),
                 bg=SIDEBAR_COLOR, fg='black').pack(pady=20)

        # 导航按钮
        nav_buttons = [
            ("录入信息", self._show_input_page),
            ("查询信息", self._show_query_page),
            ("考试管理", self._show_exam_management_page),
            ("数据统计", self._show_statistics_page),
            ("导出数据", self._export_data)
        ]

        for text, command in nav_buttons:
            btn = tk.Button(nav_frame, text=text, command=command,
                            bg=SIDEBAR_COLOR, fg='black', font=('Microsoft YaHei', 11),
                            relief=tk.FLAT, justify=tk.LEFT, anchor=tk.W, padx=20, pady=15)
            btn.pack(fill=tk.X)
            btn.bind("<Enter>", lambda e, b=btn: b.config(bg='#34495E'))
            btn.bind("<Leave>", lambda e, b=btn: b.config(bg=SIDEBAR_COLOR))

        # 右侧内容区
        self.content_frame = tk.Frame(self.root, bg=BG_COLOR)
        self.content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        welcome_label = tk.Label(self.content_frame, text="欢迎使用学生信息管理系统！",
                                 font=('Microsoft YaHei', 24, 'bold'), bg=BG_COLOR)
        welcome_label.pack(pady=100)

        # 系统信息
        info_text = """
        该系统用于管理学生考试成绩信息，支持：
        - 录入和查询学生成绩
        - 管理不同考试
        - 数据统计与分析
        - 导出数据到Excel

        请从左侧导航栏选择需要的功能。
        """
        info_label = tk.Label(self.content_frame, text=info_text, font=('Microsoft YaHei', 12),
                              bg=BG_COLOR, justify=tk.LEFT)
        info_label.pack(pady=20, padx=50, anchor=tk.W)

    def _logout(self):
        """处理退出登录"""
        self.current_user = None
        self.root.attributes('-fullscreen', False)  # 退出全屏
        self.root.resizable(False, False)
        self.create_login_page()

    def _show_input_page(self):
        """显示录入页面"""
        self._clear_content()
        frame = tk.Frame(self.content_frame, bg=BG_COLOR, padx=30, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # 页面标题
        tk.Label(frame, text="录入学生信息", font=('Microsoft YaHei', 18, 'bold'), bg=BG_COLOR).pack(pady=20)

        # 考试名称选择
        tk.Label(frame, text="考试名称:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=5)
        exam_name_var = tk.StringVar()
        exam_name_combobox = ttk.Combobox(frame, textvariable=exam_name_var, values=self.exam_names, width=50)
        exam_name_combobox.pack(pady=5)
        exam_name_combobox.set(self.exam_names[0] if self.exam_names else "")

        add_exam_btn = tk.Button(frame, text="创建考试", command=lambda: self._create_exam(exam_name_var),
                                 bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT)
        add_exam_btn.pack(pady=10)

        # 基本信息部分 - 使用卡片效果
        basic_frame = tk.LabelFrame(frame, text="基本信息", font=HEADER_FONT, bg=BG_COLOR, bd=2, relief=tk.SOLID)
        basic_frame.pack(fill=tk.X, padx=10, pady=15)

        # 姓名行
        name_frame = tk.Frame(basic_frame, bg=BG_COLOR)
        name_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(name_frame, text="姓名:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        self.name_entry = tk.Entry(name_frame, font=FONT, bd=1, relief=tk.SOLID)
        self.name_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.name_entry.focus()

        # 成绩行1
        score_frame1 = tk.Frame(basic_frame, bg=BG_COLOR)
        score_frame1.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(score_frame1, text="语文:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        self.chinese_entry = tk.Entry(score_frame1, font=FONT, bd=1, relief=tk.SOLID, width=15)
        self.chinese_entry.pack(side=tk.LEFT, padx=10)

        tk.Label(score_frame1, text="数学:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        self.math_entry = tk.Entry(score_frame1, font=FONT, bd=1, relief=tk.SOLID, width=15)
        self.math_entry.pack(side=tk.LEFT, padx=10)

        tk.Label(score_frame1, text="英语:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        self.english_entry = tk.Entry(score_frame1, font=FONT, bd=1, relief=tk.SOLID, width=15)
        self.english_entry.pack(side=tk.LEFT, padx=10)

        # 动态字段部分
        dynamic_frame = tk.LabelFrame(frame, text="自定义学科", font=HEADER_FONT, bg=BG_COLOR, bd=2, relief=tk.SOLID)
        dynamic_frame.pack(fill=tk.X, padx=10, pady=15)

        # 动态字段添加区域
        add_field_frame = tk.Frame(dynamic_frame, bg=BG_COLOR)
        add_field_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(add_field_frame, text="学科名称:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        self.new_field_name = tk.Entry(add_field_frame, font=FONT, bd=1, relief=tk.SOLID, width=20)
        self.new_field_name.pack(side=tk.LEFT, padx=10)

        tk.Label(add_field_frame, text="学科分数:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        self.new_field_value = tk.Entry(add_field_frame, font=FONT, bd=1, relief=tk.SOLID, width=10)
        self.new_field_value.pack(side=tk.LEFT, padx=10)

        add_button = tk.Button(
            add_field_frame,
            text="添加",
            command=self.add_dynamic_field,
            bg=BTN_BG_COLOR,
            fg=BTN_FG_COLOR,
            font=FONT,
            relief=tk.FLAT
        )
        add_button.pack(side=tk.LEFT, padx=10)

        # 动态字段显示区域
        self.dynamic_fields_frame = tk.Frame(dynamic_frame, bg=BG_COLOR)
        self.dynamic_fields_frame.pack(fill=tk.X, padx=10, pady=10)

        # 提交按钮
        button_frame = tk.Frame(frame, bg=BG_COLOR)
        button_frame.pack(pady=30)

        submit_btn = tk.Button(
            button_frame,
            text="提交",
            command=lambda: self._submit_student(exam_name_var.get()),
            bg=BTN_BG_COLOR,
            fg=BTN_FG_COLOR,
            font=HEADER_FONT,
            padx=30,
            pady=8,
            relief=tk.FLAT
        )
        submit_btn.pack()

        # 绑定回车键到提交按钮
        frame.bind('<Return>', lambda event: self._submit_student(exam_name_var.get()))

        # 清空输入框和动态字段
        self.name_entry.delete(0, tk.END)
        self.chinese_entry.delete(0, tk.END)
        self.math_entry.delete(0, tk.END)
        self.english_entry.delete(0, tk.END)

        # 清空动态字段
        for name, frame in self.dynamic_field_entries:
            frame.destroy()

        self.dynamic_fields = {}
        self.dynamic_field_entries = []

    def _create_exam(self, exam_name_var):
        exam_name = exam_name_var.get().strip()
        if not exam_name:
            messagebox.showerror("错误", "考试名称不能为空，请输入有效的考试名称。")
            return
        try:
            self.cursor.execute("INSERT INTO exams (exam_name) VALUES (?)", (exam_name,))
            self.conn.commit()
            self._load_exam_names()
            messagebox.showinfo("成功", f"考试 {exam_name} 创建成功。")
            exam_name_var.set(exam_name)
        except sqlite3.IntegrityError:
            messagebox.showerror("错误", f"考试 {exam_name} 已存在，请选择其他考试名称。")

    def add_dynamic_field(self):
        """添加动态字段"""
        field_name = self.new_field_name.get().strip()
        field_value = self.new_field_value.get().strip()

        if not field_name:
            messagebox.showerror("错误", "字段名称不能为空，请输入有效的学科名称。")
            return

        # 检查是否已存在同名的动态字段
        if field_name in self.dynamic_fields:
            messagebox.showerror("错误", f"字段 '{field_name}' 已存在，请选择其他学科名称。")
            return

        # 验证字段值
        if field_value:
            try:
                float(field_value)
            except ValueError:
                messagebox.showerror("错误", "学科分数必须为数字，请输入有效的分数。")
                return

        # 保存字段信息
        self.dynamic_fields[field_name] = field_value

        # 创建字段显示行
        field_frame = tk.Frame(self.dynamic_fields_frame, bg=BG_COLOR)
        field_frame.pack(fill=tk.X, pady=5)

        tk.Label(field_frame, text=f"{field_name}:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        tk.Label(field_frame, text=field_value, font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)

        remove_button = tk.Button(
            field_frame,
            text="删除",
            command=lambda name=field_name, frame=field_frame: self.remove_dynamic_field(name, frame),
            bg='#E74C3C',
            fg=BTN_FG_COLOR,
            font=('Microsoft YaHei', 9),
            relief=tk.FLAT,
            padx=5
        )
        remove_button.pack(side=tk.RIGHT, padx=10)

        # 保存框架引用，用于后续删除
        self.dynamic_field_entries.append((field_name, field_frame))

        # 清空输入框
        self.new_field_name.delete(0, tk.END)
        self.new_field_value.delete(0, tk.END)
        self.new_field_name.focus()

    def remove_dynamic_field(self, field_name, frame):
        """删除动态字段"""
        if field_name in self.dynamic_fields:
            del self.dynamic_fields[field_name]

        # 从视觉上移除
        frame.destroy()

        # 更新条目列表
        self.dynamic_field_entries = [(name, f) for name, f in self.dynamic_field_entries if name != field_name]

    def _submit_student(self, exam_name):
        """提交学生信息"""
        name = self.name_entry.get().strip()
        chinese = self.chinese_entry.get().strip()
        math = self.math_entry.get().strip()
        english = self.english_entry.get().strip()

        # 验证输入
        if not name:
            messagebox.showerror("错误", "姓名不能为空，请输入学生姓名。")
            return

        # 将空白项设置为“无”
        if not chinese:
            chinese = "无"
        if not math:
            math = "无"
        if not english:
            english = "无"

        try:
            if chinese != "无":
                chinese = float(chinese)
            if math != "无":
                math = float(math)
            if english != "无":
                english = float(english)
        except ValueError:
            messagebox.showerror("错误", "成绩必须为数字，请输入有效的成绩。")
            return

        try:
            self.conn.execute("BEGIN")

            # 添加新学生
            self.cursor.execute("INSERT INTO students (name, chinese, math, english, exam_name) VALUES (?, ?, ?, ?, ?)",
                                (name, chinese, math, english, exam_name))
            student_id = self.cursor.lastrowid

            # 添加自定义字段
            for field_name, field_value in self.dynamic_fields.items():
                if field_name.strip():
                    self.cursor.execute(
                        "INSERT INTO student_fields (student_id, field_name, field_value) VALUES (?, ?, ?)",
                        (student_id, field_name.strip(), field_value.strip()))

            self.conn.commit()
            messagebox.showinfo("成功", f"学生 {name} 的信息已添加。")
            self._show_query_page()

        except sqlite3.Error as e:
            self.conn.execute("ROLLBACK")
            messagebox.showerror("错误", f"操作失败: {str(e)}，请稍后再试。")

    def _show_query_page(self):
        """显示查询页面"""
        self._clear_content()
        frame = tk.Frame(self.content_frame, bg=BG_COLOR, padx=30, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # 页面标题
        tk.Label(frame, text="查询学生信息", font=('Microsoft YaHei', 18, 'bold'), bg=BG_COLOR).pack(pady=20)

        # 控制区域
        control_frame = tk.Frame(frame, bg=BG_COLOR)
        control_frame.pack(fill=tk.X, pady=10)

        # 左侧筛选区域
        filter_frame = tk.Frame(control_frame, bg=BG_COLOR)
        filter_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 考试名称选择
        tk.Label(filter_frame, text="选择考试:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=5)
        exam_name_var = tk.StringVar()
        exam_name_combobox = ttk.Combobox(filter_frame, textvariable=exam_name_var, values=self.exam_names, width=30)
        exam_name_combobox.pack(pady=5)
        exam_name_combobox.set(self.exam_names[0] if self.exam_names else "")

        # 排序选择
        sort_options = ["总成绩从高到低", "总成绩从低到高", "语文成绩从高到低", "语文成绩从低到高", "数学成绩从高到低",
                        "数学成绩从低到高", "英语成绩从高到低", "英语成绩从低到高"]

        # 存储排序变量，供后续使用
        self.sort_var = tk.StringVar()
        self.sort_var.set(sort_options[0])

        # 排序选择框
        tk.Label(filter_frame, text="排序方式:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=5)
        sort_combobox = ttk.Combobox(filter_frame, textvariable=self.sort_var, values=sort_options, width=30)
        sort_combobox.pack(pady=5)

        # 右侧按钮区域
        btn_frame = tk.Frame(control_frame, bg=BG_COLOR)
        btn_frame.pack(side=tk.RIGHT, padx=20, pady=10)

        # 刷新按钮
        refresh_btn = tk.Button(btn_frame, text="刷新数据",
                                command=lambda: self._load_data(exam_name_var.get(), self.sort_var.get()),
                                bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT)
        refresh_btn.pack(pady=5, fill=tk.X)

        # 添加学生按钮
        add_btn = tk.Button(btn_frame, text="添加学生", command=self._show_input_page,
                            bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT)
        add_btn.pack(pady=5, fill=tk.X)

        # 考试选择变化时的处理函数
        def on_exam_change(event):
            selected_exam = exam_name_var.get()
            # 重新创建表格并加载数据
            self._create_query_table(frame, selected_exam)
            self._load_data(selected_exam, self.sort_var.get())

        # 绑定考试选择变化事件
        exam_name_combobox.bind("<<ComboboxSelected>>", on_exam_change)

        # 排序变化事件
        def on_sort_change(event):
            selected_exam = exam_name_var.get()
            self._load_data(selected_exam, self.sort_var.get())

        sort_combobox.bind("<<ComboboxSelected>>", on_sort_change)

        # 表格区域
        table_frame = tk.Frame(frame, bg=BG_COLOR)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 初始创建表格
        initial_exam = exam_name_var.get()
        self._create_query_table(table_frame, initial_exam)

        # 初始加载数据
        self._load_data(initial_exam, self.sort_var.get())

    def _create_query_table(self, parent_frame, exam_name):
        """创建查询表格（根据考试动态生成列）"""
        # 移除旧表格和滚动条
        if hasattr(self, 'tree') and self.tree:
            self.tree.destroy()

        # 查询该考试下的所有自定义学科
        self.cursor.execute("""
            SELECT DISTINCT field_name 
            FROM student_fields 
            JOIN students ON student_fields.student_id = students.id 
            WHERE students.exam_name = ?
        """, (exam_name,))
        custom_fields = [row[0] for row in self.cursor.fetchall()]

        columns = ['id', 'name', 'chinese', 'math', 'english'] + custom_fields + ['operation']
        self.tree = ttk.Treeview(parent_frame, show='headings', columns=columns)

        for col in columns:
            width = 100
            if col == 'name':
                width = 150
            elif col == 'operation':
                width = 120
            self.tree.column(col, width=width, anchor=tk.CENTER)
            if col == 'id':
                col_text = 'ID'
            elif col == 'name':
                col_text = '姓名'
            elif col == 'chinese':
                col_text = '语文'
            elif col == 'math':
                col_text = '数学'
            elif col == 'english':
                col_text = '英语'
            elif col == 'operation':
                col_text = '操作'
            else:
                col_text = col
            self.tree.heading(col, text=col_text)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(parent_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Button-1>", lambda event: self._handle_query_tree_click(event, self.tree))

    def _load_data(self, exam_name, sort_option):
        """加载查询数据"""
        # 查询该考试下的所有自定义学科
        self.cursor.execute("""
            SELECT DISTINCT field_name 
            FROM student_fields 
            JOIN students ON student_fields.student_id = students.id 
            WHERE students.exam_name = ?
        """, (exam_name,))
        custom_fields = [row[0] for row in self.cursor.fetchall()]

        # 构建排序SQL
        sort_sql = ""
        if sort_option == "总成绩从高到低":
            sort_sql = "ORDER BY (CASE WHEN chinese = '无' THEN 0 ELSE chinese END + CASE WHEN math = '无' THEN 0 ELSE math END + CASE WHEN english = '无' THEN 0 ELSE english END) DESC"
        elif sort_option == "总成绩从低到高":
            sort_sql = "ORDER BY (CASE WHEN chinese = '无' THEN 0 ELSE chinese END + CASE WHEN math = '无' THEN 0 ELSE math END + CASE WHEN english = '无' THEN 0 ELSE english END) ASC"
        elif sort_option == "语文成绩从高到低":
            sort_sql = "ORDER BY CASE WHEN chinese = '无' THEN 0 ELSE chinese END DESC"
        elif sort_option == "语文成绩从低到高":
            sort_sql = "ORDER BY CASE WHEN chinese = '无' THEN 0 ELSE chinese END ASC"
        elif sort_option == "数学成绩从高到低":
            sort_sql = "ORDER BY CASE WHEN math = '无' THEN 0 ELSE math END DESC"
        elif sort_option == "数学成绩从低到高":
            sort_sql = "ORDER BY CASE WHEN math = '无' THEN 0 ELSE math END ASC"
        elif sort_option == "英语成绩从高到低":
            sort_sql = "ORDER BY CASE WHEN english = '无' THEN 0 ELSE english END DESC"
        elif sort_option == "英语成绩从低到高":
            sort_sql = "ORDER BY CASE WHEN english = '无' THEN 0 ELSE english END ASC"

        # 查询学生信息
        self.cursor.execute(f"""
            SELECT students.id, students.name, students.chinese, students.math, students.english
            FROM students
            WHERE students.exam_name = ?
            {sort_sql}
        """, (exam_name,))
        students = self.cursor.fetchall()

        # 清空表格
        for item in self.tree.get_children():
            self.tree.delete(item)

        for student in students:
            student_id = student[0]
            # 查询该学生的自定义学科成绩
            self.cursor.execute("""
                SELECT field_name, field_value 
                FROM student_fields 
                WHERE student_id = ?
            """, (student_id,))
            custom_scores = {row[0]: row[1] for row in self.cursor.fetchall()}

            values = list(student)
            for field in custom_fields:
                values.append(custom_scores.get(field, "无"))
            values.append("修改     |     删除")
            self.tree.insert("", tk.END, values=values, iid=student_id)

    def _handle_query_tree_click(self, event, tree):
        """处理查询表格点击事件"""
        region = tree.identify_region(event.x, event.y)
        if region == "cell":
            row_id = tree.identify_row(event.y)
            column = tree.identify_column(event.x)

            if column == f"#{len(tree['columns'])}":  # 操作列
                x, y, width, height = tree.bbox(row_id, column)
                click_offset = event.x - x
                if click_offset < width / 2:
                    # 修改学生信息
                    self._modify_student(row_id)
                else:
                    # 删除学生信息
                    self._delete_student(row_id)

    def _modify_student(self, student_id):
        """修改学生信息"""
        self.cursor.execute("SELECT name, chinese, math, english, exam_name FROM students WHERE id=?", (student_id,))
        student = self.cursor.fetchone()
        name, chinese, math, english, exam_name = student

        # 查询该学生的自定义学科成绩
        self.cursor.execute("""
            SELECT field_name, field_value 
            FROM student_fields 
            WHERE student_id = ?
        """, (student_id,))
        custom_scores = {row[0]: row[1] for row in self.cursor.fetchall()}

        modify_window = tk.Toplevel(self.root)
        modify_window.title("修改学生信息")
        self.center_window(modify_window, 1000, 900)
        modify_window.resizable(False, False)
        modify_window.grab_set()
        modify_window.configure(bg=BG_COLOR)

        frame = tk.Frame(modify_window, bg=BG_COLOR)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(frame, text="考试名称:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=5)
        exam_name_var = tk.StringVar()
        exam_name_combobox = ttk.Combobox(frame, textvariable=exam_name_var, values=self.exam_names, width=50)
        exam_name_combobox.pack(pady=5)
        exam_name_combobox.set(exam_name)

        # 基本信息部分 - 使用卡片效果
        basic_frame = tk.LabelFrame(frame, text="基本信息", font=HEADER_FONT, bg=BG_COLOR, bd=2, relief=tk.SOLID)
        basic_frame.pack(fill=tk.X, padx=10, pady=15)

        # 姓名行
        name_frame = tk.Frame(basic_frame, bg=BG_COLOR)
        name_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(name_frame, text="姓名:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        name_entry = tk.Entry(name_frame, font=FONT, bd=1, relief=tk.SOLID)
        name_entry.insert(0, name)
        name_entry.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # 成绩行1
        score_frame1 = tk.Frame(basic_frame, bg=BG_COLOR)
        score_frame1.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(score_frame1, text="语文:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        chinese_entry = tk.Entry(score_frame1, font=FONT, bd=1, relief=tk.SOLID, width=15)
        chinese_entry.insert(0, chinese)
        chinese_entry.pack(side=tk.LEFT, padx=10)

        tk.Label(score_frame1, text="数学:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        math_entry = tk.Entry(score_frame1, font=FONT, bd=1, relief=tk.SOLID, width=15)
        math_entry.insert(0, math)
        math_entry.pack(side=tk.LEFT, padx=10)

        tk.Label(score_frame1, text="英语:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        english_entry = tk.Entry(score_frame1, font=FONT, bd=1, relief=tk.SOLID, width=15)
        english_entry.insert(0, english)
        english_entry.pack(side=tk.LEFT, padx=10)

        # 动态字段部分
        dynamic_frame = tk.LabelFrame(frame, text="自定义学科", font=HEADER_FONT, bg=BG_COLOR, bd=2, relief=tk.SOLID)
        dynamic_frame.pack(fill=tk.X, padx=10, pady=15)

        # 动态字段显示区域
        dynamic_fields_frame = tk.Frame(dynamic_frame, bg=BG_COLOR)
        dynamic_fields_frame.pack(fill=tk.X, padx=10, pady=10)

        dynamic_field_entries = []
        for field_name, field_value in custom_scores.items():
            field_frame = tk.Frame(dynamic_fields_frame, bg=BG_COLOR)
            field_frame.pack(fill=tk.X, pady=5)

            tk.Label(field_frame, text=f"{field_name}:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
            field_entry = tk.Entry(field_frame, font=FONT, bd=1, relief=tk.SOLID, width=10)
            field_entry.insert(0, field_value)
            field_entry.pack(side=tk.LEFT, padx=10)

            remove_button = tk.Button(
                field_frame,
                text="删除",
                command=lambda name=field_name, frame=field_frame: self.remove_dynamic_field(name, frame),
                bg='#E74C3C',
                fg=BTN_FG_COLOR,
                font=('Microsoft YaHei', 9),
                relief=tk.FLAT,
                padx=5
            )
            remove_button.pack(side=tk.RIGHT, padx=10)

            dynamic_field_entries.append((field_name, field_frame, field_entry))

        # 动态字段添加区域
        add_field_frame = tk.Frame(dynamic_frame, bg=BG_COLOR)
        add_field_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(add_field_frame, text="学科名称:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        new_field_name = tk.Entry(add_field_frame, font=FONT, bd=1, relief=tk.SOLID, width=20)
        new_field_name.pack(side=tk.LEFT, padx=10)

        tk.Label(add_field_frame, text="学科分数:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        new_field_value = tk.Entry(add_field_frame, font=FONT, bd=1, relief=tk.SOLID, width=10)
        new_field_value.pack(side=tk.LEFT, padx=10)

        add_button = tk.Button(
            add_field_frame,
            text="添加",
            command=lambda: self.add_dynamic_field_modify(new_field_name, new_field_value, dynamic_fields_frame,
                                                          dynamic_field_entries),
            bg=BTN_BG_COLOR,
            fg=BTN_FG_COLOR,
            font=FONT,
            relief=tk.FLAT
        )
        add_button.pack(side=tk.LEFT, padx=10)

        def update_student():
            new_exam_name = exam_name_var.get()
            new_name = name_entry.get().strip()
            new_chinese = chinese_entry.get().strip()
            new_math = math_entry.get().strip()
            new_english = english_entry.get().strip()

            # 验证输入
            if not new_name:
                messagebox.showerror("错误", "姓名不能为空，请输入学生姓名。")
                return

            # 将空白项设置为“无”
            if not new_chinese:
                new_chinese = "无"
            if not new_math:
                new_math = "无"
            if not new_english:
                new_english = "无"

            try:
                if new_chinese != "无":
                    new_chinese = float(new_chinese)
                if new_math != "无":
                    new_math = float(new_math)
                if new_english != "无":
                    new_english = float(new_english)
            except ValueError:
                messagebox.showerror("错误", "成绩必须为数字，请输入有效的成绩。")
                return

            try:
                self.conn.execute("BEGIN")

                # 更新学生信息
                self.cursor.execute("""
                    UPDATE students 
                    SET name=?, chinese=?, math=?, english=?, exam_name=? 
                    WHERE id=?
                """, (new_name, new_chinese, new_math, new_english, new_exam_name, student_id))

                # 删除原有的自定义学科成绩
                self.cursor.execute("DELETE FROM student_fields WHERE student_id=?", (student_id,))

                # 添加新的自定义学科成绩
                for field_name, _, field_entry in dynamic_field_entries:
                    field_value = field_entry.get().strip()
                    if field_name.strip() and field_value.strip():
                        self.cursor.execute(
                            "INSERT INTO student_fields (student_id, field_name, field_value) VALUES (?, ?, ?)",
                            (student_id, field_name.strip(), field_value.strip()))

                self.conn.commit()
                messagebox.showinfo("成功", f"学生 {new_name} 的信息已更新。")
                modify_window.destroy()
                self._show_query_page()

            except sqlite3.Error as e:
                self.conn.execute("ROLLBACK")
                messagebox.showerror("错误", f"操作失败: {str(e)}，请稍后再试。")

        btn_frame = tk.Frame(frame, bg=BG_COLOR)
        btn_frame.pack(pady=30)

        submit_btn = tk.Button(
            btn_frame,
            text="提交",
            command=update_student,
            bg=BTN_BG_COLOR,
            fg=BTN_FG_COLOR,
            font=HEADER_FONT,
            padx=30,
            pady=8,
            relief=tk.FLAT
        )
        submit_btn.pack()

    def add_dynamic_field_modify(self, new_field_name, new_field_value, dynamic_fields_frame, dynamic_field_entries):
        """添加动态字段（修改页面）"""
        field_name = new_field_name.get().strip()
        field_value = new_field_value.get().strip()

        if not field_name:
            messagebox.showerror("错误", "字段名称不能为空，请输入有效的学科名称。")
            return

        # 检查是否已存在同名的动态字段
        for name, _, _ in dynamic_field_entries:
            if name == field_name:
                messagebox.showerror("错误", f"字段 '{field_name}' 已存在，请选择其他学科名称。")
                return

        # 验证字段值
        if field_value:
            try:
                float(field_value)
            except ValueError:
                messagebox.showerror("错误", "学科分数必须为数字，请输入有效的分数。")
                return

        # 创建字段显示行
        field_frame = tk.Frame(dynamic_fields_frame, bg=BG_COLOR)
        field_frame.pack(fill=tk.X, pady=5)

        tk.Label(field_frame, text=f"{field_name}:", font=FONT, bg=BG_COLOR).pack(side=tk.LEFT, padx=10)
        field_entry = tk.Entry(field_frame, font=FONT, bd=1, relief=tk.SOLID, width=10)
        field_entry.insert(0, field_value)
        field_entry.pack(side=tk.LEFT, padx=10)

        remove_button = tk.Button(
            field_frame,
            text="删除",
            command=lambda name=field_name, frame=field_frame: self.remove_dynamic_field(name, frame),
            bg='#E74C3C',
            fg=BTN_FG_COLOR,
            font=('Microsoft YaHei', 9),
            relief=tk.FLAT,
            padx=5
        )
        remove_button.pack(side=tk.RIGHT, padx=10)

        # 保存框架引用，用于后续删除
        dynamic_field_entries.append((field_name, field_frame, field_entry))

        # 清空输入框
        new_field_name.delete(0, tk.END)
        new_field_value.delete(0, tk.END)
        new_field_name.focus()

    def _delete_student(self, student_id):
        """删除学生信息"""
        self.cursor.execute("SELECT name FROM students WHERE id=?", (student_id,))
        name = self.cursor.fetchone()[0]

        if messagebox.askyesno("确认", f"确定要删除学生 {name} 的信息吗？此操作不可恢复。"):
            try:
                self.conn.execute("BEGIN")

                # 删除学生的自定义学科成绩
                self.cursor.execute("DELETE FROM student_fields WHERE student_id=?", (student_id,))

                # 删除学生信息
                self.cursor.execute("DELETE FROM students WHERE id=?", (student_id,))

                self.conn.commit()
                messagebox.showinfo("成功", f"学生 {name} 的信息已删除。")
                self._show_query_page()

            except sqlite3.Error as e:
                self.conn.execute("ROLLBACK")
                messagebox.showerror("错误", f"操作失败: {str(e)}，请稍后再试。")

    def _show_exam_management_page(self):
        """显示考试管理页面"""
        self._clear_content()
        frame = tk.Frame(self.content_frame, bg=BG_COLOR, padx=30, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # 页面标题
        tk.Label(frame, text="考试管理", font=('Microsoft YaHei', 18, 'bold'), bg=BG_COLOR).pack(pady=20)

        # 考试列表
        exam_listbox = tk.Listbox(frame, font=FONT, bg=BG_COLOR)
        exam_listbox.pack(fill=tk.BOTH, expand=True, pady=10)

        # 加载考试列表
        self.cursor.execute("SELECT exam_name FROM exams")
        exams = self.cursor.fetchall()
        for exam in exams:
            exam_listbox.insert(tk.END, exam[0])

        # 操作按钮
        btn_frame = tk.Frame(frame, bg=BG_COLOR)
        btn_frame.pack(pady=20)

        tk.Button(btn_frame, text="添加考试", command=self._show_add_exam_dialog,
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="删除考试", command=lambda: self._delete_exam(exam_listbox),
                  bg='#E74C3C', fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(side=tk.LEFT, padx=10)

    def _show_add_exam_dialog(self):
        """显示添加考试对话框"""
        add_exam_window = tk.Toplevel(self.root)
        add_exam_window.title("添加考试")
        self.center_window(add_exam_window, 400, 200)
        add_exam_window.resizable(False, False)
        add_exam_window.configure(bg=BG_COLOR)

        frame = tk.Frame(add_exam_window, bg=BG_COLOR)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(frame, text="考试名称:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=10)
        exam_name_entry = tk.Entry(frame, font=FONT, bd=1, relief=tk.SOLID)
        exam_name_entry.pack(fill=tk.X, pady=5)

        def add_exam():
            exam_name = exam_name_entry.get().strip()
            if not exam_name:
                messagebox.showerror("错误", "考试名称不能为空，请输入有效的考试名称。")
                return
            try:
                self.cursor.execute("INSERT INTO exams (exam_name) VALUES (?)", (exam_name,))
                self.conn.commit()
                self._load_exam_names()
                messagebox.showinfo("成功", f"考试 {exam_name} 创建成功。")
                add_exam_window.destroy()
                self._show_exam_management_page()
            except sqlite3.IntegrityError:
                messagebox.showerror("错误", f"考试 {exam_name} 已存在，请选择其他考试名称。")

        btn_frame = tk.Frame(frame, bg=BG_COLOR)
        btn_frame.pack(pady=20, fill=tk.X)

        tk.Button(btn_frame, text="确定", command=add_exam,
                  bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        tk.Button(btn_frame, text="取消", command=add_exam_window.destroy,
                  bg='#E74C3C', fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT).pack(
            side=tk.RIGHT, padx=10, fill=tk.X, expand=True)

    def _delete_exam(self, exam_listbox):
        """删除考试"""
        selected_index = exam_listbox.curselection()
        if not selected_index:
            messagebox.showerror("错误", "请选择要删除的考试。")
            return

        exam_name = exam_listbox.get(selected_index)
        if messagebox.askyesno("确认", f"确定要删除考试 {exam_name} 吗？此操作不可恢复。"):
            try:
                self.cursor.execute("DELETE FROM exams WHERE exam_name = ?", (exam_name,))
                self.conn.commit()
                self._load_exam_names()
                messagebox.showinfo("成功", f"考试 {exam_name} 已删除。")
                self._show_exam_management_page()
            except Exception as e:
                messagebox.showerror("错误", f"删除失败: {str(e)}，请稍后再试。")

    def _show_statistics_page(self):
        """显示数据统计页面"""
        self._clear_content()
        frame = tk.Frame(self.content_frame, bg=BG_COLOR, padx=30, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # 页面标题
        tk.Label(frame, text="数据统计与分析", font=('Microsoft YaHei', 18, 'bold'), bg=BG_COLOR).pack(pady=20)

        # 考试名称选择
        tk.Label(frame, text="选择考试:", font=HEADER_FONT, bg=BG_COLOR).pack(anchor=tk.W, pady=5)
        exam_name_var = tk.StringVar()
        exam_name_combobox = ttk.Combobox(frame, textvariable=exam_name_var, values=self.exam_names, width=30)
        exam_name_combobox.pack(pady=5)
        exam_name_combobox.set(self.exam_names[0] if self.exam_names else "")

        # 查询按钮
        query_btn = tk.Button(frame, text="查询统计数据",
                              command=lambda: self._show_statistics(exam_name_var.get(), frame),
                              bg=BTN_BG_COLOR, fg=BTN_FG_COLOR, font=FONT, relief=tk.FLAT)
        query_btn.pack(pady=20)

    def _show_statistics(self, exam_name, frame):
        """显示统计数据和图表"""
        # 查询该考试下的所有自定义学科
        self.cursor.execute("""
            SELECT DISTINCT field_name 
            FROM student_fields 
            JOIN students ON student_fields.student_id = students.id 
            WHERE students.exam_name = ?
        """, (exam_name,))
        custom_fields = [row[0] for row in self.cursor.fetchall()]

        # 查询学生信息
        self.cursor.execute("""
            SELECT students.id, students.name, students.chinese, students.math, students.english
            FROM students
            WHERE students.exam_name = ?
        """, (exam_name,))
        students = self.cursor.fetchall()

        scores = {
            'chinese': [],
            'math': [],
            'english': []
        }
        for field in custom_fields:
            scores[field] = []

        for student in students:
            student_id = student[0]
            chinese = student[2] if student[2] != "无" else 0
            math = student[3] if student[3] != "无" else 0
            english = student[4] if student[4] != "无" else 0

            scores['chinese'].append(float(chinese))
            scores['math'].append(float(math))
            scores['english'].append(float(english))

            # 查询该学生的自定义学科成绩
            self.cursor.execute("""
                SELECT field_name, field_value 
                FROM student_fields 
                WHERE student_id = ?
            """, (student_id,))
            custom_scores = {row[0]: row[1] for row in self.cursor.fetchall()}
            for field in custom_fields:
                score = custom_scores.get(field, 0)
                scores[field].append(float(score) if score != "无" else 0)

        # 处理没有学生数据的情况
        if not students:
            tk.Label(frame, text="该考试暂无学生数据", font=HEADER_FONT, bg=BG_COLOR).pack(pady=20)
            return

        # 统计信息
        stats_text = f"考试名称: {exam_name}\n"
        for subject, subject_scores in scores.items():
            if subject_scores:
                average = sum(subject_scores) / len(subject_scores)
                max_score = max(subject_scores)
                min_score = min(subject_scores)
                stats_text += f"{subject} 平均分: {average:.2f}, 最高分: {max_score}, 最低分: {min_score}\n"

        stats_label = tk.Label(frame, text=stats_text, font=FONT, bg=BG_COLOR, justify=tk.LEFT)
        stats_label.pack(pady=20, anchor=tk.W)

        # 绘制柱状图
        fig = Figure(figsize=(8, 5), dpi=100)
        ax = fig.add_subplot(111)
        subjects = list(scores.keys())
        averages = [sum(scores[subject]) / len(scores[subject]) if scores[subject] else 0 for subject in subjects]
        ax.bar(subjects, averages, color=['#3498db', '#2ecc71', '#e74c3c', '#f39c12', '#9b59b6'])
        ax.set_xlabel('学科', fontsize=12)
        ax.set_ylabel('平均分', fontsize=12)
        ax.set_title(f'{exam_name} 各学科平均分', fontsize=14)
        ax.tick_params(axis='x', rotation=45)  # 学科名称旋转45度避免重叠
        fig.tight_layout()  # 自动调整布局

        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=20, fill=tk.BOTH, expand=True)

    def _export_data(self):
        """导出数据到 Excel"""
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "学生信息"

            # 写入表头
            headers = ['姓名', '语文', '数学', '英语']
            self.cursor.execute("SELECT DISTINCT field_name FROM student_fields")
            custom_fields = [row[0] for row in self.cursor.fetchall()]
            headers.extend(custom_fields)
            ws.append(headers)

            # 写入数据
            self.cursor.execute("SELECT id, name, chinese, math, english FROM students")
            students = self.cursor.fetchall()
            for student_id, name, chinese, math, english in students:
                self.cursor.execute("""
                    SELECT field_name, field_value FROM student_fields 
                    WHERE student_id = ?
                """, (student_id,))
                custom_data = {row[0]: row[1] for row in self.cursor.fetchall()}

                row = [name, chinese, math, english]
                for field in custom_fields:
                    row.append(custom_data.get(field, "无"))

                ws.append(row)

            wb.save(file_path)
            messagebox.showinfo("成功", f"数据已导出到 {file_path}。")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}，请稍后再试。")

    def _clear_content(self):
        """清空内容区域"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def clear_window(self):
        """清空窗口内容"""
        for widget in self.root.winfo_children():
            widget.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = StudentSystem(root)
    root.mainloop()