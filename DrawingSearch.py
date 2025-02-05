import sys
import os
import io
import datetime
import base64
import glob
import threading
from tkinter import filedialog
from tkinter import ttk
from PIL import Image, ImageTk
from icon_base64 import ICON_BASE64

try:
    import tkinter as tk
    from tkinter import messagebox
except ModuleNotFoundError:
    print("Error: tkinter module is not available in this environment.")
    sys.exit(1)

# 全局变量
ver = "1.1.3"  # 版本号
search_history = []  # 用于存储最近的搜索记录，最多保存20条
changed_parts_path = None  # 用户更改的 PARTS 目录
result_frame = None  # 搜索结果的 Frame 容器
results_tree = None  # 搜索结果的 Treeview 控件
history_listbox = None  # 用于显示搜索历史的列表框
feeling_lucky_pressed = False  # 标志位，用于 "I'm Feeling Lucky!" 按钮
window_expanded = False  # 设置标志位，表示窗口是否已经扩展
window_width = 340
expand_window_width = 600
window_height = 315
shortcut_frame = None  # 用于快捷访问按钮的框架
default_parts_path = os.path.normpath("K:\\PARTS")
vault_cache = os.path.normpath("C:\\_Vault Working Folder\\Designs\\PARTS")  # Vault 缓存目录
# 快捷访问路径列表，存储显示文字和对应路径
shortcut_paths = [
    {"label": "Parts Folder", "path": "K:\\PARTS"},
    {"label": "Latest Missing List", "path": "V:\\Missing Lists\\Missing_Parts_List"},
    {"label": "Pneumatic Drawing Folder", "path": "K:\\Pneumatic Drawings"},
    {"label": "Projects Video Folder", "path": "G:\\Project Media"},
    {"label": "ERP Items Tool List", "path": "K:\\ERP TOOLS\\BELLATRX ERP ITEMS TOOL.xlsx"},
    {"label": "Equipment Labels Details", "path": "K:\\Equipment Labels\\Equipment Label - New"},
]

class Tooltip:
    def __init__(self, widget, get_text_callback, delay=500):
        self.widget = widget
        self.get_text_callback = get_text_callback  # 动态获取文字的回调
        self.delay = delay  # 延迟时间（毫秒）
        self.tooltip_window = None
        self.after_id = None
        # self.last_motion_time = 0

        self.widget.bind("<Enter>", self.schedule_show)
        self.widget.bind("<Leave>", self.hide_tooltip)
        self.widget.bind("<Motion>", self.update_position)

    def schedule_show(self, event):
        """安排显示提示"""
        self.after_id = self.widget.after(self.delay, self.show_tooltip)

    def show_tooltip(self, event=None):
        """在鼠标位置显示提示"""
        if self.tooltip_window or not self.get_text_callback:
            return

        text = self.get_text_callback()  # 动态获取当前文字
        if not text:
            return

        x, y, _, _ = self.widget.bbox("insert")  # 获取Label的位置
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + 20

        # 创建一个新的 Tooltip 窗口
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # 去掉边框
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)  # 确保窗口在最上层

        label = tk.Label(tw, text=text, justify="left", background="#ffffe0", relief="solid", borderwidth=1, font=("Arial", 9))
        label.pack(ipadx=5, ipady=3)

    def hide_tooltip(self, event=None):
        """隐藏提示"""
        if self.after_id:
            self.widget.after_cancel(self.after_id)
            self.after_id = None
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

    def update_position(self, event):
        """更新工具提示的位置"""
        ''' # 防抖处理 
        current_time = time.time()
        if current_time - self.last_motion_time > 0.1:  # 100毫秒防抖
            self.last_motion_time = current_time
        '''
        if self.tooltip_window:
            x = event.x_root + 20
            y = event.y_root + 20
            self.tooltip_window.wm_geometry(f"+{x}+{y}")

def show_warning_message(message):
    """在输入框下方显示警告信息"""
    global warning_label
    if "Tip" in message:
        warning_label.config(fg="blue")  # 提示类的信息用蓝色显示
    else:
        warning_label.config(fg="red")
    warning_label.config(text=message)

def hide_warning_message():
    """隐藏警告信息"""
    global warning_label
    if warning_label:
        warning_label.config(text="")

def open_shortcut(index):
    """打开快捷访问的路径或文件"""
    path = shortcut_paths[index]["path"]

    if os.path.exists(path):
        # 如果 label 包含 "Missing Lists" 字段，使用get_latest_file函数查找最新的文件
        if "Missing List" in shortcut_paths[index]["label"]:
            prefix_name = "Master_Missing_List"
            latest_file = get_latest_file(prefix_name, path)
            if latest_file:
                path = latest_file
            else:
                messagebox.showwarning("Warning", "No Missing List file found!")
                return

        # 如果 label 包含 "Equipment Labels Details" 字段，使用get_latest_file函数查找最新的文件
        if "Equipment Labels Details" in shortcut_paths[index]["label"]:
            prefix = "Equipment New Labels Details"
            latest_file = get_latest_file(prefix, path)
            if latest_file:
                path = latest_file
            else:
                messagebox.showwarning("Warning", "No Equipment Labels Details file found!")
                return
        
        try:
            # 如果是目录，根据操作系统选择不同的打开方式
            if os.path.isdir(path):
                os.startfile(path)  # 打开目录
            else:
                # 如果是文件，直接通过 open_file 打开
                open_file(file_path=path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open shortcut: {e}")
    else:
        messagebox.showerror("Error", f"Shortcut path does not exist: {path}")

def update_directory():
    """更新搜索目录"""
    global default_parts_path, changed_parts_path
    default_directory = default_parts_path
    new_dir = filedialog.askdirectory(initialdir=default_directory, title="Select Directory")
    if new_dir:
        new_dir = new_dir.replace('/', '\\')  # 将路径中的斜杠替换为反斜杠
        default_directory = new_dir
        directory_label.config(text=f"PARTS Directory: {default_directory}")
        changed_parts_path = new_dir

def reset_to_default_directory():
    """将搜索路径重置为默认路径"""
    global default_parts_path, changed_parts_path
    default_directory = default_parts_path  # 重置为默认路径
    directory_label.config(text=f"Default PARTS Directory: {default_directory}")
    changed_parts_path = None

def open_file(event=None, file_path=None):
    """用系统默认程序打开选中的文件"""
    if not file_path:  # 如果没有传入路径，则尝试从 Treeview 中获取
        selected_item = results_tree.selection()
        if selected_item:
            file_path = results_tree.item(selected_item, 'values')[2]  # 获取完整文件路径
        else:
            # 用户如果点击表头不做任何操作，直接返回
            return
    if file_path and os.path.exists(file_path):
        try:
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open file: {e}")
    else:
        messagebox.showerror("Error", "File not found!")

def save_search_history(query):
    """保存搜索记录并限制最多保存20条"""
    global search_history
    if query and query not in search_history:
        search_history.append(query)
        if len(search_history) > 20:
            search_history.pop(0)

def show_search_history(event):
    """在输入框下显示搜索历史"""
    global history_listbox

    # 如果存在旧的列表框，先销毁它
    if history_listbox:
        history_listbox.destroy()

    query = entry.get().lower()
    if not query:
        matching_history = search_history
    else:
        matching_history = [h for h in search_history if query in h.lower()]

    if not matching_history:
        return  # 如果没有匹配的历史记录，则不显示列表框
    
    if len(matching_history) == 1:
        if matching_history[0].lower() == query.lower():
            return

    # 创建列表框
    history_listbox = tk.Listbox(root, height=min(len(matching_history), 5))
    for item in matching_history:
        history_listbox.insert(tk.END, item)

    # 获取输入框的绝对位置
    # 因为entry放置在entry_frame中，所以需要计算相对位置，用entry获取x坐标，用entry_frame获取y坐标
    x = entry.winfo_x()
    y = entry_frame.winfo_y() + entry.winfo_height()

    # 放置列表框
    history_listbox.place(x=x, y=y, width=entry.winfo_width())
    history_listbox.bind("<ButtonRelease-1>", lambda event: select_history(event, history_listbox))

def hide_history(event):
    """点击窗口其他部分时隐藏搜索历史"""
    global history_listbox
    if history_listbox:
        widget = event.widget
        if widget != entry and widget != history_listbox:
            history_listbox.destroy()
            history_listbox = None

def select_history(event, listbox):
    """当选择历史记录时，填充到输入框并销毁列表框"""
    if not listbox.curselection():
        return  # 如果没有选中任何项，直接返回

    selection = listbox.get(listbox.curselection())
    entry.delete(0, tk.END)
    entry.insert(0, selection)
    entry.focus_set()  # 重新聚焦到输入框

    # 销毁列表框
    global history_listbox
    history_listbox.destroy()
    history_listbox = None

def search_pdf_files():
    """搜索目录下的 PDF 文件"""
    if not feeling_lucky_pressed:
        search_btn.config(state=tk.DISABLED)  # 禁用按钮
    else:
        lucky_btn.config(state=tk.DISABLED)  # 禁用按钮
    hide_warning_message()  # 清除警告信息
    query = entry.get().strip() # 去除首尾空格
    if not query:
        show_warning_message("Please enter part or assembly number!")
        if not feeling_lucky_pressed:
            search_btn.config(state=tk.NORMAL)  # 启用按钮
        else:
            lucky_btn.config(state=tk.NORMAL)  # 启用按钮
        return

    save_search_history(query)  # 保存搜索记录

    # 提取前两位字符并更新搜索路径
    prefix = query[:2]
    if changed_parts_path:
        search_directory = os.path.join(changed_parts_path, prefix)
    else:
        search_directory = os.path.join(default_parts_path, prefix)

    if not os.path.exists(search_directory):
        show_warning_message(f"Path does not exist! {search_directory}")
        show_result_list(None) # 目录不存在就清空已有搜索结果
        if not feeling_lucky_pressed:
            search_btn.config(state=tk.NORMAL)  # 启用按钮
        else:
            lucky_btn.config(state=tk.NORMAL)  # 启用按钮
        return

    # 执行搜索
    show_warning_message(f"Searching for \"{query}\", please wait...")
    search_thread = threading.Thread(target=search_pdf_files_thread, args=(query, search_directory,))
    search_thread.start()

def search_pdf_files_thread(query, search_directory):
    """使用多线程搜索目录下的 PDF 文件"""
    try:
        is_feeling_lucky = feeling_lucky_pressed
        result_files = []
        for root_dir, _, files in os.walk(search_directory):
            for file in files:
                if file.endswith(".pdf") and query.lower() in file.lower():
                    file_path = os.path.join(root_dir, file)
                    create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                    result_files.append((file, create_time, file_path))

        # 排序结果（按创建时间倒序）
        result_files.sort(key=lambda x: x[1], reverse=True)

        root.after(0, hide_warning_message)  # 使用主线程清除警告信息

        # "I'm Feeling Lucky" 功能：直接打开第一个文件，一般按创建时间排序后就是最新的revision
        if is_feeling_lucky:
            if result_files:
                root.after(0, lambda: open_file(file_path=result_files[0][2]))
                result_files = []
                root.after(0, lambda: show_result_list(result_files))
            else:
                root.after(0, lambda: show_warning_message("No matching drawing PDF found!"))
            root.after(0, lambda: lucky_btn.config(state=tk.NORMAL))  # 启用按钮
            return

        if not result_files:
            root.after(0, lambda: show_warning_message("No matching drawing PDF found!"))

        root.after(0, lambda: show_result_list(result_files))
        root.after(0, lambda: search_btn.config(state=tk.NORMAL))  # 启用按钮
    except Exception as e:
        root.after(0, lambda: messagebox.showerror("Error", f"An error occurred in search thread: {e}"))

def search_3d_files():
    """搜索目录下的 3D 文件(ipt或者iam)"""
    search_3d_btn.config(state=tk.DISABLED)  # 禁用按钮
    hide_warning_message()  # 清除警告信息
    query = entry.get().strip() # 去除首尾空格
    if not query:
        show_warning_message("Please enter part or assembly number!")
        search_3d_btn.config(state=tk.NORMAL)  # 启用按钮
        return

    save_search_history(query)  # 保存搜索记录

    # 提取前两位字符并更新搜索路径
    prefix = query[:2]
    if changed_parts_path:
        search_directory = os.path.join(changed_parts_path, prefix)
    else:
        search_directory = os.path.join(default_parts_path, prefix)

    if not os.path.exists(search_directory):
        show_warning_message(f"Path does not exist! {search_directory}")
        show_result_list(None) # 目录不存在就清空已有搜索结果
        search_3d_btn.config(state=tk.NORMAL)  # 启用按钮
        return

    # 执行搜索
    show_warning_message(f"Searching for \"{query}\", please wait...")
    search_thread = threading.Thread(target=search_3d_files_thread, args=(query, search_directory,))
    search_thread.start()

def search_3d_files_thread(query, search_directory):
    try:
        """使用多线程搜索目录下的 3D 文件"""
        result_files = []
        for root_dir, _, files in os.walk(search_directory):
            for file in files:
                if file.endswith(".iam") and query.lower() in file.lower():
                    file_path = os.path.join(root_dir, file)
                    create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                    result_files.append((file, create_time, file_path))

        for root_dir, _, files in os.walk(search_directory):
            for file in files:
                if file.endswith(".ipt") and query.lower() in file.lower():
                    file_path = os.path.join(root_dir, file)
                    create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                    result_files.append((file, create_time, file_path))

        root.after(0, hide_warning_message)  # 使用主线程清除警告信息
        if not result_files:
            root.after(0, lambda: show_warning_message("No matching 3D drawing found!"))
        # 排序结果（按文件名排序）
        result_files.sort(key=lambda x: x[0])
        root.after(0, lambda: show_result_list(result_files))
        root.after(0, lambda: search_3d_btn.config(state=tk.NORMAL))  # 启用按钮
    except Exception as e:
        root.after(0, lambda: messagebox.showerror("Error", f"An error occurred in search thread: {e}"))

def search_vault_cache():
    """搜索Vault缓存目录下的 3D 文件(ipt或者iam)"""
    search_cache_btn.config(state=tk.DISABLED)  # 禁用按钮
    hide_warning_message()  # 清除警告信息
    query = entry.get().strip() # 去除首尾空格
    if not query:
        show_warning_message("Please enter part or assembly number!")
        search_cache_btn.config(state=tk.NORMAL)  # 启用按钮
        return

    save_search_history(query)  # 保存搜索记录

    # 判断 query 字符串的第三位是否是数字，如果是数字，作为project number进行搜索
    if len(query) > 2 and query[2].isdigit():
        if '-' in query:
            proj_no = query.split('-')[0]
        else:
            proj_no = query

        sub_dir = "S" + proj_no
        prefix = "S"
        matching_directories = glob.glob(os.path.join(vault_cache, prefix, sub_dir + '*'))
        if matching_directories:
            search_directory = matching_directories[0]
        else:
            show_warning_message("No matching 3D drawings are cached. Check in Vault!")
            show_result_list(None) # 目录不存在就清空已有搜索结果
            search_cache_btn.config(state=tk.NORMAL)
            return
    else:
        # 作为part number或者assembly number，提取前两位字符并更新搜索路径
        prefix = query[:2]
        search_directory = os.path.join(vault_cache, prefix)
        if not os.path.exists(search_directory):
            show_warning_message("No matching 3D drawings are cached. Check in Vault!")
            show_result_list(None) # 目录不存在就清空已有搜索结果
            search_cache_btn.config(state=tk.NORMAL)  # 启用按钮
            return

    # 执行搜索
    show_warning_message(f"Searching for \"{query}\", please wait...")
    search_thread = threading.Thread(target=search_vault_cache_thread, args=(query, search_directory,))
    search_thread.start()

def search_vault_cache_thread(query, search_directory):
    try:
        """使用多线程搜索Vault缓存目录下的 3D 文件"""
        result_files = []
        for root_dir, _, files in os.walk(search_directory):
            for file in files:
                if file.endswith(".iam") and query.lower() in file.lower():
                    file_path = os.path.join(root_dir, file)
                    create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                    result_files.append((file, create_time, file_path))

        for root_dir, _, files in os.walk(search_directory):
            for file in files:
                if file.endswith(".ipt") and query.lower() in file.lower():
                    file_path = os.path.join(root_dir, file)
                    create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                    result_files.append((file, create_time, file_path))

        root.after(0, hide_warning_message)  # 使用主线程清除警告信息
        if not result_files:
            root.after(0, lambda: show_warning_message("No matching 3D drawing cached to the local disk!"))
        if len(query) > 2 and query[2].isdigit():
            # 如果是project number，按文件名正序排列
            result_files.sort(key=lambda x: x[0])
        else:
            # 如果是part number或者assembly number，按创建时间倒序排列
            result_files.sort(key=lambda x: x[1], reverse=True)

        root.after(0, lambda: show_result_list(result_files))
        root.after(0, lambda: search_cache_btn.config(state=tk.NORMAL))  # 启用按钮
        root.after(0, lambda: show_warning_message("Tip: Searched from cache, may not be the latest update!"))
    except Exception as e:
        root.after(0, lambda: messagebox.showerror("Error", f"An error occurred in search thread: {e}"))

def show_result_list(result_files):
    """显示搜索结果"""
    global result_frame, results_tree
    if not result_files:
        if result_frame:
            result_frame.destroy()
            if window_expanded:
                root.geometry(f"{expand_window_width}x{window_height}")
            else:
                root.geometry(f"{window_width}x{window_height}")
        return

    # 创建结果显示区域
    if result_frame:
        result_frame.destroy()
    result_frame = tk.Frame(root)
    result_frame.pack(fill=tk.BOTH, expand=True, pady=0)
    tip_label = tk.Label(result_frame, text="Double click to open the file", fg="blue")
    tip_label.pack(padx=5, pady=0, anchor="w")

    # 添加 Treeview 控件显示结果
    columns = ("File Name", "Created Time", "Path")
    results_tree = ttk.Treeview(result_frame, columns=columns, show="headings")
    results_tree.heading("File Name", text="File Name", anchor="w")
    results_tree.heading("Created Time", text="Created Time", anchor="w")
    results_tree.column("File Name", width=150, anchor="w")
    results_tree.column("Created Time", width=135, anchor="w")
    results_tree.column("Path", width=0, stretch=tk.NO)  # 隐藏第三列

    # 创建一个垂直滚动条并将其与 Treeview 关联
    scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=results_tree.yview)
    results_tree.configure(yscrollcommand=scrollbar.set)

    # 设置 Treeview 表头和行样式
    style = ttk.Style()
    style.configure("Treeview.Heading", background="#A9A9A9", foreground="black", font=("Arial", 10, "bold"))
    style.configure("Treeview", rowheight=25)
    style.map("Treeview", background=[('selected', '#347083')])

    # 插入搜索结果
    for index, item in enumerate(result_files):
        tag = 'evenrow' if index % 2 == 0 else 'oddrow'
        results_tree.insert("", tk.END, values=(item[0], item[1], item[2]), tags=(tag,))
    
    results_tree.tag_configure('evenrow', background='#E6F7FF')
    results_tree.tag_configure('oddrow', background='white')

    results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    results_tree.bind("<Double-1>", open_file)

    # 动态调整窗口大小以显示结果
    root.update_idletasks()
    new_height = 420 + len(result_files) * 20 if result_files else window_height
    if window_expanded:
        root.geometry(f"{expand_window_width}x{min(new_height, 600)}")
    else:
        root.geometry(f"{window_width}x{min(new_height, 600)}")

'''
def show_about():
    messagebox.showinfo("About", f"This is a mini-app for quickly accessing\rdrawings on BellatRx computers.\
                        \r\rIf you have any questions or suggestions,\rplease feel free to contact me.\
                        \r\rVersion: {ver}\rDeveloped by: Wei Tang\rContact: wtweitang@hotmail.com")
'''

def show_about():
    # 创建自定义关于窗口
    about_win = tk.Toplevel(root)
    about_win.attributes("-topmost", True)
    about_win.title("About")
    about_win.geometry("380x280")
    about_win.resizable(False, False)
    
    # 设置窗口图标（复用主窗口图标）
    about_win.tk.call("wm", "iconphoto", about_win._w, icon)
    
    # 主容器框架
    main_frame = tk.Frame(about_win)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 左侧图标区域
    icon_frame = tk.Frame(main_frame, width=100)
    icon_frame.pack(side=tk.LEFT, fill=tk.Y)
    
    try:
        # 解码Base64图标并调整大小
        icon_data = base64.b64decode(ICON_BASE64)
        img = Image.open(io.BytesIO(icon_data))
        img = img.resize((64, 64), Image.Resampling.LANCZOS)  # 调整图标尺寸
        tk_img = ImageTk.PhotoImage(img)
        icon_label = tk.Label(icon_frame, image=tk_img)
        icon_label.image = tk_img  # 保持引用
        icon_label.pack(pady=20)
    except Exception as e:
        print(f"Error loading icon: {e}")

    # 右侧文本区域
    text_frame = tk.Frame(main_frame)
    text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=20)

    # 文本内容
    about_text = [
        f"Drawing Search - Version {ver}",
        "\nThis is a mini-app for quickly accessing",
        "drawings on BellatRx computers.",
        "\nIf you have any questions or suggestions,",
        "please feel free to contact me.",
        "\nDeveloped by: Wei Tang",
    ]

    # 文本标签
    for i, text in enumerate(about_text):
        if i == 0:
            label = tk.Label(text_frame, text=text, font=("Arial", 10, "bold"), anchor="w")
        else:
            label = tk.Label(text_frame, text=text, anchor="w")
        label.pack(anchor="w", fill=tk.X)

    # 邮箱按钮和地址
    email_frame = tk.Frame(text_frame)
    email_frame.pack(anchor="w")

    email_btn = tk.Button(email_frame, text="Email me", bg="#4CAF50", command=lambda: send_email())
    email_btn.pack(side=tk.LEFT, padx=0)

    email_label = tk.Label(email_frame, text=": wtweitang@hotmail.com")
    email_label.pack(side=tk.LEFT)

    ok_button = tk.Button(about_win, text="OK", width=12, height=1, command=about_win.destroy)
    ok_button.pack(padx=20, pady=15, side=tk.RIGHT)

    # 窗口居中
    about_win.update_idletasks()
    window_width = about_win.winfo_width()
    window_height = about_win.winfo_height()
    position_right = int(about_win.winfo_screenwidth()/2 - window_width/2)
    position_down = int(about_win.winfo_screenheight()/3 - window_height/2)
    about_win.geometry(f"+{position_right}+{position_down}")

def send_email():
    """打开默认邮件客户端发送邮件"""
    import webbrowser
    try:
        webbrowser.open("mailto:wtweitang@hotmail.com?subject=Drawing%20Search%20Feedback")
    except Exception as e:
        messagebox.showerror("Error", f"Cannot open email client: {e}")

def reset_window():
    """恢复主窗口到初始状态"""
    global result_frame, results_tree, history_listbox, window_expanded, shortcut_frame
    entry.delete(0, tk.END)  # 清空输入框
    hide_warning_message()  # 清除警告信息
    if result_frame:
        result_frame.destroy()
        result_frame = None
        results_tree = None
    root.geometry(f"{window_width}x{window_height}")  # 恢复初始窗口大小
    if window_expanded:
        expand_btn.config(text="Quick Access   >>")  # 改为 ">>"
        window_expanded = not window_expanded  # 切换状态
        if shortcut_frame:
            shortcut_frame.destroy()
            shortcut_frame = None

def feeling_lucky():
    """设置标志位并执行搜索"""
    global feeling_lucky_pressed
    feeling_lucky_pressed = True
    search_pdf_files()
    feeling_lucky_pressed = False

def get_latest_file(prefix_name, directory):
    """查找并返回给定目录中最新修改的 Excel 文件的绝对路径"""
    latest_file = None
    latest_time = 0

    # 遍历目录中的所有文件
    for filename in os.listdir(directory):
        # 排除excel打开时产生的临时文件
        if filename.startswith("~"):
            continue
        if (filename.endswith(".xlsx") or filename.endswith(".xls")) and filename.startswith(prefix_name):
            file_path = os.path.join(directory, filename)
            file_mtime = os.path.getmtime(file_path)
            if file_mtime > latest_time:
                latest_time = file_mtime
                latest_file = file_path

    return latest_file

def toggle_window_size():
    """切换窗口大小和按钮文本，并动态显示快捷访问按钮"""
    global window_expanded, shortcut_frame, result_frame
    if window_expanded:
        # 如果已有搜索结果，保持显示搜索结果
        if result_frame:
            height = root.winfo_height()  # 获取当前窗口高度
            root.geometry(f"{window_width}x{height}")
        else:
            # 收缩窗口，隐藏快捷按钮框架
            root.geometry(f"{window_width}x{window_height}")  # 恢复到原始大小
        expand_btn.config(text="Quick Access   >>")  # 改为 ">>"
        if shortcut_frame:
            shortcut_frame.destroy()
            shortcut_frame = None
    else:
        # 如果已有搜索结果，保持显示搜索结果
        if result_frame:
            height = root.winfo_height()  # 获取当前窗口高度
            root.geometry(f"{expand_window_width}x{height}")
        else:
            # 扩展窗口，显示快捷按钮框架
            root.geometry(f"{expand_window_width}x{window_height}")  # 扩展窗口大小
        expand_btn.config(text="Quick Access   <<")  # 改为 "<<"

        # 创建快捷按钮框架
        if not shortcut_frame:
            shortcut_frame = tk.Frame(root)
            shortcut_frame.place(x=350, y=42, width=200, height=250)  # 定位到右侧扩展区域

        for i, shortcut in enumerate(shortcut_paths):
            btn = tk.Button(
                shortcut_frame, 
                text=shortcut["label"], 
                width=100, 
                command=lambda i=i: open_shortcut(i)
            )
            btn.pack(pady=6, padx=10, anchor="w")

    window_expanded = not window_expanded  # 切换状态

def center_window(root, width, height):
    """将窗口显示在屏幕中央偏上"""
    # 获取屏幕宽度和高度
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # 计算窗口左上角的坐标
    x = (screen_width - width) // 2
    y = (screen_height - height) // 5
    
    # 设置窗口的大小和位置
    root.geometry(f"{width}x{height}+{x}+{y}")

# 窗口置顶
def toggle_topmost():
    # 根据复选框的状态设置窗口是否置顶
    is_checked = topmost_var.get()
    root.attributes("-topmost", is_checked)

# 为 Entry 小部件创建一个右键菜单
def create_entry_context_menu(entry_widget):
    # 创建菜单
    context_menu = tk.Menu(root, tearoff=0)
    
    # 定义菜单项及其功能
    def cut_text():
        entry_widget.event_generate("<<Cut>>")

    def copy_text():
        entry_widget.event_generate("<<Copy>>")

    def paste_text():
        entry_widget.event_generate("<<Paste>>")

    # 添加菜单项
    context_menu.add_command(label="Copy", command=copy_text)
    context_menu.add_command(label="Cut", command=cut_text)
    context_menu.add_command(label="Paste", command=paste_text)

    # 绑定右键事件
    def show_context_menu(event):
        context_menu.tk_popup(event.x_root, event.y_root)

    # 将右键单击事件绑定到 Entry 小部件
    entry_widget.bind("<Button-3>", show_context_menu)

# 创建主窗口
try:
    root = tk.Tk()
    # 将 Base64 解码为二进制
    icon_data = base64.b64decode(ICON_BASE64)
    # 通过 BytesIO 读取 ICO 图标
    icon_image = Image.open(io.BytesIO(icon_data))
    icon = ImageTk.PhotoImage(icon_image)
    # 窗口居中偏上显示
    center_window(root, window_width, window_height)
    root.title("Drawing Search")
    # 设置窗口图标
    root.tk.call("wm", "iconphoto", root._w, icon)
    root.geometry(f"{window_width}x{window_height}")  # 初始窗口大小

    root.resizable(False, False)
    
    # 输入框和标签分开为两行
    label_frame = tk.Frame(root)
    label_frame.pack(pady=10, anchor="w", fill="x")

    # 标签放在第一行
    prompt_label = tk.Label(label_frame, text="Input Part / Assembly / Project Number :", anchor="w")
    prompt_label.pack(side=tk.LEFT, padx=(20, 0))

    # 添加置顶选项
    # 创建一个 IntVar 绑定复选框的状态（0 未选中，1 选中）
    topmost_var = tk.IntVar()

    # 创建复选框，用于控制窗口置顶
    checkbox = tk.Checkbutton(label_frame, text="📌", font=("Arial", 12), variable=topmost_var, command=toggle_topmost)
    checkbox.pack(anchor="e", padx=5)
    Tooltip(checkbox, lambda: "Pin to top", delay=500)

    # 创建输入框框架
    entry_frame = tk.Frame(root)
    entry_frame.pack(pady=0, anchor="w", fill="x")
    entry = tk.Entry(entry_frame, width=25, font=("Arial", 16))
    entry.pack(padx=20, anchor="w")
    create_entry_context_menu(entry)
    entry.focus()
    entry.bind("<Return>", lambda event: search_pdf_files())
    entry.bind("<Button-1>", show_search_history)  # 点击输入框时显示历史记录
    entry.bind("<KeyRelease>", show_search_history)  # 输入时实时更新匹配历史
    # 用于显示警告信息的标签
    warning_label = tk.Label(entry_frame, text="", fg="red", anchor="w")
    warning_label.pack(fill="x", padx=20)

    # 用户点击非 Listbox 或 Entry 区域时销毁 Listbox
    root.bind("<Button-1>", hide_history)

    # 添加按钮框架
    button_frame = tk.Frame(root)
    button_frame.pack(pady=5, padx=5, anchor="w")

    # Search PDF 按钮
    search_btn = tk.Button(button_frame, text="Search PDF Drawing", width=18, command=search_pdf_files)
    search_btn.grid(row=0, column=0, padx=15, pady=10)
    Tooltip(search_btn, lambda: "Search for PDF files matching the entered keywords", delay=500)

    # I'm Feeling Lucky 按钮
    lucky_btn = tk.Button(button_frame, text="I'm Feeling Lucky!", width=18, command=feeling_lucky)
    lucky_btn.grid(row=0, column=1, padx=15, pady=10)
    Tooltip(lucky_btn, lambda: "Open the latest revision of the PDF drawing", delay=500)

    # Search 3D drawing 按钮
    search_3d_btn = tk.Button(button_frame, text="Search 3D Drawing", width=18, command=search_3d_files)
    search_3d_btn.grid(row=1, column=0, padx=15, pady=10)
    Tooltip(search_3d_btn, lambda: "Search for 3D files (.iam/.ipt) matching the entered keywords", delay=500)

    # Search vault cache 按钮
    search_cache_btn = tk.Button(button_frame, text="Search in Vault Cache", width=18, command=search_vault_cache)
    search_cache_btn.grid(row=1, column=1, padx=15, pady=10)
    Tooltip(search_cache_btn, lambda: "Search 3D drawings (.iam/.ipt) from local Vault cache", delay=500)

    # Reset 按钮
    reset_btn = tk.Button(button_frame, text="Reset", width=18, command=reset_window)
    reset_btn.grid(row=2, column=0, padx=15, pady=8)
    Tooltip(reset_btn, lambda: "Reset the window to the initial state", delay=500)

    # 扩展按钮
    expand_btn = tk.Button(button_frame, text="Quick Access   >>", width=18, command=toggle_window_size)
    expand_btn.grid(row=2, column=1, padx=15, pady=8)
    Tooltip(expand_btn, lambda: "Shortcuts to frequently used folders and files", delay=500)

    # 显示默认目录及更改功能
    directory_frame = tk.Frame(root)
    directory_frame.pack(anchor="w", padx=20, pady=5, fill="x")
    directory_label = tk.Label(directory_frame, text=f"Default PARTS Directory: {default_parts_path}", font=("Arial", 8), width=34, anchor="w")
    directory_label.pack(side=tk.LEFT)
    Tooltip(directory_label, lambda: directory_label.cget("text"), delay=500)

    # Change 按钮
    change_label = tk.Label(directory_frame, text="Change", fg="blue", cursor="hand2", font=("Arial", 8, "underline"))
    change_label.pack(side=tk.LEFT, padx=3)
    Tooltip(change_label,  lambda: "Select a new PARTS directory", delay=500)
    change_label.bind("<Button-1>", lambda event: update_directory())

    # Default 按钮
    default_label = tk.Label(directory_frame, text="Default", fg="blue", cursor="hand2", font=("Arial", 8, "underline"))
    default_label.pack(side=tk.LEFT, padx=0)
    Tooltip(default_label,  lambda: "Reset the PARTS directory to default", delay=500)
    default_label.bind("<Button-1>", lambda event: reset_to_default_directory())

    # About 按钮
    about_frame = tk.Frame(root)
    about_frame.pack(anchor="e", padx=0, pady=5, fill="x")
    about_label = tk.Label(about_frame, text="ⓘ", fg="black", cursor="hand2", font=("Fixedsys", 12, "bold"))
    about_label.pack(anchor="e", padx=5, pady=5)
    Tooltip(about_label,  lambda: "About", delay=500)
    about_label.bind("<Button-1>", lambda event: show_about())

    # 运行主循环
    root.mainloop()
except Exception as e:
    print(f"An error occurred: {e}")
