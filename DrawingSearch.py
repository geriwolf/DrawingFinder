import sys
import os
import io
import re
import datetime
import collections
import base64
import glob
import threading
import ctypes
import fitz  # PyMuPDF
from tkinter import filedialog
from tkinter import ttk
from PIL import Image, ImageTk
from logo import ICON_BASE64

try:
    import tkinter as tk
    from tkinter import messagebox
except ModuleNotFoundError:
    print("Error: tkinter module is not available in this environment.")
    sys.exit(1)

# 全局变量
ver = "1.3.2"  # 版本号
search_history = []  # 用于存储最近的搜索记录，最多保存20条
changed_parts_path = None  # 用户更改的 PARTS 目录
result_frame = None  # 搜索结果的 Frame 容器
results_tree = None  # 搜索结果的 Treeview 控件
result_files_pdf = None  # 存储pdf搜索结果
result_files_3d = None  # 存储3d文件iam和ipt搜索结果
last_query = None  # 上一次的搜索结果
history_listbox = None  # 用于显示搜索历史的列表框
window_expanded = False  # 设置标志位，表示窗口是否已经扩展
window_width = 345
expand_window_width = 560
window_height = 300
stop_event = threading.Event()
active_threads = set()
shortcut_frame = None  # 用于快捷访问按钮的框架
default_parts_path = os.path.normpath("K:\\PARTS") # 默认 PARTS 目录
vault_cache = os.path.normpath("C:\\_Vault Working Folder\\Designs\\PARTS")  # Vault 缓存目录
# 全局缓存字典，键为目录路径，值为该目录下的所有文件信息列表
directory_cache = collections.OrderedDict()  # 使用 OrderedDict 维护缓存顺序
cache_max_size = 10  # 设置缓存最大条目数，防止缓存过大
cache_lock = threading.Lock()  # 用于保护缓存的线程锁
thumbnail_win = None
# 快捷访问路径列表，存储按钮上显示的文字和对应路径
shortcut_paths = [
    {"label": "Parts Folder", "path": "K:\\PARTS"},
    {"label": "Latest Missing List", "path": "V:\\Missing Lists\\Missing_Parts_List"},
    {"label": "Pneumatic Drawing Folder", "path": "K:\\Pneumatic Drawings"},
    {"label": "Projects Video Folder", "path": "G:\\Project Media"},
    {"label": "ERP Items Tool List", "path": "K:\\ERP TOOLS\\BELLATRX ERP ITEMS TOOL.xlsx"},
    {"label": "Equipment Labels Details", "path": "K:\\Equipment Labels\\Equipment Label - New"},
]

class Tooltip:
    def __init__(self, widget, get_text_callback, delay=500, movedelay=16):
        self.widget = widget
        self.get_text_callback = get_text_callback  # 动态获取文字
        self.delay = delay  # 延迟时间（毫秒）
        self.tooltip_window = None
        self.after_id = None
        self.last_update = 0  # 记录上次更新时间
        self.movedelay = movedelay  # 16ms ≈ 60FPS（相当于节流）

        self.widget.bind("<Enter>", self.schedule_show)
        self.widget.bind("<Leave>", self.hide_tooltip)
        self.widget.bind("<Motion>", self.throttle_move)

    def schedule_show(self, event):
        """安排显示提示"""
        if self.after_id:
            self.widget.after_cancel(self.after_id)
        self.after_id = self.widget.after(self.delay, lambda: self.show_tooltip(event))

    def show_tooltip(self, event=None):
        """显示 Tooltip"""
        if self.tooltip_window or not self.get_text_callback:
            return

        text = self.get_text_callback()
        if not text:
            return
        
        # 获取鼠标的位置
        x = event.x_root + int(10*sf)
        y = event.y_root + int(10*sf)
        # 创建 Tooltip 窗口
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # 去掉边框
        tw.attributes("-topmost", True)  # 确保窗口始终在最上层
        tw.wm_geometry(f"+{x}+{y}")

        # 创建 Tooltip Label
        label = ttk.Label(tw, text=text, style="Tooltip.TLabel", relief="solid", borderwidth=1, font=("Segoe UI", 9), padding=(5, 1, 5, 1))
        label.pack()

    def hide_tooltip(self, event=None):
        """隐藏 Tooltip"""
        if self.after_id:
            self.widget.after_cancel(self.after_id)
            self.after_id = None
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

    def throttle_move(self, event):
        """节流鼠标移动事件，减少 Tooltip 位置更新频率"""
        now = event.time
        if now - self.last_update < self.movedelay:
            return
        self.last_update = now
        self.update_position(event)

    def update_position(self, event):
        """更新 Tooltip 位置"""
        if self.tooltip_window:
            x = event.x_root + int(10*sf)
            y = event.y_root + int(10*sf)
            self.tooltip_window.wm_geometry(f"+{x}+{y}")

def show_warning_message(message, color):
    """在输入框下方显示警告信息"""
    global warning_label
    if warning_label is None:
        return
    warning_label.config(text=message, foreground=color)

def hide_warning_message():
    """隐藏警告信息"""
    global warning_label
    if warning_label:
        warning_label.config(text="")

def open_shortcut(index):
    """打开快捷访问的路径或文件"""
    entry.focus() # 焦点重新回到输入框
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
            if os.path.isdir(path):
                # 如果是目录，打开目录
                os.startfile(path)
            else:
                # 如果是文件，通过 open_file 打开
                open_file(file_path=path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open shortcut: {e}")
    else:
        messagebox.showerror("Error", f"Shortcut path does not exist: {path}")

def update_directory():
    """更新搜索目录"""
    global changed_parts_path, last_query
    new_dir = filedialog.askdirectory(initialdir=default_parts_path, title="Select Directory")
    if new_dir:
        new_dir = new_dir.replace('/', '\\')  # 将路径中的斜杠替换为反斜杠
        directory_label.config(text=f"PARTS Directory: {new_dir}")
        changed_parts_path = new_dir
        last_query = None
        
def reset_to_default_directory():
    """将搜索路径重置为默认路径"""
    global changed_parts_path, last_query
    directory_label.config(text=f"Default PARTS Directory: {default_parts_path}")
    changed_parts_path = None
    last_query = None

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
        history_listbox.insert(0, item) # 最新的搜索记录显示在列表最上面

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

def disable_search_button():
    """禁用所有搜索按钮"""
    search_btn.config(state=tk.DISABLED)
    lucky_btn.config(state=tk.DISABLED)
    search_3d_btn.config(state=tk.DISABLED)
    search_cache_btn.config(state=tk.DISABLED)

def enable_search_button():
    """启用所有搜索按钮"""
    search_btn.config(state=tk.NORMAL)
    lucky_btn.config(state=tk.NORMAL)
    search_3d_btn.config(state=tk.NORMAL)
    search_cache_btn.config(state=tk.NORMAL)

def build_directory_cache_thread(search_directory):
    """
    遍历指定目录，构建目录缓存。
    每个文件信息是一个元组：(文件名, 创建时间, 文件路径)
    """
    global directory_cache, active_threads
    thread = threading.current_thread()

    # 检查线程是否已经在运行
    if any(t.name == f"cache_thread_{search_directory}" for t in active_threads):
        return
    
    active_threads.add(thread)
    thread.name = f"cache_thread_{search_directory}"
    
    # 缓存开始，更改cache_label的颜色
    root.after(0, lambda: cache_label.config(foreground="red"))

    try:
        files_info = []
        for root_dir, _, files in os.walk(search_directory):
            if stop_event.is_set():  # 检查是否需要终止
                return
            for file in files:
                if stop_event.is_set():  # 检查是否需要终止
                    return
                if file.endswith((".pdf", ".iam", ".ipt")):
                    file_path = os.path.join(root_dir, file)
                    try:
                        create_time = os.path.getctime(file_path)
                    except Exception as e:
                        create_time = 0
                    files_info.append((file, create_time, file_path))

        if stop_event.is_set():  # 检查停止标志并返回
            return

        # 使用线程锁保护directory_cache
        with cache_lock:
            # 如果缓存数量超过 cache_max_size 限制，则删除最老的未使用项（LRU）
            if len(directory_cache) >= cache_max_size:
                directory_cache.popitem(last=False)  # 移除最老的未使用项

            # 添加新缓存数据，并将其移动到末尾（表示最近使用）
            directory_cache[search_directory] = files_info
            directory_cache.move_to_end(search_directory)

    except Exception as e:
        root.after(0, lambda: messagebox.showerror("Error", f"An error occurred in cache thread: {e}"))

    finally:
        active_threads.discard(thread)  # 线程结束后移除
        show_cache_status()

def get_cache_str():
    # 获取cache的目录并返回字符串
    cache_pattern = re.compile(r"cache_thread_.*?\\([^\\,]+)(?=\s|,|$)")
    caching_list = [match.group(1) for item in active_threads if (match := cache_pattern.search(str(item)))]
    cached_dir = [re.search(r'\\([^\\]+)$', key).group(1) for key in directory_cache.keys()]
    if not caching_list:
        # 没有正在缓存的目录
        if not cached_dir:
            # 没有已缓存的目录
            return f"No cache"
        else:
            # 有已缓存的目录
            return f"Cache completed: [{', '.join(cached_dir)}]"
    else:
        # 有正在缓存的目录
        if not cached_dir:
            # 没有已缓存的目录
            return f"Caching in progress: [{', '.join(caching_list)}]"
        else:
            # 有已缓存的目录
            return f"Caching in progress: [{', '.join(caching_list)}]\rCache completed: [{', '.join(cached_dir)}]"

def show_cache_status():
    # 获取cache的状态并设置颜色
    cache_pattern = re.compile(r"cache_thread_.*?\\([^\\,]+)(?=\s|,|$)")
    caching_list = [match.group(1) for item in active_threads if (match := cache_pattern.search(str(item)))]

    # 如果点击了重置按钮，直接改为灰色
    if stop_event.is_set():
        root.after(0, lambda: cache_label.config(foreground="lightgray"))
        return
    
    if not caching_list:
        root.after(0, lambda: cache_label.config(foreground="lime"))
    else:
        root.after(0, lambda: cache_label.config(foreground="red"))

def get_cached_directory(search_directory):
    """
    获取缓存的目录信息，如果存在则返回，否则返回 None。
    并将访问的缓存项移动到末尾（表示最近使用）
    """
    if search_directory in directory_cache:
        directory_cache.move_to_end(search_directory)
        return directory_cache[search_directory]
    return None

def search_files(query, search_type=None):
    """根据不同搜索类型，搜索PARTS目录下的文件"""
    global last_query

    # 检查进程是否存在，防止频繁搜索
    if any("search_files_thread" in t.name for t in active_threads):
        last_query = None
        return
    
    disable_search_button() # 禁用搜索按钮
    hide_warning_message()  # 清除警告信息
    
    
    if not query:
        show_warning_message("Please enter any number or project name!", "red")
        enable_search_button() # 启用搜索按钮
        last_query = None
        return

    # 检查是否包含非法字符
    if any(char in query for char in "*.?+^$[]{}|\\()"):
        show_warning_message("Invalid characters in search query!", "red")
        enable_search_button() # 启用搜索按钮
        last_query = None
        return

    save_search_history(query)  # 保存搜索记录
    last_query = query
    # 提取前两位字符并更新搜索路径
    prefix = query[:2]
    if changed_parts_path:
        search_directory = os.path.join(changed_parts_path, prefix)
    else:
        search_directory = os.path.join(default_parts_path, prefix)

    if not os.path.exists(search_directory):
        show_warning_message(f"Path does not exist! {search_directory}", "red")
        show_result_list(None) # 目录不存在就清空已有搜索结果
        enable_search_button() # 启用搜索按钮
        last_query = None
        return

    # 执行搜索
    show_warning_message(f"Searching... Please wait.", "red")
    query = query.lower()
    # 对STK的project number进行特殊处理
    if query.startswith("stk") and len(query) > 3:
        if query[3] == '-' or query[3] == ' ':
            query = query[:3] + '.*' + query[4:]
        else:
            query = query[:3] + '.*' + query[3:]

    stop_event.clear()  # 确保上一次的停止信号被清除
    search_thread = threading.Thread(target=search_files_thread, args=(query, search_directory, search_type))
    search_thread.start()

def search_files_thread(query, search_directory, search_type):
    """使用多线程搜索目录下的文件"""
    global active_threads, directory_cache, result_files_pdf, result_files_3d

    # 获取当前线程并添加到活动线程集合中
    thread = threading.current_thread()
    active_threads.add(thread)

    try:
        if "*" in query or "." in query:
            # 只有包含通配符才使用正则
            regex_pattern = re.compile(query, re.IGNORECASE)
            match_func = lambda file: regex_pattern.search(file)
        else:
            query = query.lower()
            match_func = lambda file: query in file.lower()
                
        result_files_pdf = []
        result_files_3d = []
        # 从缓存中取出所有文件信息
        all_files = get_cached_directory(search_directory)
        if all_files is None:
            # 如果缓存中没有该目录的记录，则在后台建立缓存
            # 检查是否已经有当前搜索目录的缓存线程在运行
            if not any(t.name == f"cache_thread_{search_directory}" for t in active_threads):
                cache_thread = threading.Thread(target=build_directory_cache_thread, args=(search_directory,))
                cache_thread.name = f"cache_thread_{search_directory}"
                cache_thread.start()

            # 直接遍历目录下的文件，不使用缓存
            i = 50
            for root_dir, _, files in os.walk(search_directory):
                if stop_event.is_set():  # 检查是否需要终止
                    return
                for file in files:
                    if stop_event.is_set():  # 检查是否需要终止
                        return
                    # 每遍历50个文件，显示一次文件名，体现搜索过程
                    if i == 50:
                        root.after(0, lambda: show_warning_message(f"Searching... Please wait.  {file}", "red"))
                        i = 0
                    i += 1
                    # 根据不同文件类型，存储到不同list里
                    if file.endswith(".pdf") and match_func(file):
                        file_path = os.path.join(root_dir, file)
                        create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                        result_files_pdf.append((file, create_time, file_path))  # (文件名, 创建时间, 文件路径)
                    elif (file.endswith(".iam") or file.endswith(".ipt")) and match_func(file):
                        file_path = os.path.join(root_dir, file)
                        create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                        result_files_3d.append((file, create_time, file_path))  # (文件名, 创建时间, 文件路径)

        else:
            # 使用缓存中的文件信息
            i = 50
            for file_info in all_files:
                if stop_event.is_set():  # 检查是否需要终止
                    return
                file_name = file_info[0]
                # 每遍历50个文件，显示一次文件名，体现搜索过程
                if i == 50:
                    root.after(0, lambda: show_warning_message(f"Searching... Please wait.  {file_name}", "red"))
                    i = 0
                i += 1
                # 根据不同文件类型，存储到不同list里
                if file_name.endswith(".pdf") and match_func(file_name):
                    # 格式化创建时间
                    create_time = datetime.datetime.fromtimestamp(file_info[1]).strftime("%Y-%m-%d %H:%M:%S")
                    result_files_pdf.append((file_name, create_time, file_info[2]))  # (文件名, 创建时间, 文件路径)
                elif (file_name.endswith(".iam") or file_name.endswith(".ipt")) and match_func(file_name):
                    # 格式化创建时间
                    create_time = datetime.datetime.fromtimestamp(file_info[1]).strftime("%Y-%m-%d %H:%M:%S")
                    result_files_3d.append((file_name, create_time, file_info[2]))  # (文件名, 创建时间, 文件路径)

        if stop_event.is_set():  # 检查停止标志并返回
            return
        # pdf搜索结果排序（按创建时间倒序）
        result_files_pdf.sort(key=lambda x: x[1], reverse=True)
        # 3d文件搜索结果排序（按文件名排序）
        result_files_3d.sort(key=lambda x: x[0])

        root.after(0, hide_warning_message)  # 使用主线程清除警告信息

        if search_type == "pdf" or search_type == "lucky":
            if not result_files_pdf:
                # 如果没有搜索到匹配的文件，显示警告信息
                root.after(0, lambda: show_warning_message("No matching drawing PDF found!", "red"))
            else:
                if search_type == "lucky":
                    # "I'm Feeling Lucky" 功能：直接打开第一个文件，一般按创建时间排序后就是最新的revision
                    file_path = result_files_pdf[0][2]  # 复制file_path的值传给open_file，避免result_files被修改后值为空
                    root.after(0, lambda: open_file(file_path=file_path))
                    # result_files = [] # 清空搜索结果，不显示在界面上 (注释原因：不清空，方便用户查看搜索结果)
            # 显示搜索结果
            root.after(0, lambda: show_result_list(result_files_pdf, search_type))
        if search_type == "3d":
            if not result_files_3d:
                # 如果没有搜索到匹配的文件，显示警告信息
                root.after(0, lambda: show_warning_message("No matching 3D drawing found! Try using Vault Cache.", "red"))
            # 显示搜索结果
            root.after(0, lambda: show_result_list(result_files_3d))
        
        root.after(0, lambda: enable_search_button())  # 启用搜索按钮

    except Exception as e:
        root.after(0, lambda: messagebox.showerror("Error", f"An error occurred in search thread: {e}"))

    finally:
        active_threads.discard(thread)  # 线程结束后移除

def search_pdf_files():
    """执行pdf搜索"""
    global last_query
    entry.focus()  # 保持焦点在输入框
    query = entry.get().strip() # 去除首尾空格
    if query == last_query:
        # 如果搜索关键字跟上一次一样，直接调用上一次的搜索结果
        if not result_files_pdf:
            # 如果结果是空，显示警告信息，移除之前显示的搜索结果
            show_warning_message("No matching drawing PDF found!", "red")
            close_result_list()
        else:
            disable_search_button()
            show_warning_message(f"Searching... Please wait.", "red")
            # 后台执行show_result_list
            root.after(10, lambda: (show_result_list(result_files_pdf, search_type="pdf"), hide_warning_message(), enable_search_button()))
    else:
        # 如果搜索关键字跟上一次不一样，重新搜索
        last_query = query
        search_files(query, search_type="pdf")

def feeling_lucky():
    """执行feeling lucky搜索"""
    global last_query
    entry.focus()  # 保持焦点在输入框
    query = entry.get().strip() # 去除首尾空格
    if query == last_query:
        # 如果搜索关键字跟上一次一样，直接调用上一次的搜索结果
        if not result_files_pdf:
            # 如果结果是空，显示警告信息，移除之前显示的搜索结果
            show_warning_message("No matching drawing PDF found!", "red")
            close_result_list()
        else:
            disable_search_button()
            show_warning_message(f"Searching... Please wait.", "red")
            file_path = result_files_pdf[0][2]  # 复制排在第一个的file_path的值传给open_file
            open_file(file_path=file_path)  # 打开第一个pdf文件
            # 后台执行show_result_list
            root.after(10, lambda: (show_result_list(result_files_pdf, search_type="lucky"), hide_warning_message(), enable_search_button()))
    else:
        # 如果搜索关键字跟上一次不一样，重新搜索
        last_query = query
        search_files(query, search_type="lucky")

def search_3d_files():
    """执行3d文件搜索"""
    global last_query
    entry.focus()  # 保持焦点在输入框
    query = entry.get().strip() # 去除首尾空格
    if query == last_query:
        # 如果搜索关键字跟上一次一样，直接调用上一次的搜索结果
        if not result_files_3d:
            # 如果结果是空，显示警告信息，移除之前显示的搜索结果
            show_warning_message("No matching 3D drawing found! Try using Vault Cache.", "red")
            close_result_list()
        else:
            disable_search_button()
            show_warning_message(f"Searching... Please wait.", "red")
            # 后台执行show_result_list
            root.after(10, lambda: (show_result_list(result_files_3d), hide_warning_message(), enable_search_button()))
    else:
        # 如果搜索关键字跟上一次不一样，重新搜索
        last_query = query
        search_files(query, search_type="3d")

def search_vault_cache():
    """搜索Vault缓存目录下的 3D 文件(ipt或者iam)"""
    global last_query
    entry.focus()  # 保持焦点在输入框
    disable_search_button() # 禁用搜索按钮
    hide_warning_message()  # 清除警告信息
    query = entry.get().strip() # 去除首尾空格
    if not query:
        show_warning_message("Please enter any number or project name!", "red")
        enable_search_button() # 启用搜索按钮
        return
    
    # 检查是否包含非法字符
    if any(char in query for char in "*.?+^$[]{}|\\()"):
        show_warning_message("Invalid characters in search query!", "red")
        enable_search_button() # 启用搜索按钮
        return

    save_search_history(query)  # 保存搜索记录

    matching_directories = []
    search_directory = None
    real_query = None
    last_query = None

    if not os.path.exists(vault_cache):
        # 如果Vault缓存目录不存在，提示用户使用Vault
        show_warning_message(f"Vault cache not found! Please use Vault instead.", "red")
        show_result_list(None) # 目录不存在就清空已有搜索结果
        enable_search_button() # 启用搜索按钮
        return

    if len(query) > 2 and query[2].isdigit():
        # 用户输入的是project number，可能包括2012789，2012789-100，2012789 100等格式
        # 也可能输入不完整的数字，在后面匹配目录时，用通配符*去匹配
        if " " in query:
            # 如果输入的project number带空格，替换为空格为-
            query = query.replace(" ", "-")
        if '-' in query:
            # 查询字段带后缀，提取project number
            proj_no = query.split('-')[0]
        else:
            proj_no = query
        sub_dir = proj_no
    elif query.lower().startswith("stk"):
        # 用户输入的是stk number，可能包括stk-100，stk 100，stk100等格式
        if len(query) > 3:
            if query[3] == '-' or query[3] == ' ':
                sub_dir = query[:3] + '*' + query[4:]
            else:
                sub_dir = query[:3] + '*' + query[3:]
        else:
            # 只输入了stk，没有具体的stk number
            sub_dir = query[:3] + '*'
    elif len(query) > 2 and query[0].isdigit() and not query[2].isdigit():
        # 用户输入的是part number或者assembly number，可能是22A010123或者13B010456
        # 需要到PARTS/22这样的路径下去搜索
        prefix = query[:2]
        search_directory = os.path.join(vault_cache, prefix)
        if not os.path.exists(search_directory):
            show_warning_message("No matching 3D drawings are cached. Check in Vault!", "red")
            show_result_list(None) # 目录不存在就清空已有搜索结果
            enable_search_button() # 启用搜索按钮
            return
    else:
        # 任何其他字符串，都当作是project name去匹配，去PARTS/S路径下查找匹配的目录
        if not os.path.exists(os.path.join(vault_cache, "S")):
            show_warning_message("No matching 3D drawings are cached. Check in Vault!", "red")
            show_result_list(None) # 目录不存在就清空已有搜索结果
            enable_search_button() # 启用搜索按钮
            return
        for dir_name in os.listdir(os.path.join(vault_cache, "S")):
            dir_path= os.path.join(vault_cache, "S", dir_name)
            if os.path.isdir(dir_path) and query.lower() in dir_name.lower():
                matching_directories.append(dir_path)
        # 根据匹配到的目录数量进行处理
        if len(matching_directories) == 1:
            search_directory = matching_directories[0]
            if not query[0].isdigit() and not query.lower().startswith("stk"):
                sub_dir = os.path.basename(search_directory)
                path_split = sub_dir.find(" ")
                if path_split != -1:
                    real_query = sub_dir[1:path_split]
                else:
                    real_query = sub_dir[1:]
        elif len(matching_directories) > 1:
            # 如果有匹配到多个目录，让用户选择
            selected_dir = ask_user_to_select_directory(matching_directories)
            if selected_dir:
                search_directory = selected_dir
                sub_dir = os.path.basename(search_directory)
                if sub_dir.lower().startswith("stk"):
                    # 选择了STK开头的project
                    if sub_dir[3] == " " and sub_dir[4].isdigit():
                        # 处理“STK 103 Project Name” 这样的project number
                        sub_dir_split = sub_dir.split(" ")
                        if len(sub_dir_split) >= 2:
                            query = " ".join(sub_dir_split[:2])
                    else:
                        # 处理“STK-103 Project Name” 这样的project number
                        end_index = sub_dir.find(" ")
                        if end_index != -1:
                            query = sub_dir[0:end_index]
                        else:
                            # 处理"STK103"这样的project number
                            query = sub_dir
                else:
                    if sub_dir.lower().startswith("s"):
                        # 选择了S开头的project
                        end_index = sub_dir.find(" ")
                        if end_index != -1:
                            query = sub_dir[1:end_index]
                    else:
                        # 如果存在既不是STK开头的，也不是S开头的project目录
                        end_index = sub_dir.find(" ")
                        if end_index != -1:
                            query = sub_dir[0:end_index]
                        else:
                            query = sub_dir
            else:
                show_warning_message("Cancelled!", "red")
                enable_search_button() # 启用搜索按钮
                return
        else:
            # 如果用户输入的关键字匹配不到任何project，直接当作子目录去PARTS下搜索，如PARTS/XX
            prefix = query[:2]
            search_directory = os.path.join(vault_cache, prefix)
            if not os.path.exists(search_directory):
                show_warning_message("No matching 3D drawings are cached. Check in Vault!", "red")
                show_result_list(None) # 目录不存在就清空已有搜索结果
                enable_search_button() # 启用搜索按钮
                return

    if not search_directory:
        prefix = "S"
        matching_directories = glob.glob(os.path.join(vault_cache, prefix, '*' + sub_dir + '*'))
        # 根据匹配到的目录数量进行处理
        if len(matching_directories) == 1:
            search_directory = matching_directories[0]
        elif len(matching_directories) > 1:
            # 如果有匹配到多个目录，让用户选择
            selected_dir = ask_user_to_select_directory(matching_directories)
            if selected_dir:
                search_directory = selected_dir
                sub_dir = os.path.basename(search_directory)
                if sub_dir.lower().startswith("stk"):
                    # 选择了STK开头的project
                    if sub_dir[3] == " " and sub_dir[4].isdigit():
                        # 处理“STK 103 Project Name” 这样的project number
                        sub_dir_split = sub_dir.split(" ")
                        if len(sub_dir_split) >= 2:
                            query = " ".join(sub_dir_split[:2])
                    else:
                        # 处理“STK-103 Project Name” 这样的project number
                        end_index = sub_dir.find(" ")
                        if end_index != -1:
                            query = sub_dir[0:end_index]
                        else:
                            # 处理"STK103"这样的project number
                            query = sub_dir
                else:
                    if sub_dir.lower().startswith("s"):
                        # 选择了S开头的project
                        end_index = sub_dir.find(" ")
                        if end_index != -1:
                            query = sub_dir[1:end_index]
                    else:
                        # 如果存在既不是STK开头的，也不是S开头的project目录
                        end_index = sub_dir.find(" ")
                        if end_index != -1:
                            query = sub_dir[0:end_index]
                        else:
                            query = sub_dir
            else:
                show_warning_message("Cancelled!", "red")
                enable_search_button() # 启用搜索按钮
                return
        else:
            show_warning_message("No matching 3D drawings are cached. Check in Vault!", "red")
            show_result_list(None) # 目录不存在就清空已有搜索结果
            enable_search_button() # 启用搜索按钮
            return

    # 执行搜索
    show_warning_message(f"Searching... Please wait.", "red")
    if query.lower().startswith("stk"):
        if len(query) > 3:
            if query[3] == '-' or query[3] == ' ':
                query = query[:3] + '*' + query[4:]
            else:
                query = query[:3] + '*' + query[3:]
        else:
            # 只输入了stk，没有具体的stk number
            query = query[:3] + '*'
    
    if real_query:
        query = real_query

    stop_event.clear()  # 确保上一次的停止信号被清除
    search_thread = threading.Thread(target=search_vault_cache_thread, args=(query, search_directory,))
    search_thread.start()

def search_vault_cache_thread(query, search_directory):
    """使用多线程搜索Vault缓存目录下的 3D 文件"""
    global active_threads
    thread = threading.current_thread()
    active_threads.add(thread)
    try:
        if "*" in query or "." in query:
            # 只有包含通配符才使用正则
            regex_pattern = re.compile(query.replace("*", ".*"), re.IGNORECASE)
            match_func = lambda file: regex_pattern.search(file)
        else:
            query = query.lower()
            match_func = lambda file: query in file.lower()

        result_files = []
        # 替换通配符为正则表达式
        i = 50
        for root_dir, _, files in os.walk(search_directory):
            if stop_event.is_set():  # 检查是否需要终止
                return
            for file in files:
                if stop_event.is_set():  # 检查是否需要终止
                    return
                # 每遍历50个文件，显示一次文件名，体现搜索过程
                if i == 50:
                    root.after(0, lambda: show_warning_message(f"Searching... Please wait.  {file}", "red"))
                    i = 0
                i += 1
                if (file.endswith(".iam") or file.endswith(".ipt")) and match_func(file):
                    file_path = os.path.join(root_dir, file)
                    create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                    result_files.append((file, create_time, file_path))

        if stop_event.is_set():  # 检查停止标志并返回
            return
        root.after(0, hide_warning_message)  # 使用主线程清除警告信息

        if not result_files:
            root.after(0, lambda: show_warning_message("No matching 3D drawings are cached. Check in Vault!", "red"))
        else:
            root.after(0, lambda: show_warning_message("Tip: Searched from cache, may not be the latest update!", "blue"))
        if len(query) > 2 and query[2].isdigit():
            # 如果是project number，按文件名正序排列
            result_files.sort(key=lambda x: x[0])
        else:
            # 如果是part number或者assembly number，按创建时间倒序排列
            result_files.sort(key=lambda x: x[1], reverse=True)

        root.after(0, lambda: show_result_list(result_files))
        root.after(0, lambda: enable_search_button())  # 启用搜索按钮
    except Exception as e:
        root.after(0, lambda: messagebox.showerror("Error", f"An error occurred in search thread: {e}"))

    finally:
        active_threads.discard(thread)  # 线程结束后移除

def ask_user_to_select_directory(directories):
    """弹出对话框让用户选择project目录"""

    def on_list_select(event):
        # 列表单击激活选择按钮
        select_btn.config(state=tk.NORMAL)

    def on_double_click(event):
        # 列表双击选择
        selected_index = event.widget.curselection()
        if selected_index:
            selected_dir[0] = directories[selected_index[0]]
            choice_win.destroy()

    def on_select():
        # 按钮选择
        selected_index = listbox.curselection()
        if selected_index:
            selected_dir[0] = directories[selected_index[0]]
            choice_win.destroy()

    choice_win = tk.Toplevel(root)
    choice_win.withdraw()  # 先隐藏窗口
    choice_win.attributes("-topmost", True)
    choice_win.iconphoto(True, icon) # 设置窗口图标
    choice_win.title("Select Project")
    choice_win_width = int(280*sf)
    choice_win_height = int(200*sf)
    choice_win.geometry(f"{choice_win_width}x{choice_win_height}")
    choice_win.resizable(False, False)

    # 窗口位置，跟随主窗口初始大小居中显示，方便鼠标选取
    choice_win.update_idletasks()
    position_right = int(root.winfo_x() + window_width/2 - choice_win_width/2)
    position_down = int(root.winfo_y() + window_height/2 - choice_win_height/2)
    choice_win.geometry(f"+{position_right}+{position_down}")
    choice_win.deiconify() # 显示窗口

    # 主容器
    frame = tk.Frame(choice_win)
    frame.pack(fill=tk.BOTH, expand=True, padx=(int(15*sf), int(5*sf)), pady=int(10*sf))

    # 提示文本
    label = ttk.Label(frame, text="Multiple projects were found, please select:", font=("Segoe UI", 9), anchor="w")
    label.pack(fill=tk.X)

    # 目录列表框
    list_frame = tk.Frame(frame)
    list_frame.pack(fill=tk.BOTH, expand=True, pady=int(5*sf))
    listbox = tk.Listbox(list_frame, width=20, height=6, font=("Segoe UI", 9))
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=int(5*sf))

    # 创建Scrollbar并将其与Listbox关联
    scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox.config(yscrollcommand=scrollbar.set)

    # 填充目录列表
    for path in directories:
        display_name = os.path.basename(path)  # 只显示目录名
        if display_name.lower().startswith("s") and not display_name.lower().startswith("stk"):
            display_name = display_name[1:]
        listbox.insert(tk.END, f" {display_name}")  # 每行最前面留了一个空格，显示更好

    for i in range(listbox.size()):
        if i % 2 == 0:
            listbox.itemconfig(i, {'bg': '#E6F7FF'})  # 浅蓝色
        else:
            listbox.itemconfig(i, {'bg': 'white'})  # 白色

    selected_dir = [None]  # 用列表存储选择结果
    listbox.bind("<<ListboxSelect>>", on_list_select)  # 单击事件
    listbox.bind("<Double-1>", on_double_click)  # 双击事件
    btn_frame = tk.Frame(frame)
    btn_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=int(15*sf), pady=int(3*sf))

    cancel_btn = ttk.Button(btn_frame, text="Cancel", width=8, style="All.TButton", command=choice_win.destroy)
    cancel_btn.pack(side=tk.RIGHT, padx=0)

    select_btn = ttk.Button(btn_frame, text="Select Project", width=12, style="All.TButton", command=on_select)
    select_btn.pack(side=tk.RIGHT, padx=int(15*sf))
    select_btn.config(state=tk.DISABLED) # 默认禁用选择按钮

    # 等待窗口关闭
    choice_win.wait_window()
    return selected_dir[0]

def close_result_list():
    """移除搜索结果"""
    global result_frame, results_tree, window_expanded, thumbnail_check, thumbnail_win
    if result_frame:
        results_tree.destroy()
        results_tree = None
        result_frame.destroy()
        result_frame = None
        thumbnail_check.forget()
        if thumbnail_win and thumbnail_win.winfo_exists():
            thumbnail_win.destroy()  # 关闭缩略图窗口
        # 取反窗口扩展标志位，通过toggle_window_size()保持当前状态
        window_expanded = not window_expanded
        toggle_window_size()

def sort_treeview(col, columns):
    # 排序函数
    global results_tree
    # 获取 Treeview 中的所有数据
    data_list = [(results_tree.item(item, "values"), item) for item in results_tree.get_children("")]
    
    # 按照指定列进行排序
    col_index = columns.index(col)  # 获取列索引
    # 读取当前列的排序状态
    reverse = results_tree.sort_states[col]
    sorted_data = sorted(data_list, key=lambda x: x[0][col_index], reverse=reverse)

    # 删除原数据
    for item in results_tree.get_children(""):
        results_tree.delete(item)

    # 重新插入排序后的数据，并应用交错颜色
    for index, (values, _) in enumerate(sorted_data):
        tag = 'evenrow' if index % 2 == 0 else 'oddrow'
        results_tree.insert("", tk.END, values=values, tags=(tag,))

    # 切换排序状态
    results_tree.sort_states[col] = not reverse

    # 更新表头箭头
    arrow = "  ▲" if not reverse else "  ▼"
    for c in columns:
        results_tree.heading(c, text=c)  # 先重置所有表头
    results_tree.heading(col, text=col + arrow, command=lambda: sort_treeview(col, columns))

def show_result_list(result_files, search_type=None):
    """显示搜索结果"""
    global result_frame, results_tree, thumbnail_check, thumbnail_win

    # 销毁原先的缩略图窗口
    if thumbnail_win and thumbnail_win.winfo_exists():
        thumbnail_win.destroy()

    # 如果没有搜索结果，移除搜索结果显示区域
    if not result_files:
        if result_frame:
            results_tree.destroy()
            results_tree = None
            result_frame.destroy()
            result_frame = None
            if 'thumbnail_check' in globals() and thumbnail_check:
                thumbnail_check.forget()  # 如果没有搜索结果，隐藏 thumbnail_check
            if window_expanded:
                root.geometry(f"{expand_window_width}x{window_height}")
            else:
                root.geometry(f"{window_width}x{window_height}")
        return

    # 显示搜索结果数量
    count = len(result_files)
    msg = f"{count} file{'s' if count!=1 else ''} found. Double-click to open."

    # 创建结果显示区域
    if result_frame:
        results_tree.destroy()
        results_tree = None
        result_frame.destroy()
        result_frame = None
    result_frame = tk.Frame(root)
    result_frame.pack(fill=tk.BOTH, expand=True, pady=0)
    tip_frame = tk.Frame(result_frame)
    tip_frame.pack(padx=0, pady=0, fill="x")
    # 搜索结果数量
    tip_label = ttk.Label(tip_frame, text=msg, font=("Segoe UI", 9), foreground="blue")
    tip_label.pack(padx=int(20*sf), pady=0, side=tk.LEFT)
    # 显示一个关闭按钮用来移除搜索结果
    close_btn = ttk.Button(tip_frame, text="❌", style="Close.TButton", width=3, command=close_result_list)
    close_btn.pack(padx=(0, int(16*sf)), pady=0, side=tk.RIGHT)
    Tooltip(close_btn, lambda: "Remove search results list", delay=500)

    # 如果是搜索pdf文件，就显示thumbnail_check按钮
    if thumbnail_check and search_type in ("pdf", "lucky"):
        thumbnail_check.pack(side=tk.LEFT, padx=int(20*sf))
    else:
        thumbnail_check.forget()

    # 设置 Treeview 表头和行样式
    style = ttk.Style()
    style.configure("Treeview.Heading", padding=(0, int(4*sf)), background="#A9A9A9", foreground="black", font=("Segoe UI", 9, "bold"))
    style.configure("Treeview", font=("Segoe UI", 9), rowheight=int(25*sf))
    style.map("Treeview", background=[('selected', '#347083')])

    # 添加 Treeview 控件显示结果
    columns = ("File Name", "Created Time", "Path")
    results_tree = ttk.Treeview(result_frame, columns=columns, show="headings")
    results_tree.pack(fill=tk.BOTH, expand=True, padx=(int(17*sf), 0), pady=0)
    # 在 results_tree 上存储排序状态
    results_tree.sort_states = {col: False for col in columns}  # False 表示升序, True 表示降序
    for col in columns:
        results_tree.heading(col, text=col, anchor="w", command=lambda c=col: sort_treeview(c, columns))
    results_tree.column("File Name", width=150, anchor="w")
    results_tree.column("Created Time", width=135, anchor="w")
    results_tree.column("Path", width=0, stretch=tk.NO)  # 隐藏第三列

    # 创建一个垂直滚动条并将其与 Treeview 关联
    scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=results_tree.yview)
    results_tree.configure(yscrollcommand=scrollbar.set)

    # 插入搜索结果
    for index, item in enumerate(result_files):
        tag = 'evenrow' if index % 2 == 0 else 'oddrow'
        results_tree.insert("", tk.END, values=(item[0], item[1], item[2]), tags=(tag,))
    
    results_tree.tag_configure('evenrow', background='#E6F7FF')
    results_tree.tag_configure('oddrow', background='white')

    results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    results_tree.bind("<Double-1>", open_file)
    results_tree.bind("<<TreeviewSelect>>", on_tree_select)

    # 动态调整窗口大小以显示结果
    root.update_idletasks()
    new_height = int(420*sf) + len(result_files) * int(25*sf) if result_files else window_height
    if window_expanded:
        root.geometry(f"{expand_window_width}x{min(new_height, int(540*sf))}")
    else:
        root.geometry(f"{window_width}x{min(new_height, int(540*sf))}")

def get_pdf_page_orientation(pdf_path):
    """获取 PDF 第一页的方向（横向 or 纵向）"""
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]  # 获取第一页
        width, height = page.rect.width, page.rect.height
        # 判断pdf是横向还是竖向
        if width > height:
            return "landscape", width, height
        else:
            return "portrait", height, width
    except Exception as e:
        show_warning_message(f"Unknown PDF orientation: {e}", "red")
        return "landscape", width, height # 失败时默认横向

def generate_pdf_thumbnail(pdf_path, thumbnail_size=(220, 170)):
    """生成 PDF 文件的缩略图，220x170是根据letter纸张比例设置"""
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]  # 读取第一页
        pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))  # 缩小 50% 生成更小的图片
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail(thumbnail_size)  # 生成缩略图
        return ImageTk.PhotoImage(img)
    except Exception as e:
        show_warning_message(f"Unable to generate PDF thumbnail: {e}", "red")
        return None

def on_tree_select(event):
    """当选中某个搜索结果时，如果是 PDF 文件，则在独立窗口显示缩略图"""
    global thumbnail_win, results_tree, thumbnail_check, thumbnail_var

    hide_warning_message()  # 清除警告信息

    def close_window(event=None):
        # 关闭缩略图窗口
        global thumbnail_win
        if thumbnail_win and thumbnail_win.winfo_exists():
            thumbnail_win.destroy()
            thumbnail_win = None

    selected_item = results_tree.selection()
    if not selected_item:
        return
    # 如果复选框存在且未勾选，关闭缩略图显示并返回
    if 'thumbnail_check' in globals() and thumbnail_check and not thumbnail_var.get():
        if thumbnail_win and thumbnail_win.winfo_exists():
            thumbnail_win.destroy()
            thumbnail_win = None
        return
    file_path = results_tree.item(selected_item, "values")[2]  # 获取文件路径

    if file_path.lower().endswith(".pdf"):
        orientation, width, height = get_pdf_page_orientation(file_path)  # 判断 PDF 方向，获取长宽数据
        # 缩略图缩小1/5
        long_edge = int(width / 5 * sf)
        short_edge = int(height / 5 *sf)
        if orientation == "landscape":
            thumbnail = generate_pdf_thumbnail(file_path, (long_edge, short_edge))
        else:
            thumbnail = generate_pdf_thumbnail(file_path, (short_edge, long_edge))
        if thumbnail:
            if thumbnail_win and thumbnail_win.winfo_exists():
                thumbnail_win.destroy()  # 先销毁旧窗口

            # 创建一个新的独立窗口
            thumbnail_win = tk.Toplevel(root)
            thumbnail_win.configure(bg="orange")
            thumbnail_win.overrideredirect(True)
            if topmost_var.get():
                # 如果主窗口置顶，缩略图窗口也置顶
                thumbnail_win.attributes("-topmost", True)
            # 根据纸张方向设置窗口的大小
            if orientation == "landscape":
                thumbnail_win_width = long_edge + int(10*sf)
                thumbnail_win_height = short_edge + int(10*sf)
            else:
                thumbnail_win_width = short_edge + int(10*sf)
                thumbnail_win_height = long_edge + int(10*sf)
            thumbnail_win.geometry(f"{thumbnail_win_width}x{thumbnail_win_height}")  # 设置窗口大小
            thumbnail_win.resizable(False, False)

            # 显示缩略图
            label = ttk.Label(thumbnail_win, image=thumbnail, anchor="center")
            label.pack(padx=int(5*sf), pady=int(5*sf))
            label.image = thumbnail  # 保持引用，防止被垃圾回收

            # 用于关闭缩略图窗口的label
            close_label = ttk.Label(thumbnail_win, text="✕", style="Close.TLabel")
            close_label.place(relx=1.0, x=int(-5*sf), y=int(5*sf), anchor="ne")  # 右上角
            close_label.bind("<Button-1>", close_window)  # 绑定点击事件

            # 缩略图窗口出现主窗口左侧
            x = root.winfo_x() - thumbnail_win_width + int(7*sf)
            y = root.winfo_y() + window_height + int(48*sf)
            thumbnail_win.geometry(f"+{x}+{y}")
    else:
        if thumbnail_win and thumbnail_win.winfo_exists():
            thumbnail_win.destroy()  # 关闭缩略图窗口

def on_main_window_move(event):
    """当主窗口移动时，让 thumbnail_win 也随之移动"""
    global thumbnail_win
    if thumbnail_win and thumbnail_win.winfo_exists():
        # 计算新的位置
        x = root.winfo_x() - thumbnail_win.winfo_width() + int(7*sf)
        y = root.winfo_y() + window_height + int(48*sf)
        thumbnail_win.geometry(f"+{x}+{y}")

def on_focus_in(event):
    # 如果焦点在主窗口，调出缩略图窗口
    global thumbnail_win
    if root.state() != "iconic" and root.state() != "withdrawn":
        if thumbnail_win and thumbnail_win.winfo_exists():
            if thumbnail_win.state() == "withdrawn":  # 当窗口被隐藏时重新显示
                thumbnail_win.deiconify()
            thumbnail_win.lift()  # 提升缩略图窗口到最前

def on_window_state_change(event=None):
    # 当窗口状态改变时，隐藏或显示缩略图窗口
    global thumbnail_win
    if root.state() == "iconic" or root.state() == "withdrawn":  # 窗口最小化或隐藏
        if thumbnail_win and thumbnail_win.winfo_exists():
            thumbnail_win.withdraw()  # 隐藏缩略图窗口
    else:  # 窗口恢复
        if thumbnail_win and thumbnail_win.winfo_exists():
            thumbnail_win.deiconify()  # 显示缩略图窗口
            thumbnail_win.lift()  # 提升缩略图窗口到最前

def show_about():
    """自定义关于信息的窗口"""
    entry.focus()  # 保持焦点在输入框
    if hasattr(root, 'about_win') and root.about_win.winfo_exists():
        # 如果窗口已经打开，就返回
        root.about_win.lift()  # 如果存在，提升窗口到最前
        return

    # 创建自定义关于窗口
    about_win = tk.Toplevel(root)
    root.about_win = about_win  # 将窗口绑定到 root 的属性上
    about_win.withdraw()  # 先隐藏窗口
    about_win.attributes("-topmost", True)
    about_win.title("About")
    about_win_width = int(370*sf)
    about_win_height = int(275*sf)
    about_win.geometry(f"{about_win_width}x{about_win_height}")
    about_win.resizable(False, False)

    # 窗口关闭时重置标志位
    def on_close():
        del root.about_win  # 删除引用
        about_win.destroy()  

    about_win.protocol("WM_DELETE_WINDOW", on_close)

    # 窗口位置，跟随主窗口居中显示，不考虑Treeview高度
    about_win.update_idletasks()
    position_right = int(root.winfo_x() + root.winfo_width()/2 - about_win_width/2)
    position_down = int(root.winfo_y() + window_height - about_win_height + 6)
    about_win.geometry(f"+{position_right}+{position_down}")
    about_win.deiconify() # 显示窗口
    
    # 设置窗口图标（复用主窗口图标）
    about_win.iconphoto(True, icon)
    
    # 主容器框架
    main_frame = tk.Frame(about_win)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=int(10*sf), pady=int(10*sf))

    # 左侧图标区域
    icon_frame = tk.Frame(main_frame, width=100)
    icon_frame.pack(side=tk.LEFT, fill=tk.Y, padx=int(15*sf))
    
    try:
        # 解码Base64图标并调整大小
        icon_data = base64.b64decode(ICON_BASE64)
        img = Image.open(io.BytesIO(icon_data))
        img = img.resize((int(64*sf), int(64*sf)), Image.Resampling.LANCZOS)  # 调整图标尺寸
        tk_img = ImageTk.PhotoImage(img)
        icon_label = ttk.Label(icon_frame, image=tk_img)
        icon_label.image = tk_img  # 保持引用
        icon_label.pack(pady=int(30*sf))
    except Exception as e:
        print(f"Error loading icon: {e}")

    # 右侧文本区域
    text_frame = tk.Frame(main_frame)
    text_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=int(15*sf))

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
            label = ttk.Label(text_frame, text=text, font=("Segoe UI", 10, "bold"), anchor="w")
        else:
            label = ttk.Label(text_frame, text=text, font=("Segoe UI", 9), anchor="w")
        label.pack(anchor="w", fill=tk.X)

    # 邮箱按钮和地址
    email_frame = tk.Frame(text_frame)
    email_frame.pack(anchor="w")

    email_label = ttk.Label(email_frame, text="Email me", foreground="blue", cursor="hand2", font=("Segoe UI", 9, "underline"))
    email_label.pack(side=tk.LEFT, padx=0)
    email_label.bind("<Button-1>", lambda event: send_email())

    email_label = ttk.Label(email_frame, text=": wtweitang@hotmail.com", font=("Segoe UI", 9))
    email_label.pack(side=tk.LEFT)

    # 设置OK按钮的样式，主要想增加按钮高度，以便于不移动鼠标即可点击按钮关闭窗口
    style.configure("AboutOK.TButton", font=("Segoe UI", 9), padding=(5, 5))
    ok_button = ttk.Button(about_win, text="OK", style="AboutOK.TButton", command=on_close)
    ok_button.pack(padx=int(15*sf), pady=int(15*sf), side=tk.RIGHT)

def send_email():
    """打开默认邮件客户端发送邮件"""
    import webbrowser
    try:
        webbrowser.open("mailto:wtweitang@hotmail.com?subject=Drawing%20Search%20Feedback")
    except Exception as e:
        messagebox.showerror("Error", f"Cannot open email client: {e}")

def reset_window():
    """恢复主窗口到初始状态，停止搜索进程，清空缓存"""
    global result_frame, results_tree, window_expanded, shortcut_frame, last_query, thumbnail_win

    entry.focus()  # 保持焦点在输入框

    # 触发停止事件
    stop_event.set()

    # 等待所有线程结束
    for thread in list(active_threads):
        thread.join(timeout=0.5)  # 最多等 0.5 秒
    
    # 清除所有线程引用
    active_threads.clear()
    
    # 重置停止事件，以便下一次搜索可以正常启动
    stop_event.clear()

    # 清空目录缓存
    directory_cache.clear()

    # 重置cache label颜色
    cache_label.config(foreground="lightgray")

    # 隐藏 thumbnail_check
    thumbnail_check.forget()

    # 关闭缩略图窗口
    if thumbnail_win and thumbnail_win.winfo_exists():
        thumbnail_win.destroy()

    # 清除上次搜索关键字记录
    last_query = None
    
    entry.delete(0, tk.END)  # 清空输入框
    hide_warning_message()  # 清除警告信息
    enable_search_button() # 启用搜索按钮

    # 移除显示的搜索结果
    if result_frame:
        results_tree.destroy()
        results_tree = None
        result_frame.destroy()
        result_frame = None
    root.geometry(f"{window_width}x{window_height}")  # 恢复初始窗口大小
    if window_expanded:
        expand_btn.config(text="Quick Access   ❯❯")  # 改为 "❯❯"
        window_expanded = not window_expanded  # 切换状态
        if shortcut_frame:
            shortcut_frame.destroy()
            shortcut_frame = None

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

    entry.focus()  # 保持焦点在输入框
    if window_expanded:
        # 如果已有搜索结果，保持显示搜索结果
        if result_frame:
            height = root.winfo_height()  # 获取当前窗口高度
            root.geometry(f"{window_width}x{height}")
        else:
            # 收缩窗口，隐藏快捷按钮框架
            root.geometry(f"{window_width}x{window_height}")  # 恢复到原始大小
        expand_btn.config(text="Quick Access   ❯❯")  # 改为 "❯❯"
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
        expand_btn.config(text="Quick Access   ❮❮")  # 改为 "❮❮"

        # 先清空快捷按钮框架，防止从 mini 窗口切换回来时重复生成
        if shortcut_frame:
            shortcut_frame.destroy()
            shortcut_frame = None

        # 创建快捷按钮框架
        shortcut_frame = tk.Frame(root)
        shortcut_frame.place(x=int(340*sf), y=int(43*sf), width=int(200*sf), height=int(210*sf))  # 定位到右侧扩展区域

        for i, shortcut in enumerate(shortcut_paths):
            btn = ttk.Button(
                shortcut_frame, 
                text=shortcut["label"],
                width=100,
                style="All.TButton",
                command=lambda i=i: open_shortcut(i)
            )
            btn.pack(padx=int(10*sf), pady=int(5*sf), anchor="w")

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

def toggle_topmost():
    # 根据复选框的状态设置窗口是否置顶
    entry.focus()  # 保持焦点在输入框
    is_checked = topmost_var.get()
    root.attributes("-topmost", is_checked)
    # 如果缩略图窗口存在，也设置其置顶
    if thumbnail_win and thumbnail_win.winfo_exists():
        thumbnail_win.attributes("-topmost", is_checked)

def create_entry_context_menu(entry_widget):
    # 为 Entry 小部件创建一个右键菜单
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

def open_mini_window():
    # 打开 mini 窗口
    # 隐藏缩略图窗口
    if thumbnail_win and thumbnail_win.winfo_exists():
        thumbnail_win.withdraw()
    # 隐藏主窗口
    root.withdraw()
    
    def on_close():
        # 关闭 mini 窗口时，获取当前的位置，用于主窗口显示
        global window_expanded
        # 获取当前主窗口是否扩展的状态
        expanded_status = window_expanded
        new_position_left = int(mini_win.winfo_x() - root.winfo_width()/2 + mini_win_width/2)
        new_position_top = mini_win.winfo_y()
        mini_win.destroy()
        show_window(new_position_left, new_position_top, expanded_status)

    def clear_mini_entry(event=None):
        # 清空输入框内容
        mini_entry.delete(0, tk.END)  # 清空输入框内容
        clear_label_mini.place_forget()  # 清空后，隐藏 X

    def update_label_color_mini(event=None):
        # 判断输入框内容是否为空
        if mini_entry.get():
            clear_label_mini.place(in_=mini_entry, relx=1.0, rely=0.5, anchor='e', x=int(-10*sf))  # 显示 X
        else:
            clear_label_mini.place_forget()  # 空内容时，隐藏 X`

    # 创建 mini 窗口
    mini_win = tk.Toplevel(root)
    mini_win.withdraw()  # 先隐藏窗口
    mini_win.title("Drawing Search")
    mini_win_width = int(230*sf)
    mini_win_height = int(35*sf)
    mini_win.geometry(f"{mini_win_width}x{mini_win_height}")
    mini_win.attributes("-topmost", True) # 窗口置顶
    mini_win.attributes('-alpha', 0.6)  # 设置窗口透明度
    mini_win.resizable(False, False)

    # 设置窗口图标（复用主窗口图标）
    mini_win.iconphoto(True, icon)
    # 窗口位置，跟随主窗口居中显示
    mini_win.update_idletasks()
    position_left = int(root.winfo_x() + root.winfo_width()/2 - mini_win_width/2)
    position_top = root.winfo_y()
    mini_win.geometry(f"+{position_left}+{position_top}")
    mini_win.deiconify() # 显示mini窗口
    root.winfo_width()/2 - mini_win_width/2
    # 创建 mini 窗口的框架
    mini_frame = tk.Frame(mini_win)
    mini_frame.pack(pady=int(5*sf))

    # 在框架中添加一个输入框
    mini_entry = ttk.Entry(mini_frame, font=("Segoe UI", 12), width=13)
    mini_entry.pack(side="left", pady=0, padx=int(5*sf))
    mini_entry.focus()

    # 定义 mini 窗口的搜索操作
    def on_search_mini(event=None):
        mini_entry.focus() # 焦点回位到输入框
        query = mini_entry.get().strip()
        if query:
            # 将 mini 窗口输入内容传递到主窗口的输入框
            entry.delete(0, tk.END)
            entry.insert(0, query)
            on_close() # 关闭 mini 窗口，显示主窗口
            # 调用搜索pdf函数
            search_pdf_files()
    
    # 绑定回车键
    mini_entry.bind("<Return>", on_search_mini)

    # 清除 mini Entry 内容的Label，初始不显示
    clear_label_mini = ttk.Label(mini_entry, text="✕", font=('Segoe UI', 9), foreground='red', style="Clear.TLabel", cursor="arrow")
    # 监听 mini Entry 内容变化来更新 Label 的颜色， 同时实时更新匹配历史
    mini_entry.bind("<KeyRelease>", lambda event: (update_label_color_mini(event)))
    mini_entry.bind("<FocusIn>", update_label_color_mini)     # 当 mini Entry 获取焦点时检查内容

    # 绑定点击事件：点击 Label 的 X 就清空 Entry
    clear_label_mini.bind("<Button-1>", clear_mini_entry)

    # 添加搜索按钮
    search_btn_mini = ttk.Button(mini_frame, text="Search", width=10, style="All.TButton", command=on_search_mini)
    search_btn_mini.pack(side="right", padx=int(5*sf))

    # 如果用户直接关闭 mini 窗口，则重新显示主窗口
    mini_win.protocol("WM_DELETE_WINDOW", on_close)

def show_window(new_position_left, new_position_top, expanded_status):
    # 根据新位置显示主窗口
    global window_expanded
    root.geometry(f"{window_width}x{window_height}+{new_position_left}+{new_position_top}")
    # 如果主窗口隐藏前是扩展状态，恢复到扩展状态，并恢复搜索结果的显示（如果有） 
    if expanded_status:
        window_expanded = False
        toggle_window_size()
    else:
        window_expanded = True
        toggle_window_size()
    root.deiconify()  # 显示root窗口
    if thumbnail_win and thumbnail_win.winfo_exists():
        thumbnail_win.deiconify()  # 显示缩略图窗口
        thumbnail_win.lift()
    entry.focus_set()  # 设置焦点到输入框

def on_root_close():
    # 关闭主窗口时清除所有未完成的线程
    global thumbnail_win
    stop_event.set()  # 发送退出信号
    # 等待所有线程结束
    for thread in list(active_threads):
        thread.join(timeout=0.5)  # 最多等 0.5 秒
    if thumbnail_win and thumbnail_win.winfo_exists():
        thumbnail_win.destroy()  # 关闭缩略图窗口
    root.destroy()  # 关闭窗口

def clear_entry(event=None):
    # 清空输入框内容
    entry.delete(0, tk.END)  # 清空输入框内容
    clear_label.place_forget()  # 清空后，隐藏 X

def update_label_color(event=None):
    # 判断输入框内容是否为空
    if entry.get():
        clear_label.place(in_=entry, relx=1.0, rely=0.5, anchor='e', x=int(-10*sf))  # 显示 X
    else:
        clear_label.place_forget()  # 空内容时，隐藏 X

def debounce(func, delay=200):
    """装饰器函数，用于防抖"""
    def wrapper(*args, **kwargs):
        if hasattr(wrapper, 'after_id'):
            root.after_cancel(wrapper.after_id)
        wrapper.after_id = root.after(delay, lambda: func(*args, **kwargs))
    return wrapper

# 创建主窗口
try:
    # 适配系统缩放比例
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
    sf = ScaleFactor/100
    tk_sf = sf*(96/72)
    
    root = tk.Tk()
    root.withdraw()  # 先隐藏窗口
    root.tk.call('tk', 'scaling', tk_sf)  # 设置tk的缩放比例,调整控件和字体大小
    # 将 Base64 解码为二进制
    icon_data = base64.b64decode(ICON_BASE64)
    # 通过 BytesIO 读取 ICO 图标
    icon_image = Image.open(io.BytesIO(icon_data))
    icon = ImageTk.PhotoImage(icon_image)
    # 设置窗口图标
    root.iconphoto(True, icon)
    root.title("Drawing Search")
    # 绑定窗口事件
    root.bind("<Configure>", on_main_window_move)
    root.bind("<FocusIn>", debounce(on_focus_in))
    root.bind("<Visibility>", debounce(on_window_state_change))
    root.bind("<Unmap>", debounce(on_window_state_change))
    # 根据系统缩放比例调整窗口大小
    window_width = int(window_width*sf)
    window_height = int(window_height*sf)
    expand_window_width = int(expand_window_width*sf)
    root.geometry(f"{window_width}x{window_height}")  # 初始窗口大小
    root.resizable(False, False)
    # 窗口居中偏上显示
    center_window(root, window_width, window_height)
    root.deiconify() # 显示窗口
    root.protocol("WM_DELETE_WINDOW", on_root_close)  # 绑定关闭事件，清除所有线程
    
    # 第一行标签的框架
    label_frame = tk.Frame(root)
    label_frame.pack(pady=(int(15*sf), int(5*sf)), anchor="w", fill="x")

    # 标签放在第一行
    prompt_label = ttk.Label(label_frame, text="Input Part / Assembly / Project Number :", font=("Segoe UI", 9), anchor="w")
    prompt_label.pack(side=tk.LEFT, padx=(int(20*sf), 0))

    # 创建ttk控件的 Style
    style = ttk.Style()
    style.configure("Top.TCheckbutton", font=("Segoe UI", 10)) 
    style.configure("Warning.TLabel", font=("Segoe UI", 9), foreground="red")
    style.configure("All.TButton", font=("Segoe UI", 9))
    style.configure("Close.TButton", font=("Segoe UI", 7))
    style.map("Close.TButton", foreground=[("active", "red"), ("!active", "black")])  # 搜索结果列表的关闭按钮，鼠标悬停变红
    style.configure("Change.TLabel", font=("Segoe UI", 8, "underline"), foreground="blue")
    style.configure("About.TLabel", font=("Segoe UI", 12, "bold"))
    style.configure("Cache.TLabel", font=("Segoe UI", 9), foreground="lightgray")
    style.configure("Tooltip.TLabel", background="#ffffe0")
    style.configure("Clear.TLabel", background="white")
    style.configure("Thumbnail.TCheckbutton", font=("Segoe UI", 9))
    style.configure("Close.TLabel", foreground="red", background="white", font=("Segoe UI", 8))

    # 添加置顶选项
    # 创建一个 IntVar 绑定复选框的状态（0 未选中，1 选中）
    topmost_var = tk.IntVar()

    # 创建复选框，用于控制窗口置顶
    checkbox = ttk.Checkbutton(label_frame, text="📌", variable=topmost_var, style="Top.TCheckbutton", command=toggle_topmost)
    checkbox.pack(side=tk.RIGHT, padx=int(10*sf))
    Tooltip(checkbox, lambda: "Pin to top", delay=500)

    # 添加切换mini窗口的按钮
    mini_search_label = ttk.Label(label_frame, text="🍀", font=("Segoe UI", 10), cursor="hand2")
    mini_search_label.pack(side=tk.RIGHT, padx=int(5*sf))
    mini_search_label.bind("<Button-1>", lambda event: open_mini_window())
    Tooltip(mini_search_label, lambda: "Switch to mini window", delay=500)

     # 按钮宽度，输入框宽度，Label宽度和位置，无法根据缩放比例在布局内进行同比例调整，所以指定具体值
    if ScaleFactor == 100:
        btn_width = 21
        entry_width = 27
        parts_dir_width = 33
        parts_y_position = int(5*sf-3)
    elif ScaleFactor == 125:
        btn_width = 20
        entry_width = 25
        parts_dir_width = 36
        parts_y_position = int(5*sf)
    elif ScaleFactor == 150:
        btn_width = 19
        entry_width = 26
        parts_dir_width = 35
        parts_y_position = int(5*sf+4)
    elif ScaleFactor == 175:
        btn_width = 21
        entry_width = 26
        parts_dir_width = 36
        parts_y_position = int(5*sf+10)
    else:
        btn_width = 20
        entry_width = 25
        parts_dir_width = 35
        parts_y_position = int(5*sf)

    # 创建输入框框架
    entry_frame = tk.Frame(root)
    entry_frame.pack(pady=0, anchor="w", fill="x")
    entry = ttk.Entry(entry_frame, width=entry_width, font=("Segoe UI", 16))
    entry.pack(padx=int(20*sf), pady=int(5*sf), anchor="w")
    create_entry_context_menu(entry)
    entry.focus()
    entry.bind("<Return>", lambda event: search_pdf_files())
    entry.bind("<Button-1>", show_search_history)  # 点击输入框时显示历史记录

    # 清除 Entry 内容的Label，初始不显示
    clear_label = ttk.Label(entry_frame, text="✕", font=('Segoe UI', 10), foreground='red', style="Clear.TLabel", cursor="arrow")
    # 监听 Entry 内容变化来更新 Label 的颜色， 同时实时更新匹配历史
    entry.bind("<KeyRelease>", lambda event: (show_search_history(event), update_label_color(event)))
    entry.bind("<FocusIn>", update_label_color)     # 当 Entry 获取焦点时检查内容

    # 绑定点击事件：点击 Label 的 X 就清空 Entry
    clear_label.bind("<Button-1>", clear_entry)

    # 用于显示警告信息的标签
    warning_label = ttk.Label(entry_frame, text="", style="Warning.TLabel", anchor="w")
    warning_label.pack(fill="x", padx=int(20*sf))
    Tooltip(warning_label, lambda: warning_label.cget("text"), delay=500)

    # 用户点击非 Listbox 或 Entry 区域时销毁 Listbox
    root.bind("<Button-1>", hide_history)

    # 添加按钮框架
    button_frame = tk.Frame(root)
    button_frame.pack(padx=int(20*sf), pady=int(5*sf), anchor="w")

    # Search PDF 按钮
    search_btn = ttk.Button(button_frame, text="Search PDF Drawing", width=btn_width, style="All.TButton", command=search_pdf_files)
    search_btn.grid(row=0, column=0, padx=(int(5*sf), int(10*sf)), pady=int(8*sf))
    Tooltip(search_btn, lambda: "Search for PDF files matching the entered keywords", delay=500)

    # I'm Feeling Lucky 按钮
    lucky_btn = ttk.Button(button_frame, text="I'm Feeling Lucky!", style="All.TButton", width=btn_width, command=feeling_lucky)
    lucky_btn.grid(row=0, column=1, padx=(int(10*sf), int(5*sf)), pady=int(8*sf))
    Tooltip(lucky_btn, lambda: "Open the latest revision of the PDF drawing", delay=500)

    # Search 3D drawing 按钮
    search_3d_btn = ttk.Button(button_frame, text="Search 3D Drawing", style="All.TButton", width=btn_width, command=search_3d_files)
    search_3d_btn.grid(row=1, column=0, padx=(int(5*sf), int(10*sf)), pady=int(8*sf))
    Tooltip(search_3d_btn, lambda: "Search for 3D files (.iam/.ipt) matching the entered keywords", delay=500)

    # Search vault cache 按钮
    search_cache_btn = ttk.Button(button_frame, text="Search in Vault Cache", style="All.TButton", width=btn_width, command=search_vault_cache)
    search_cache_btn.grid(row=1, column=1, padx=(int(10*sf), int(5*sf)), pady=int(8*sf))
    Tooltip(search_cache_btn, lambda: "Search 3D drawings (.iam/.ipt) from local Vault cache\rSupport searching by project name", delay=500)

    # Reset 按钮
    reset_btn = ttk.Button(button_frame, text="Reset", width=btn_width, style="All.TButton", command=reset_window)
    reset_btn.grid(row=2, column=0, padx=(int(5*sf), int(10*sf)), pady=int(8*sf))
    Tooltip(reset_btn, lambda: "Reset the window to default, stop the current search\rand clear search cache", delay=500)

    # 扩展按钮
    expand_btn = ttk.Button(button_frame, text="Quick Access   ❯❯", width=btn_width, style="All.TButton", command=toggle_window_size)
    expand_btn.grid(row=2, column=1, padx=(int(10*sf), int(5*sf)), pady=int(8*sf))
    Tooltip(expand_btn, lambda: "Shortcuts to frequently used folders and files", delay=500)

    # 显示默认目录及更改功能
    directory_frame = tk.Frame(root)
    directory_frame.pack(anchor="w", padx=int(20*sf), pady=parts_y_position, fill="x")
    directory_label = ttk.Label(directory_frame, text=f"Default PARTS Directory: {default_parts_path}", font=("Segoe UI", 8), width=parts_dir_width, anchor="w")
    directory_label.pack(side=tk.LEFT)
    Tooltip(directory_label, lambda: directory_label.cget("text"), delay=500)

    # Change 按钮
    change_label = ttk.Label(directory_frame, text="Change", style="Change.TLabel", cursor="hand2")
    change_label.pack(side=tk.LEFT, padx=int(8*sf))
    Tooltip(change_label,  lambda: "Select a new PARTS directory", delay=500)
    change_label.bind("<Button-1>", lambda event: update_directory())

    # Default 按钮
    default_label = ttk.Label(directory_frame, text="Default", style="Change.TLabel", cursor="hand2")
    default_label.pack(side=tk.LEFT, padx=0)
    Tooltip(default_label,  lambda: "Reset the PARTS directory to default", delay=500)
    default_label.bind("<Button-1>", lambda event: reset_to_default_directory())

    # About 按钮
    about_frame = ttk.Frame(root)
    about_frame.pack(anchor="e", padx=0, pady=0, fill="x")
    about_label = ttk.Label(about_frame, text="ⓘ", style="About.TLabel", cursor="hand2")
    about_label.pack(side=tk.RIGHT, padx=int(10*sf), pady=(int(3*sf), int(8*sf)))
    Tooltip(about_label,  lambda: "About", delay=500)
    about_label.bind("<Button-1>", lambda event: show_about())

    # 显示缓存状态, 灰色无缓存，绿色缓存已完成，红色正在缓存
    cache_label = ttk.Label(about_frame, text="●", style="Cache.TLabel")
    cache_label.pack(side=tk.RIGHT, padx=0, pady=(int(3*sf), int(8*sf)))
    tooltip_instance = Tooltip(cache_label, get_cache_str, delay=500)

    # 添加 thumbnail_check 复选框
    thumbnail_var = tk.BooleanVar(value=True)  # 默认选中
    thumbnail_check = ttk.Checkbutton(
        about_frame, text="Thumbnails", variable=thumbnail_var, style="Thumbnail.TCheckbutton", 
        command=lambda: (on_tree_select(None), entry.focus())
    )
    Tooltip(thumbnail_check, lambda: "Show thumbnails", delay=500)

    # 运行主循环
    root.mainloop()
except Exception as e:
    print(f"An error occurred: {e}")
    messagebox.showerror("Error", f"An error occurred: {e}")