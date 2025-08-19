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
import subprocess
import locale
import urllib.request
import json
import webbrowser
from tkinter import filedialog
from tkinter import Menu, ttk
from PIL import Image, ImageTk
from logo import ICON_BASE64
from language import LANGUAGES
from partname_util import generate_partname_dat

try:
    import tkinter as tk
    from tkinter import messagebox
except ModuleNotFoundError:
    print("Error: tkinter module is not available in this environment.")
    sys.exit(1)

# 全局变量
ver = "1.5.1"  # 版本号
current_language = "en"  # 当前语言（默认英文）
previous_language = None # 切换语言前的上一个语言
search_history = []  # 用于存储最近的搜索记录，最多保存20条
changed_parts_path = None  # 用户更改的 PARTS 目录
result_frame = None  # 搜索结果的 Frame 容器
results_tree = None  # 搜索结果的 Treeview 控件
result_files_pdf = None  # 存储pdf搜索结果
result_files_3d = None  # 存储3d文件iam和ipt搜索结果
result_files_name = None  # 存储partname搜索结果
last_query = None  # 上一次的搜索结果
history_frame = None  # 用于显示搜索历史的 Frame
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
preview_win = None
last_file = None # 用于记录上一次在预览窗口中打开的文件
# 全局点击计数器，用于刷新缓存功能
refresh_cache_click_count = 0
refresh_cache_click_first_time = None
release_url = "https://api.github.com/repos/geriwolf/DrawingFinder/releases/latest"
# 快捷访问路径列表，存储按钮上显示的文字和对应路径
shortcut_paths = [
    {"label": "PARTS Folder", "path": "K:\\PARTS"},
    {"label": "Latest Missing List", "path": "V:\\Missing Lists\\Missing_Parts_List"},
    {"label": "Pneumatic Drawing Folder", "path": "K:\\Pneumatic Drawings"},
    {"label": "Projects Video Folder", "path": "G:\\Project Media"},
    {"label": "ERP Items Tool List", "path": "K:\\ERP TOOLS\\BELLATRX ERP ITEMS TOOL.xlsx"},
    {"label": "Equipment Labels Details", "path": "K:\\Equipment Labels\\Equipment Label - New"},
]

class Tooltip:
    """显示鼠标悬停提示的类"""
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

class SearchAnimation:
    """搜索时显示动画的类"""
    def __init__(self, parent, width=35, height=20, radius=3):
        self.radius = int(radius*sf)  # 圆的半径随屏幕缩放比例调整
        self.width = int(width*sf)  # 宽度随屏幕缩放比例调整
        self.height = height  # 高度固定
        self.canvas = tk.Canvas(parent, width=self.width, height=self.height, highlightthickness=0, bg="white")
        self.x = self.radius + 1  # 初始x坐标，不紧贴左边缘，留出1px的间隙
        # 创建一个圆形的 Canvas 作为动画元素
        self.oval = self.canvas.create_oval(
            self.x - self.radius, self.height // 2 - self.radius,
            self.x + self.radius, self.height // 2 + self.radius,
            fill="blue", outline=""
        )
        self.direction = 1
        self.running = False

    def start(self):
        if self.running:
            return  # 防止重复启动动画
        self.running = True
        # search_hint 已在主函数定义
        x = search_hint.winfo_x() + search_hint.winfo_width() - self.canvas.winfo_reqwidth() - 5  # canvas布局向左偏移5px，不贴边
        y = search_hint.winfo_y() + (search_hint.winfo_height() - self.canvas.winfo_reqheight()) // 2  # 垂直居中
        self.canvas.place(x=x, y=y)
        self.animate()

    def animate(self):
        if not self.running:
            return

        # 获取当前坐标
        x1, _, x2, _ = self.canvas.coords(self.oval)
        step = int(2*sf)  # 每次移动的步长，随屏幕缩放比例调整
        # 计算圆心的位置
        center_x = (x1 + x2) / 2
        # 计算下一帧中心位置
        next_center_x = center_x + self.direction * step
        # 如果下一帧将超出边界，就改方向
        if next_center_x - self.radius < 0 or next_center_x + self.radius > self.width:
            self.direction *= -1

        self.canvas.move(self.oval, self.direction * step, 0)
        self.canvas.after(30, self.animate)

    def stop(self):
        if not self.running:
            return
        self.running = False
        self.canvas.place_forget()

def show_warning_message(message, color="blue"):
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
    entry_focus()  # 焦点重新回到输入框
    path = shortcut_paths[index]["path"]

    if os.path.exists(path):
        # 如果 label 包含 "Missing Lists" 字段，使用get_latest_file函数查找最新的文件
        if "Missing List" in shortcut_paths[index]["label"]:
            prefix_name = "Master_Missing_List"
            latest_file = get_latest_file(prefix_name, path)
            if latest_file:
                path = latest_file
            else:
                messagebox.showwarning(LANGUAGES[current_language]['warning'], LANGUAGES[current_language]['no_missing_list'])
                return

        # 如果 label 包含 "Equipment Labels Details" 字段，使用get_latest_file函数查找最新的文件
        if "Equipment Labels Details" in shortcut_paths[index]["label"]:
            prefix = "Equipment New Labels Details"
            latest_file = get_latest_file(prefix, path)
            if latest_file:
                path = latest_file
            else:
                messagebox.showwarning(LANGUAGES[current_language]['warning'], LANGUAGES[current_language]['no_labels_file'])
                return
        
        try:
            if os.path.isdir(path):
                # 如果是目录，打开目录
                os.startfile(path)
            else:
                # 如果是文件，通过 open_file 打开
                open_file(file_path=path)
        except Exception as e:
            messagebox.showerror(LANGUAGES[current_language]['error'], f"{LANGUAGES[current_language]['failed_to_open_shortcut']}: {e}")
    else:
        messagebox.showerror(LANGUAGES[current_language]['error'], f"{LANGUAGES[current_language]['shortcut_not_exist']}: {path}")

def update_directory():
    """更新搜索目录"""
    global changed_parts_path, last_query
    new_dir = filedialog.askdirectory(initialdir=default_parts_path, title="Darwing Finder")
    if new_dir:
        new_dir = new_dir.replace('/', '\\')  # 将路径中的斜杠替换为反斜杠
        directory_label.config(text=f"{LANGUAGES[current_language]['parts_dir']} {new_dir}")
        changed_parts_path = new_dir
        last_query = None
        
def reset_to_default_directory():
    """将搜索路径重置为默认路径"""
    global changed_parts_path, last_query
    directory_label.config(text=f"{LANGUAGES[current_language]['default_parts_dir']} {default_parts_path}")
    changed_parts_path = None
    last_query = None

def open_file(event=None, file_path=None):
    """用系统默认程序打开选中的文件"""
    if not file_path:
        selected_item = results_tree.selection()  # 获取选中的项
        if not selected_item or not results_tree.exists(selected_item[0]):
            return
        # 判断results_tree的列数，获取默认文件路径
        if len(results_tree["columns"]) == 4:
            file_path = results_tree.item(selected_item[0], 'values')[2]  # 获取文件路径
        elif len(results_tree["columns"]) == 6:
            values = results_tree.item(selected_item[0], 'values')
            file_path = values[2] or values[3] or values[4]
            if not file_path:
                show_warning_message(LANGUAGES[current_language]['no_drawing_open'], "red")
                return

    if file_path and os.path.exists(file_path):
        try:
            os.startfile(file_path)
        except Exception as e:
            show_warning_message(f"{LANGUAGES[current_language]['failed_to_open_file']}: {e}", "red")
    else:
        show_warning_message(f"{LANGUAGES[current_language]['file_not_found']}: {file_path}", "red")

def save_search_history(query):
    """保存搜索记录并限制最多保存20条"""
    global search_history
    if query and query not in search_history:
        search_history.append(query)
        if len(search_history) > 20:
            search_history.pop(0)

def show_search_history(event):
    """在输入框下显示搜索历史"""
    global history_listbox, history_frame

    # 如果存在旧的列表框，先销毁它
    if history_listbox:
        history_listbox.destroy()
    if history_frame:
        history_frame.destroy()

    query = entry.get().lower()
    if search_hint.cget("text") != LANGUAGES[current_language]['enter_search'] and not search_anim.running:
        search_hint.config(text=LANGUAGES[current_language]['enter_search'])
    if not query:
        matching_history = search_history
    else:
        matching_history = [h for h in search_history if query in h.lower()]

    if not matching_history:
        return  # 如果没有匹配的历史记录，则不显示列表框
    
    if len(matching_history) == 1:
        if matching_history[0].lower() == query.lower():
            return

    # 创建一个 Frame 包含 Listbox 和 Scrollbar
    history_frame = tk.Frame(
        root,
        bd=0,
        relief="solid",
        highlightthickness=1,
        highlightbackground="SystemHighlight",  # 蓝色边框（未聚焦）
        highlightcolor="SystemHighlight",       # 蓝色边框（聚焦时）
        bg="white"
    )

    border_thickness = 1  # 边框厚度

    def get_listbox_height_width(root, rows, columns):
        # 用临时 Listbox 获取高度和宽度的像素值
        test = tk.Listbox(root, height=rows, width=columns)
        test.place(x=-1000, y=-1000)
        root.update_idletasks()
        height = test.winfo_height()
        width = test.winfo_width()
        test.destroy()
        return height, width

    row_number = min(len(matching_history), 5)  # 最多显示5条记录
    height_px, width_px = get_listbox_height_width(root, row_number, 12)  # 获取listbox高度和宽度
    # 创建 Listbox
    history_listbox = tk.Listbox(history_frame)
    history_listbox.place(x=0, y=0, width=width_px, height=height_px)

    # 如果搜索历史记录超过5条，添加滚动条
    if len(matching_history) > 5:
        # 创建垂直滚动条
        scrollbar = tk.Scrollbar(history_frame, orient=tk.VERTICAL, command=history_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        history_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar_width = 17  # 滚动条宽度
    else:
        scrollbar_width = 0  # 没有滚动条时宽度为0
       
    for item in matching_history:
        history_listbox.insert(0, item) # 最新的搜索记录显示在列表最上面

    # 获取输入框的绝对位置
    # 因为entry放置在entry_frame中，所以需要计算相对位置，用entry获取x坐标，用entry_frame获取y坐标
    # 放置整体 Frame
    x = entry.winfo_x()
    # entry在entry_frame中往下偏移了5px，为了使listbox的蓝色边框与entry的边框重叠，这里使用了4px
    y = entry_frame.winfo_y() + int(4*sf) + entry.winfo_height()
    history_frame.place(x=x, y=y, width=width_px + int(scrollbar_width*sf) + border_thickness * 2, height=height_px + border_thickness * 2)

    # 绑定点击事件
    history_listbox.bind("<ButtonRelease-1>", lambda event: select_history(event, history_listbox))

def hide_history(event):
    """点击窗口其他部分时隐藏搜索历史"""
    global history_listbox, history_frame

    if history_listbox:
        clicked_widget = event.widget
        # 如果点击的不是 Entry 且不是 Listbox 或 Scrollbar，就隐藏
        if clicked_widget not in (entry, history_listbox) and not str(clicked_widget).startswith(str(history_frame)):
            history_listbox.destroy()
            history_listbox = None
            if 'history_frame' in globals() and history_frame:
                history_frame.destroy()
                history_frame = None


def select_history(event, listbox):
    """当选择历史记录时，填充到输入框并销毁列表框"""
    if not listbox.curselection():
        return  # 如果没有选中任何项，直接返回

    selection = listbox.get(listbox.curselection())
    entry.delete(0, tk.END)
    entry.insert(0, selection)
    entry.focus_set()  # 重新聚焦到输入框

    # 销毁列表框
    global history_listbox, history_frame
    history_listbox.destroy()
    history_listbox = None
    if 'history_frame' in globals() and history_frame:
        history_frame.destroy()
        history_frame = None

def disable_search_button():
    """禁用所有搜索按钮"""
    search_btn.config(state=tk.DISABLED)
    lucky_btn.config(state=tk.DISABLED)
    search_3d_btn.config(state=tk.DISABLED)
    search_cache_btn.config(state=tk.DISABLED)
    search_partname_btn.config(state=tk.DISABLED)
    search_hint.config(text="")  # 清空文字
    root.update_idletasks()  # 更新界面
    search_anim.start()  # 启动动画

def enable_search_button():
    """启用所有搜索按钮"""
    search_btn.config(state=tk.NORMAL)
    lucky_btn.config(state=tk.NORMAL)
    search_3d_btn.config(state=tk.NORMAL)
    search_cache_btn.config(state=tk.NORMAL)
    search_partname_btn.config(state=tk.NORMAL)
    search_anim.stop()  # 停止动画
    if entry.get() and result_frame is None:
        # 如果输入框有内容且没有搜索结果，显示回车搜索的提示信息
        search_hint.config(text=LANGUAGES[current_language]['enter_search'])

def build_directory_cache_thread(search_directory):
    """
    遍历指定目录，构建目录缓存。
    每个目录缓存包括：
        - files_info: 当前目录所有文件的列表，每个元素为 (文件名, 创建时间, 文件路径)
        - idw_index: 前9位 part_number -> idw 文件路径的字典
    """
    global directory_cache, active_threads
    thread = threading.current_thread()

    # 检查线程是否已经在运行
    if any(t.name == f"cache_thread_{search_directory}" for t in active_threads):
        return
    
    active_threads.add(thread)
    thread.name = f"cache_thread_{search_directory}"
    
    # 缓存开始，更改cache_label以及refresh_cache_label的颜色
    root.after(0, lambda: cache_label.config(foreground="red"))
    root.after(0, lambda: refresh_cache_label.config(foreground="lightgray"))

    try:
        files_info = []
        for root_dir, _, files in os.walk(search_directory):
            if stop_event.is_set():  # 检查是否需要终止
                return

            sort_files = sorted(files)  # 按文件名排序
            for file in sort_files:
                if stop_event.is_set():  # 检查是否需要终止
                    return
                if file.endswith((".pdf", ".iam", ".ipt", ".idw")):
                    file_path = os.path.join(root_dir, file)
                    try:
                        create_time = os.path.getctime(file_path)
                    except Exception as e:
                        create_time = 0
                    files_info.append((file, create_time, file_path))

        if stop_event.is_set():  # 检查停止标志并返回
            return

        # 构建 IDW 索引
        idw_index = {}
        for f, _, path in files_info:
            if f.endswith(".idw"):
                part_number = f[:9]
                idw_index[part_number] = path

        # 使用线程锁保护directory_cache
        with cache_lock:
            # 如果缓存数量超过 cache_max_size 限制，则删除最老的未使用项（LRU）
            if len(directory_cache) >= cache_max_size:
                directory_cache.popitem(last=False)  # 移除最老的未使用项

            # 存入缓存，value 为字典，包含 files_info 和 idw_index
            directory_cache[search_directory] = {
                "files_info": files_info,
                "idw_index": idw_index
            }
            # 将缓存移动到末尾（表示最近使用）
            directory_cache.move_to_end(search_directory)

    except Exception as e:
        root.after(0, lambda e=e: messagebox.showerror(
            LANGUAGES[current_language]['error'],
            f"{LANGUAGES[current_language]['error_cache']}: {e}"
        ))

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
            return f"{LANGUAGES[current_language]['no_cache']}"
        else:
            # 有已缓存的目录
            return f"{LANGUAGES[current_language]['cache_completed']} [{', '.join(cached_dir)}]"
    else:
        # 有正在缓存的目录
        if not cached_dir:
            # 没有已缓存的目录
            return f"{LANGUAGES[current_language]['cache_in_progress']} [{', '.join(caching_list)}]"
        else:
            # 有已缓存的目录
            return f"{LANGUAGES[current_language]['cache_in_progress']} [{', '.join(caching_list)}]\r{LANGUAGES[current_language]['cache_completed']} [{', '.join(cached_dir)}]"

def show_cache_status():
    # 获取cache的状态并设置cache_label和refresh_cache_label颜色
    cache_pattern = re.compile(r"cache_thread_.*?\\([^\\,]+)(?=\s|,|$)")
    caching_list = [match.group(1) for item in active_threads if (match := cache_pattern.search(str(item)))]

    # 重置按钮不清空缓存，所以这段注释掉
    # 如果点击了重置按钮，直接改为灰色
    #if stop_event.is_set():
    #    root.after(0, lambda: cache_label.config(foreground="lightgray"))  # 设置cache_label为灰色
    #    root.after(0, lambda: refresh_cache_label.config(foreground="#F0F0F0"))  # 隐藏refresh_cache_label
    #    return
    
    if not caching_list:
        if not directory_cache:
            # 没有正在缓存的目录，也没有已缓存的目录
            root.after(0, lambda: cache_label.config(foreground="lightgray"))
            root.after(0, lambda: refresh_cache_label.config(foreground="#F0F0F0"))  # 隐藏refresh_cache_label
        else:
            # 没有正在缓存的目录，但有已缓存的目录
            root.after(0, lambda: cache_label.config(foreground="lime"))  # 设置cache_label为绿色
            root.after(0, lambda: refresh_cache_label.config(foreground="black"))  # 显示refresh_cache_label
    else:
        # 有正在缓存的目录
        root.after(0, lambda: cache_label.config(foreground="red"))  # 设置cache_label为红色
        root.after(0, lambda: refresh_cache_label.config(foreground="lightgray")) # 设置refresh_cache_label为灰色

def get_cached_directory(search_directory):
    """
    获取缓存的目录信息，如果存在则返回，否则返回 None。
    并将访问的缓存项移动到末尾（表示最近使用）
    """
    if search_directory in directory_cache:
        directory_cache.move_to_end(search_directory)
        return directory_cache[search_directory]
    return None

def refresh_cache():
    """对已缓存目录进行刷新缓存操作"""
    global directory_cache, last_query
    # 刷新前先清空目录缓存
    cached_dirs = list(directory_cache.keys())
    directory_cache.clear()
    last_query = None
    # 对于已经建立缓存的所有目录，重新创建缓存
    for search_directory in cached_dirs:
        # 重新启动缓存线程
        thread = threading.Thread(target=build_directory_cache_thread, args=(search_directory,), daemon=True)
        thread.start()

def on_refresh_cache_click(event):
    """通过点击refresh_cache_label的次数5秒内达到5次来刷新缓存"""
    global refresh_cache_click_count, refresh_cache_click_first_time
    # 如果没有缓存目录，直接返回
    if not directory_cache:
        return
    # 检查是否有缓存线程在运行，如果有则不刷新
    if any(t.name.startswith("cache_thread") for t in active_threads):
        return

    now = datetime.datetime.now()

    if refresh_cache_click_first_time is None:
        # 如果第一次点击，记录当前时间，计数器置0
        refresh_cache_click_first_time = now
        refresh_cache_click_count = 0
        # 如果显示的是"Cache is refreshing"的提示，但是没有任何缓存线程在运行，说明缓存已经刷新完成，则隐藏提示并返回
        if warning_label.cget("text") == LANGUAGES[current_language]['cache_refreshing'] and not any(t.name.startswith("cache_thread") for t in active_threads):
            hide_warning_message()
            return
    elif (now - refresh_cache_click_first_time).total_seconds() > 5:
        # 如果第一次点击距离当前点击超过5秒，重新计时，计数器置0
        refresh_cache_click_first_time = now
        refresh_cache_click_count = 0
        if warning_label.cget("text").startswith(LANGUAGES[current_language]['continue_click_1']):
            # 如果提示信息是"Continue clicking"，说明未能在5秒内点击5次，给出提示超时，并返回
            show_warning_message(LANGUAGES[current_language]['timeout_click'], "red")
            return
        elif warning_label.cget("text") == LANGUAGES[current_language]['timeout_click']:
            # 如果提示信息是"Timeout click"，说明距离上一次超时点击又超时5s，则隐藏提示，并返回
            hide_warning_message()
            return
    elif ((now - refresh_cache_click_first_time).total_seconds()) <= 5:
        # 如果第一次点击距离当前点击在5秒内，继续计数
        if warning_label.cget("text") == LANGUAGES[current_language]['timeout_click']:
            # 如果提示信息是"Timeout click"，说明上一次点击已经超时，但当前点击距离前一次在5s内，则隐藏提示，并继续计数
            hide_warning_message()
    
    refresh_cache_click_count += 1

    # 如果点击次数达到2次，给出提示继续点击可刷新缓存
    if 2 <= refresh_cache_click_count < 5:
        remaining = 5 - refresh_cache_click_count
        show_warning_message(f"{LANGUAGES[current_language]['continue_click_1']} {remaining} {LANGUAGES[current_language]['continue_click_2']}", "blue")

    # 如果点击次数达到5次，刷新缓存
    if refresh_cache_click_count >= 5:
        refresh_cache_click_count = 0
        refresh_cache_click_first_time = None
        refresh_cache()
        show_warning_message(LANGUAGES[current_language]['cache_refreshing'], "blue")

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
        show_warning_message(LANGUAGES[current_language]['enter_number'], "red")
        enable_search_button() # 启用搜索按钮
        last_query = None
        return

    # 检查是否包含非法字符
    if any(char in query for char in "*.?+^$[]{}|\\()"):
        show_warning_message(LANGUAGES[current_language]['invalid_characters'], "red")
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
        show_warning_message(f"{LANGUAGES[current_language]['path_not_exist']} {search_directory}", "red")
        show_result_list(None) # 目录不存在就清空已有搜索结果
        enable_search_button() # 启用搜索按钮
        last_query = None
        return

    # 执行搜索
    show_warning_message(LANGUAGES[current_language]['searching'], "red")
    query = query.lower()
    # 对STK的project number进行特殊处理
    if query.startswith("stk") and len(query) > 3:
        if query[3] == '-' or query[3] == ' ':
            query = query[:3] + '.*' + query[4:]
        else:
            query = query[:3] + '.*' + query[3:]

    stop_event.clear()  # 确保上一次的停止信号被清除
    search_thread = threading.Thread(target=search_files_thread, args=(query, search_directory, search_type), daemon=True)
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
        cache_entry = get_cached_directory(search_directory)
        if cache_entry is None:
            # 如果缓存中没有该目录的记录，则在后台建立缓存
            # 检查是否已经有当前搜索目录的缓存线程在运行
            if not any(t.name == f"cache_thread_{search_directory}" for t in active_threads):
                cache_thread = threading.Thread(target=build_directory_cache_thread, args=(search_directory,), daemon=True)
                cache_thread.name = f"cache_thread_{search_directory}"
                cache_thread.start()

            # 定义变量用于比对上一次循环中检查的idw文件路径结果
            last_idw_file_path = ""
            last_idw_exist = ""
            def get_idw_file_path(file):
                """获取对应的idw文件路径，如果存在则返回路径，否则返回空字符串"""
                nonlocal last_idw_file_path, last_idw_exist  # 声明使用外层变量
                p_number = file[:9]  # 获取文件名的前9个字符作为part number
                # idw文件名和路径
                idw_file = f"{p_number}.idw"
                idw_file_path = os.path.join(root_dir, idw_file)

                if idw_file_path == last_idw_file_path:
                    # 和上一次一样，直接复用
                    idw_exist = last_idw_exist
                else:
                    # 不一样，重新判断并更新变量
                    idw_exist = idw_file_path if os.path.isfile(idw_file_path) else ""  # 如果idw文件存在，返回路径，否则返回空字符串
                    last_idw_file_path = idw_file_path
                    last_idw_exist = idw_exist
                return idw_exist
            i = 50
            # 直接遍历目录下的文件，不使用缓存
            for root_dir, _, files in os.walk(search_directory):
                if stop_event.is_set():  # 检查是否需要终止
                    return
                # 对文件名进行排序，因为os.walk返回的文件列表是无序的
                sorted_files = sorted(files)
                for file in sorted_files:
                    if stop_event.is_set():  # 检查是否需要终止
                        return
                    # 每遍历50个文件，显示一次文件名，体现搜索过程
                    if i == 50:
                        root.after(0, lambda: show_warning_message(f"{LANGUAGES[current_language]['searching']}  {file}", "red"))
                        i = 0
                    i += 1
                    # 根据不同文件类型，存储到不同list里
                    if file.endswith(".pdf") and match_func(file):
                        file_path = os.path.join(root_dir, file)
                        create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                        idw_exist = get_idw_file_path(file)  # 获取对应的idw文件路径
                        result_files_pdf.append((file, create_time, file_path, idw_exist))  # (文件名，创建时间，文件路径，idw文件路径)
                    elif (file.endswith(".iam") or file.endswith(".ipt")) and match_func(file):
                        file_path = os.path.join(root_dir, file)
                        create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                        idw_exist = get_idw_file_path(file)  # 获取对应的idw文件路径
                        result_files_3d.append((file, create_time, file_path, idw_exist))  # (文件名，创建时间，文件路径，idw文件路径)

        else:
            # 使用缓存中的文件信息
            all_files = cache_entry["files_info"]  # 从缓存中获取所有文件信息
            idw_index = cache_entry["idw_index"]  # 从缓存中获取idw索引字典
            i = 50
            for file_info in all_files:
                if stop_event.is_set():  # 检查是否需要终止
                    return
                file_name = file_info[0]
                # 每遍历50个文件，显示一次文件名，体现搜索过程
                if i == 50:
                    root.after(0, lambda: show_warning_message(f"{LANGUAGES[current_language]['searching']}  {file_name}", "red"))
                    i = 0
                i += 1
                # 根据不同文件类型，存储到不同list里
                if file_name.endswith(".pdf") and match_func(file_name):
                    # 格式化创建时间
                    create_time = datetime.datetime.fromtimestamp(file_info[1]).strftime("%Y-%m-%d %H:%M:%S")
                    # 从缓存的索引字典中检查有没有对应的idw文件
                    p_number = file_name[:9]  # 获取文件名的前9个字符作为part number
                    # 检查idw索引字典，没有则返回空字符串
                    idw_exist = idw_index.get(p_number, "")
                    result_files_pdf.append((file_name, create_time, file_info[2], idw_exist))  # (文件名，创建时间，文件路径，idw文件路径)
                elif (file_name.endswith(".iam") or file_name.endswith(".ipt")) and match_func(file_name):
                    # 格式化创建时间
                    create_time = datetime.datetime.fromtimestamp(file_info[1]).strftime("%Y-%m-%d %H:%M:%S")
                    # 从缓存的索引字典中检查有没有对应的idw文件
                    p_number = file_name[:9]  # 获取文件名的前9个字符作为part number
                    # 检查idw索引字典，没有则返回空字符串
                    idw_exist = idw_index.get(p_number, "")
                    result_files_3d.append((file_name, create_time, file_info[2], idw_exist))  # (文件名，创建时间，文件路径，idw文件路径)

        if stop_event.is_set():  # 检查停止标志并返回
            return
        # pdf搜索结果排序（根据搜索的关键字不同，进行不同排序）
        if len(query) > 2 and query[2].isdigit():
            # 如果是project number，按文件名正序排列，方便根据文件名查找
            result_files_pdf.sort(key=lambda x: x[0])
        else:
            # 如果是part number或者assembly number，按创建时间倒序排列，方便查找最新的revision
            result_files_pdf.sort(key=lambda x: x[1], reverse=True)
        # 3d文件搜索结果排序（按文件名排序）
        result_files_3d.sort(key=lambda x: x[0])

        root.after(0, hide_warning_message)  # 使用主线程清除警告信息

        if search_type == "pdf" or search_type == "lucky":
            if not result_files_pdf:
                # 如果没有搜索到匹配的文件，显示警告信息
                root.after(0, lambda: show_warning_message(LANGUAGES[current_language]['no_matching_pdf'], "red"))
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
                root.after(0, lambda: show_warning_message(LANGUAGES[current_language]['no_matching_3d'], "red"))
            # 显示搜索结果
            root.after(0, lambda: show_result_list(result_files_3d))
        
        root.after(0, lambda: enable_search_button())  # 启用搜索按钮

    except Exception as e:
        root.after(0, lambda e=e: messagebox.showerror(LANGUAGES[current_language]['error'], f"{LANGUAGES[current_language]['error_search']}: {e}"))

    finally:
        active_threads.discard(thread)  # 线程结束后移除

def fill_entry_from_clipboard(widget=None):
    """如果输入框为空，尝试从剪贴板读取内容并填入，检查是否是part number格式"""
    target_widget = widget if widget else entry
    if not target_widget.get().strip():
        try:
            clipboard_content = root.clipboard_get().strip()
            # 检查剪贴板内容是否符合part number格式（两位数字开头），并且长度不超过20个字符
            if re.match(r'^\d{2}', clipboard_content) and len(clipboard_content) <= 20:
                target_widget.delete(0, tk.END)
                target_widget.insert(0, clipboard_content)
        except tk.TclError:
            pass  # 剪贴板为空或不可访问

def search_pdf_files():
    """执行pdf搜索"""
    global last_query, history_listbox, history_frame

    # 当敲回车执行搜索时，隐藏搜索历史列表
    if history_listbox:
        history_listbox.destroy()
        history_listbox = None
    if history_frame:
        history_frame.destroy()
        history_frame = None

    entry_focus()  # 保持焦点在输入框

    fill_entry_from_clipboard(widget=entry)  # 从剪贴板读取内容并填入
    
    query = entry.get().strip() # 去除首尾空格
    if query == last_query:
        # 如果搜索关键字跟上一次一样，直接调用上一次的搜索结果
        if not result_files_pdf:
            # 如果结果是空，显示警告信息，移除之前显示的搜索结果
            show_warning_message(LANGUAGES[current_language]['no_matching_pdf'], "red")
            close_result_list()
        else:
            disable_search_button()
            show_warning_message(f"{LANGUAGES[current_language]['searching']}", "red")
            # 后台执行show_result_list
            root.after(10, lambda: (show_result_list(result_files_pdf, search_type="pdf"), hide_warning_message(), enable_search_button()))
    else:
        # 如果搜索关键字跟上一次不一样，重新搜索
        last_query = query
        search_files(query, search_type="pdf")

def feeling_lucky():
    """执行feeling lucky搜索"""
    global last_query
    entry_focus()  # 保持焦点在输入框

    fill_entry_from_clipboard(widget=entry)  # 从剪贴板读取内容并填入
    
    query = entry.get().strip() # 去除首尾空格
    if query == last_query:
        # 如果搜索关键字跟上一次一样，直接调用上一次的搜索结果
        if not result_files_pdf:
            # 如果结果是空，显示警告信息，移除之前显示的搜索结果
            show_warning_message(LANGUAGES[current_language]['no_matching_pdf'], "red")
            close_result_list()
        else:
            disable_search_button()
            show_warning_message(f"{LANGUAGES[current_language]['searching']}", "red")
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
    entry_focus()  # 保持焦点在输入框

    fill_entry_from_clipboard(widget=entry)  # 从剪贴板读取内容并填入

    query = entry.get().strip() # 去除首尾空格
    if query == last_query:
        # 如果搜索关键字跟上一次一样，直接调用上一次的搜索结果
        if not result_files_3d:
            # 如果结果是空，显示警告信息，移除之前显示的搜索结果
            show_warning_message(LANGUAGES[current_language]['no_matching_3d'], "red")
            close_result_list()
        else:
            disable_search_button()
            show_warning_message(f"{LANGUAGES[current_language]['searching']}", "red")
            # 后台执行show_result_list
            root.after(10, lambda: (show_result_list(result_files_3d), hide_warning_message(), enable_search_button()))
    else:
        # 如果搜索关键字跟上一次不一样，重新搜索
        last_query = query
        search_files(query, search_type="3d")

def search_vault_cache():
    """搜索Vault缓存目录下的 3D 文件(ipt或者iam)"""
    global last_query
    entry_focus()  # 保持焦点在输入框

    fill_entry_from_clipboard(widget=entry)  # 从剪贴板读取内容并填入

    disable_search_button() # 禁用搜索按钮
    hide_warning_message()  # 清除警告信息
    query = entry.get().strip() # 去除首尾空格
    if not query:
        show_warning_message(LANGUAGES[current_language]['enter_number'], "red")
        enable_search_button() # 启用搜索按钮
        return
    
    # 检查是否包含非法字符
    if any(char in query for char in "*.?+^$[]{}|\\()"):
        show_warning_message(LANGUAGES[current_language]['invalid_characters'], "red")
        enable_search_button() # 启用搜索按钮
        return

    save_search_history(query)  # 保存搜索记录

    matching_directories = []
    search_directory = None
    real_query = None
    last_query = None

    if not os.path.exists(vault_cache):
        # 如果Vault缓存目录不存在，提示用户使用Vault
        show_warning_message(LANGUAGES[current_language]['vault_cache_not_found'], "red")
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
            show_warning_message(LANGUAGES[current_language]['no_matching_3d_cache'], "red")
            show_result_list(None) # 目录不存在就清空已有搜索结果
            enable_search_button() # 启用搜索按钮
            return
    else:
        # 任何其他字符串，都当作是project name去匹配，去PARTS/S路径下查找匹配的目录
        if not os.path.exists(os.path.join(vault_cache, "S")):
            show_warning_message(LANGUAGES[current_language]['no_matching_3d_cache'], "red")
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
                show_warning_message(LANGUAGES[current_language]['cancelled'], "red")
                enable_search_button() # 启用搜索按钮
                return
        else:
            # 如果用户输入的关键字匹配不到任何project，直接当作子目录去PARTS下搜索，如PARTS/XX
            prefix = query[:2]
            search_directory = os.path.join(vault_cache, prefix)
            if not os.path.exists(search_directory):
                show_warning_message(LANGUAGES[current_language]['no_matching_3d_cache'], "red")
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
                show_warning_message(LANGUAGES[current_language]['cancelled'], "red")
                enable_search_button() # 启用搜索按钮
                return
        else:
            show_warning_message(LANGUAGES[current_language]['no_matching_3d_cache'], "red")
            show_result_list(None) # 目录不存在就清空已有搜索结果
            enable_search_button() # 启用搜索按钮
            return

    # 执行搜索
    show_warning_message(LANGUAGES[current_language]['searching'], "red")
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
    search_thread = threading.Thread(target=search_vault_cache_thread, args=(query, search_directory,), daemon=True)
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
                    root.after(0, lambda: show_warning_message(f"{LANGUAGES[current_language]['searching']}  {file}", "red"))
                    i = 0
                i += 1
                if (file.endswith(".iam") or file.endswith(".ipt")) and match_func(file):
                    file_path = os.path.join(root_dir, file)
                    create_time = datetime.datetime.fromtimestamp(os.path.getctime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
                    placeholder = "" # 对于Vault缓存搜索，没有对应的idw文件，最后一段用空字符串占位，便于在treeview中处理
                    result_files.append((file, create_time, file_path, placeholder))

        if stop_event.is_set():  # 检查停止标志并返回
            return
        root.after(0, hide_warning_message)  # 使用主线程清除警告信息

        if not result_files:
            root.after(0, lambda: show_warning_message(LANGUAGES[current_language]['no_matching_3d_cache'], "red"))
        else:
            root.after(0, lambda: show_warning_message(LANGUAGES[current_language]['not_latest'], "blue"))
            result_files.sort(key=lambda x: x[0])  # 搜索结果按文件名排序

        root.after(0, lambda: show_result_list(result_files))
        root.after(0, lambda: enable_search_button())  # 启用搜索按钮
    except Exception as e:
        root.after(0, lambda e=e: messagebox.showerror(LANGUAGES[current_language]['error'], f"{LANGUAGES[current_language]['error_search']}: {e}"))

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
    choice_win.title(LANGUAGES[current_language]['select_project'])
    choice_win_width = int(280*sf)
    choice_win_height = int(200*sf)
    choice_win.geometry(f"{choice_win_width}x{choice_win_height}")
    choice_win.resizable(False, False)

    # 窗口位置，跟随主窗口初始大小居中显示，方便鼠标选取
    root.update_idletasks()  # 刷新主窗口状态
    choice_win.update_idletasks()
    position_right = int(root.winfo_x() + window_width/2 - choice_win_width/2)
    position_down = int(root.winfo_y() + window_height/2 - choice_win_height/2)
    choice_win.geometry(f"+{position_right}+{position_down}")
    choice_win.deiconify() # 显示窗口

    # 主容器
    frame = tk.Frame(choice_win)
    frame.pack(fill=tk.BOTH, expand=True, padx=(int(15*sf), int(5*sf)), pady=int(10*sf))

    # 提示文本
    label = ttk.Label(frame, text=LANGUAGES[current_language]['multiple_projects'], font=("Segoe UI", 9), anchor="w")
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

    cancel_btn = ttk.Button(btn_frame, text=LANGUAGES[current_language]['cancel'], width=8, style="All.TButton", command=choice_win.destroy)
    cancel_btn.pack(side=tk.RIGHT, padx=0)

    select_btn = ttk.Button(btn_frame, text=LANGUAGES[current_language]['select_project'], width=15, style="All.TButton", command=on_select)
    select_btn.pack(side=tk.RIGHT, padx=int(15*sf))
    select_btn.config(state=tk.DISABLED) # 默认禁用选择按钮

    # 等待窗口关闭
    choice_win.wait_window()
    return selected_dir[0]

def search_partname():
    """根据关键字搜索part name， part name的数据存在于JSON文件partname.dat中"""
    global last_query
    # 获取程序的路径
    if getattr(sys, 'frozen', False):
        # 打包后的 .exe 环境
        base_dir = os.path.dirname(sys.executable)
    else:
        # 正常调试环境（解释器运行）
        base_dir = os.path.dirname(os.path.abspath(__file__))  # 获取当前脚本所在目录
    
    # 生成partname.dat的完整路径
    partname_dat = os.path.join(base_dir, "partname.dat")

    # 如果partname.dat不存在，后台生成
    if not os.path.exists(partname_dat):
        gen_partname(partname_dat)
    else:
        # 进行part name搜索处理        
        entry_focus()  # 保持焦点在输入框

        fill_entry_from_clipboard(widget=entry)  # 从剪贴板读取内容并填入

        disable_search_button() # 禁用搜索按钮
        hide_warning_message()  # 清除警告信息
        query = entry.get().strip().lower() # 去除首尾空格
        if not query:
            show_warning_message(LANGUAGES[current_language]['enter_number'], "red")
            enable_search_button() # 启用搜索按钮
            return
        
        # 检查是否包含非法字符
        if any(char in query for char in "*.?+^$[]{}|\\()"):
            show_warning_message(LANGUAGES[current_language]['invalid_characters'], "red")
            enable_search_button() # 启用搜索按钮
            return

        save_search_history(query)  # 保存搜索记录
        last_query = None

        # 执行搜索
        show_warning_message(LANGUAGES[current_language]['searching'], "red")
        query = query.lower()

        stop_event.clear()  # 确保上一次的停止信号被清除
        search_thread = threading.Thread(target=search_partname_thread, args=(query, partname_dat), daemon=True)
        search_thread.start()

def search_partname_thread(query, partname_dat):
    """使用多线程读取part name数据文件并搜索"""
    global active_threads, result_files_name

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
        
        if not hasattr(search_partname, "data"):
            with open(partname_dat, "r", encoding="utf-8") as f:
                search_partname.data = json.load(f)

        
        result_files_name = []
        # 遍历partname.dat中的数据，查找匹配的part name
        for item in search_partname.data.values():
            if stop_event.is_set():  # 检查停止标志并返回
                return
            id_ = item.get("id", "")
            description = item.get("description", "")
            if query in id_.lower() or query in description.lower():
                pdf = item.get("pdf", "")
                ipt = item.get("ipt", "")
                iam = item.get("iam", "")

                # 根据drawing类型生成对应字符串用于在结果中显示
                drawing = []
                if pdf: drawing.append("PDF")
                if ipt: drawing.append("IPT")
                if iam: drawing.append("IAM")
                drawing_type = ", ".join(drawing)

                result_files_name.append([
                    id_,
                    description,
                    pdf,
                    ipt,
                    iam,
                    drawing_type
                ])

        if stop_event.is_set():  # 检查停止标志并返回
            return
        
        root.after(0, hide_warning_message)  # 使用主线程清除警告信息

        if not result_files_name:
            root.after(0, lambda: show_warning_message(LANGUAGES[current_language]['no_matching_name'], "red"))
        else:
            hide_warning_message()
        
        # 显示搜索结果
        root.after(0, lambda: show_result_list(result_files_name, search_type="name"))
        root.after(0, lambda: enable_search_button())  # 启用搜索按钮

    except Exception as e:
        root.after(0, lambda e=e: messagebox.showerror(LANGUAGES[current_language]['error'], f"{LANGUAGES[current_language]['error_search']}: {e}"))

    finally:
        active_threads.discard(thread)  # 线程结束后移除

def gen_partname(partname_dat):
    """后台生成 partname.dat"""
    thread_name = "gen_partname_thread"

    def on_done():
        hide_warning_message()  # 隐藏生成中提示
        messagebox.showinfo("Part Name", f"{LANGUAGES[current_language]['partname_generated']}\n\n"
                            f"{LANGUAGES[current_language]['file_path']}\n{partname_dat}\n\n"
                            f"{LANGUAGES[current_language]['delete_data']}")

    def run():
        if changed_parts_path:
            # 如果用户修改了parts路径，使用新的路径
            parts_path = changed_parts_path
        else:
            # 否则使用默认的parts路径
            parts_path = default_parts_path
        try:
            generate_partname_dat(partname_dat, parts_path, callback=on_done)
        finally:
            active_threads.discard(thread)

    # 检查是否已有线程在运行
    if any(t.name == thread_name for t in active_threads):
        show_warning_message(LANGUAGES[current_language]['partname_generating'], "blue")
        return

    # 提示用户第一次搜索Part Name，没有part name数据文件，将在后台生成
    answer = messagebox.askokcancel(
        f"{LANGUAGES[current_language]['first_partname_1']}",
        f"{LANGUAGES[current_language]['first_partname_2']}\n\n"
        f"{LANGUAGES[current_language]['first_partname_3']}"
    )
    if answer:
        thread = threading.Thread(target=run, daemon=True)
        thread.name = thread_name
        active_threads.add(thread)
        thread.start()

def close_result_list():
    """移除搜索结果"""
    global result_frame, results_tree, window_expanded, preview_check, preview_win
    if result_frame:
        results_tree.destroy()
        results_tree = None
        result_frame.destroy()
        result_frame = None
        preview_check.forget()
        if preview_win and preview_win.winfo_exists():
            preview_win.destroy()  # 关闭预览窗口
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

def setup_click_tooltip():
    """设置鼠标点击results_tree的第二列时弹出浮动窗口，用于partname过长时显示完整内容"""
    # 记录上一次鼠标所点击的条目 ID（用于避免重复触发）
    last_clicked_item = None
    # 提示窗口和标签
    tip_win = None
    tip_label = None
    tip_text = ""      # 存储准备显示的文字

    def on_click(event):
        """鼠标移动事件处理函数"""
        nonlocal last_clicked_item, tip_win, tip_label, tip_text

        # 获取鼠标当前所处的区域（如 cell, heading 等）
        region = results_tree.identify("region", event.x, event.y)
        # 获取鼠标当前点击的哪一列（例如 "#2" 表示第二列）
        col = results_tree.identify_column(event.x)

        # 仅当鼠标点击在“第二列的单元格”时才显示提示
        if region == "cell" and col == "#2":
            item_id = results_tree.identify_row(event.y)

            # 如果鼠标点击了新的单元格（不是上一次的），则更新提示
            if item_id and item_id != last_clicked_item:
                last_clicked_item = item_id
                hide_tip()
                values = results_tree.item(item_id, 'values')

                # 确保第二列存在内容
                if len(values) > 1 and values[1].strip():
                    tip_text = values[1].strip()  # 第二列的内容作为提示文本

                    # 延迟 200 毫秒后显示窗口
                    results_tree.after(200, lambda: show_tip(event.x_root, event.y_root, tip_text))
                else:
                    # 如果第二列没有内容，隐藏提示窗口
                    hide_tip()
        else:
            # 如果鼠标不在第二列或不在 cell 区域，重置状态并隐藏提示
            last_clicked_item = None
            tip_text = ""
            hide_tip()

    def show_tip(x_root, y_root, text):
        """显示提示窗口"""
        nonlocal tip_win, tip_label
        # 获取屏幕上的鼠标位置，稍微偏移避免挡住指针
        x = x_root + 10
        y = y_root + 10

        tip_win = tk.Toplevel(results_tree)
        tip_win.overrideredirect(True)    # 无边框
        tip_win.attributes("-topmost", True)  # 保持最前
        tip_win.geometry(f"+{x}+{y}")

        tip_label = ttk.Label(
            tip_win, text=text,
            relief="solid", borderwidth=1,
            background="#ffffe0", padding=(5, 2)
        )
        tip_label.pack()

    def hide_tip(event=None):
        """隐藏提示窗口"""
        nonlocal tip_win
        if tip_win and tip_win.winfo_exists():
            tip_win.destroy()
            tip_win = None

    # 绑定事件：鼠标点击和离开时触发相应事件
    results_tree.bind("<Button-1>", on_click)
    results_tree.bind("<Leave>", hide_tip)

def show_result_list(result_files, search_type=None):
    """显示搜索结果"""
    global result_frame, results_tree, preview_check, preview_win

    # 销毁原先的预览窗口
    if preview_win and preview_win.winfo_exists():
        preview_win.destroy()

    # 如果没有搜索结果，移除搜索结果显示区域
    if not result_files:
        if result_frame:
            results_tree.destroy()
            results_tree = None
            result_frame.destroy()
            result_frame = None
            if 'preview_check' in globals() and preview_check:
                preview_check.forget()  # 如果没有搜索结果，隐藏 preview_check
            if window_expanded:
                root.geometry(f"{expand_window_width}x{window_height}")
            else:
                root.geometry(f"{window_width}x{window_height}")
        return

    # 显示搜索结果数量
    count = len(result_files)
    msg = f"{count} {LANGUAGES[current_language]['item']}{'s' if count!=1 else ''} {LANGUAGES[current_language]['found_open']}"
    search_hint.config(text=LANGUAGES[current_language]['esc_return'])

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
    tip_label.pack(padx=(int(20*sf), 0), pady=0, side=tk.LEFT)
    # 显示一个关闭按钮用来移除搜索结果
    close_btn = ttk.Button(tip_frame, text="❌", style="Close.TButton", width=3, command=close_result_list)
    close_btn.pack(padx=(0, int(16*sf)), pady=0, side=tk.RIGHT)
    Tooltip(close_btn, lambda: LANGUAGES[current_language]['remove_search_results'], delay=500)

    # 如果是搜索pdf文件，就显示preview_check按钮
    if preview_check and search_type in ("pdf", "lucky", "name"):
        preview_check.pack(side=tk.LEFT, padx=int(20*sf))
    else:
        preview_check.forget()

    # 设置 Treeview 表头和行样式
    style = ttk.Style()
    style.configure("Treeview.Heading", padding=(0, int(4*sf)), background="#A9A9A9", foreground="black", font=("Segoe UI", 9, "bold"))
    style.configure("Treeview", font=("Segoe UI", 9), rowheight=int(25*sf))
    style.map("Treeview", background=[('selected', '#347083')])

    # 添加 Treeview 控件显示结果
    if search_type == "name":
        # 如果是part name搜索，显示part name相关信息
        # 定义列表表头
        columns = (LANGUAGES[current_language]['part_no'], LANGUAGES[current_language]['name'], "PDF", "IPT", "IAM", LANGUAGES[current_language]['drawing'])
        results_tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        results_tree.pack(fill=tk.BOTH, expand=True, padx=(int(17*sf), 0), pady=0)
        # 在 results_tree 上存储排序状态
        results_tree.sort_states = {col: False for col in columns}  # False 表示升序, True 表示降序
        for col in columns:
            results_tree.heading(col, text=col, anchor="w", command=lambda c=col: sort_treeview(c, columns))
        results_tree.column(LANGUAGES[current_language]['part_no'], width=int(60*sf), anchor="w")
        results_tree.column(LANGUAGES[current_language]['name'], width=int(150*sf), anchor="w")
        # 设置PDF、IPT、IAM列宽为0，隐藏这些列
        results_tree.column("PDF", width=0, stretch=tk.NO)
        results_tree.column("IPT", width=0, stretch=tk.NO)
        results_tree.column("IAM", width=0, stretch=tk.NO)
        results_tree.column(LANGUAGES[current_language]['drawing'], width=int(70*sf), anchor="w")
    else:
        # 定义列表表头
        columns = (LANGUAGES[current_language]['file_name'], LANGUAGES[current_language]['created_time'], "Path", "IDW")
        results_tree = ttk.Treeview(result_frame, columns=columns, show="headings")
        results_tree.pack(fill=tk.BOTH, expand=True, padx=(int(17*sf), 0), pady=0)
        # 在 results_tree 上存储排序状态
        results_tree.sort_states = {col: False for col in columns}  # False 表示升序, True 表示降序
        for col in columns:
            results_tree.heading(col, text=col, anchor="w", command=lambda c=col: sort_treeview(c, columns))
        results_tree.column(LANGUAGES[current_language]['file_name'], width=int(150*sf), anchor="w")
        results_tree.column(LANGUAGES[current_language]['created_time'], width=int(135*sf), anchor="w")
        results_tree.column("Path", width=0, stretch=tk.NO)  # 隐藏第三列pdf文件路径
        results_tree.column("IDW", width=0, stretch=tk.NO)  # 隐藏第四列idw文件路径

    # 创建一个垂直滚动条并将其与 Treeview 关联
    scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=results_tree.yview)
    results_tree.configure(yscrollcommand=scrollbar.set)

    # 插入搜索结果
    for index, item in enumerate(result_files):
        tag = 'evenrow' if index % 2 == 0 else 'oddrow'
        if search_type == "name":
            # 如果是part name搜索，插入part name相关信息
            results_tree.insert("", tk.END, values=(item[0], item[1], item[2], item[3], item[4], item[5]), tags=(tag,))
        else:
            results_tree.insert("", tk.END, values=(item[0], item[1], item[2], item[3]), tags=(tag,))
    
    results_tree.tag_configure('evenrow', background='#E6F7FF')
    results_tree.tag_configure('oddrow', background='white')

    results_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    results_tree.bind("<Button-3>", on_right_click)  # 右键菜单
    results_tree.bind("<Double-1>", open_file)
    results_tree.bind("<Return>", open_file)
    results_tree.bind("<<TreeviewSelect>>", on_tree_select)

    # 绑定键盘事件，按下 ESC 键时让 Entry 获得焦点
    def focus_entry(event=None):
        entry_focus()  # 让 Entry 获得焦点
        entry.selection_range(0, tk.END)  # 选中所有文字
    results_tree.bind("<Escape>", focus_entry)

    # 鼠标点击空白区域时清除选中项并关闭预览窗口
    def clear_selection_if_blank(event):
        # 判断用户点击的位置是否在有效的行上
        select_item = results_tree.identify_row(event.y)  # 获取点击位置的行 ID
        if not select_item:  # 如果没有行 ID，说明点击的是空白区域
            results_tree.selection_remove(results_tree.selection())  # 清除选中项
            if preview_win and preview_win.winfo_exists():
                preview_win.destroy()  # 关闭预览窗口
    results_tree.bind("<Button-1>", clear_selection_if_blank)

    # 动态调整窗口大小以显示结果
    root.update_idletasks()
    new_height = int(420*sf) + len(result_files) * int(25*sf) if result_files else window_height
    if window_expanded:
        root.geometry(f"{expand_window_width}x{min(new_height, int(540*sf))}")
    else:
        root.geometry(f"{window_width}x{min(new_height, int(540*sf))}")

    if results_tree.get_children():
        # 如果有搜索结果，焦点移至第一个结果
        first_item = results_tree.get_children()[0]
        results_tree.focus_set()
        results_tree.focus(first_item)
        results_tree.selection_set(first_item)

    if search_type == "name":
        setup_click_tooltip()  # 如果part name名字过长，通过鼠标悬停显示完整名字

def on_right_click(event):
    """给 Treeview 添加右键菜单"""
    item = results_tree.identify_row(event.y)
    if not item:
        return
    
    menu = Menu(results_tree, tearoff=0)

    results_tree.selection_set(item)
    # 根据results_tree的列数，添加不同的右键菜单选项
    # 如果是pdf或者3d搜索，添加打开文件和打开文件所在目录，以及打开idw文件和复制part number的选项
    # 如果是part name搜索，添加打开PDF、IPT、IAM文件，以及复制part number和part name的选项
    if len(results_tree["columns"]) == 4:
        file_name = results_tree.item(item, 'values')[0]
        part_number = file_name.split(".")[0] # 取不含扩展名的文件名作为part number
        if len(part_number) > 15:
            display_part_number = part_number[:12] + "..."
        else:
            display_part_number = part_number
        file_path = results_tree.item(item, 'values')[2]
        idw_path = results_tree.item(item, 'values')[3]
        # 打开文件
        def open_file_right_menu():
            if os.path.exists(file_path):
                os.startfile(file_path)
            else:
                show_warning_message(f"{LANGUAGES[current_language]['file_not_found']}: {file_path}", "red")
        menu.add_command(label=LANGUAGES[current_language]['open'], command=open_file_right_menu)

        # 打开文件所在目录并选中文件
        def open_file_location():
            folder = os.path.dirname(file_path)
            if os.path.exists(folder):
                # 使用 explorer /select 来选中文件
                subprocess.run(["explorer", "/select,", file_path])
        menu.add_command(label=LANGUAGES[current_language]['open_file_location'], command=open_file_location)

        # 打开IDW文件
        if idw_path:
            def open_idw():
                if os.path.exists(idw_path):
                    os.startfile(idw_path)
                else:
                    show_warning_message(f"{LANGUAGES[current_language]['idw_not_found']}: {idw_path}", "red")
            menu.add_command(label=LANGUAGES[current_language]['open_idw'], command=open_idw)
        
        # 复制Part Number
        def copy_part_number():
            root.clipboard_clear()
            root.clipboard_append(part_number)
        menu.add_command(label=f'{LANGUAGES[current_language]['copy']} "{display_part_number}"', command=copy_part_number)
    elif len(results_tree["columns"]) == 6:
        # 如果是part name搜索，添加打开PDF、IPT、IAM的选项，和复制Part Number和Part Name的选项
        part_number = results_tree.item(item, 'values')[0]
        part_name = results_tree.item(item, 'values')[1]
        pdf_path = results_tree.item(item, 'values')[2]
        ipt_path = results_tree.item(item, 'values')[3]
        iam_path = results_tree.item(item, 'values')[4]

        if pdf_path:
            def open_pdf():
                if os.path.exists(pdf_path):
                    os.startfile(pdf_path)
                else:
                    show_warning_message(f"{LANGUAGES[current_language]['pdf_not_found']}: {pdf_path}", "red")
            menu.add_command(label=LANGUAGES[current_language]['open_pdf'], command=open_pdf)

        if ipt_path:
            def open_ipt():
                if os.path.exists(ipt_path):
                    os.startfile(ipt_path)
                else:
                    show_warning_message(f"{LANGUAGES[current_language]['ipt_not_found']}: {ipt_path}", "red")
            menu.add_command(label=LANGUAGES[current_language]['open_ipt'], command=open_ipt)

        if iam_path:
            def open_iam():
                if os.path.exists(iam_path):
                    os.startfile(iam_path)
                else:
                    show_warning_message(f"{LANGUAGES[current_language]['iam_not_found']}: {iam_path}", "red")
            menu.add_command(label=LANGUAGES[current_language]['open_iam'], command=open_iam)

        # 复制Part Number和Part Name
        def copy_part_number_name():
            part_number_name = f"{part_number} - {part_name}"
            root.clipboard_clear()
            root.clipboard_append(part_number_name)
        menu.add_command(label=f'{LANGUAGES[current_language]['copy']} Part No. & Name', command=copy_part_number_name)

    menu.post(event.x_root, event.y_root)

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
        show_warning_message(f"{LANGUAGES[current_language]['unable_pdf']}: {e}", "red")
        return None, None, None # 失败时返回None

def generate_pdf_img(pdf_path, preview_size=(330, 255)):
    """生成 PDF 文件的缩略图，330x255是根据letter纸张比例设置"""
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]  # 读取第一页
        if sf < 2.5:
            pix = page.get_pixmap(matrix=fitz.Matrix(0.5, 0.5))  # 缩小 50% 生成更小的图片
        else:
            pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))  # 大尺寸屏幕，如果屏幕缩放值很大，50%比例的图片会太小，所以用100%比例
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail(preview_size)  # 生成缩略图供预览
        return ImageTk.PhotoImage(img)
    except Exception as e:
        show_warning_message(f"{LANGUAGES[current_language]['unable_preview']}: {e}", "red")
        return None

def generate_preview_thread(file_path):
    """生成预览缩略图的线程函数"""
    global active_threads, last_file
    
    # 获取当前线程并添加到活动线程集合中
    thread = threading.current_thread()
    active_threads.add(thread)

    try:
        orientation, width, height = get_pdf_page_orientation(file_path)  # 判断 PDF 方向，获取长宽数据
        if not orientation:
            # 文件读取失败，重置 last_file
            last_file = None
            return
        # 缩略图缩小1/5
        long_edge = int(width / 5 * sf)
        short_edge = int(height / 5 * sf)
        if orientation == "landscape":
            # 横向
            if short_edge > int(200 * sf):
                short_edge = int(200 * sf)  # 限制高度最大为200*sf，防止缩略图过大
                long_edge = int(short_edge * width / height)
            preview = generate_pdf_img(file_path, (long_edge, short_edge))
        else:
            # 纵向
            if long_edge > int(200 * sf):
                long_edge = int(200 * sf)  # 限制高度最大为200*sf，防止缩略图过大
                short_edge = int(long_edge * height / width)
            preview = generate_pdf_img(file_path, (short_edge, long_edge))

        if not preview:
            # 生成缩略图失败，重置 last_file
            last_file = None
            return
        
        # 在主线程中显示预览窗口
        # 防止连续点击太快导致线程竞争，在显示预览窗口前检查 last_file 是否与当前文件相同
        root.after(0, lambda: show_preview_window(preview, file_path, orientation, long_edge, short_edge)
                   if last_file == file_path else None)
    except Exception as e:
        show_warning_message(f"{LANGUAGES[current_language]['unable_preview']}: {e}", "red")
        last_file = None
    
    finally:
        active_threads.discard(thread)  # 线程结束后移除

def show_preview_window(preview, file_path, orientation, long_edge, short_edge):
    """显示预览缩略图窗口"""
    global preview_win

    hide_warning_message()  # 清除警告信息
    if preview_win and preview_win.winfo_exists():
        preview_win.destroy()  # 销毁旧窗口
        preview_win = None
    # 创建一个新的独立窗口
    preview_win = tk.Toplevel(root)
    preview_win.configure(bg="orange")
    preview_win.overrideredirect(True)
    if topmost_var.get():

        # 如果主窗口置顶，预览窗口也置顶
        preview_win.attributes("-topmost", True)
    # 根据纸张方向设置窗口的大小
    if orientation == "landscape":
        preview_win_width = long_edge + int(10*sf)
        preview_win_height = short_edge + int(10*sf)
    else:
        preview_win_width = short_edge + int(10*sf)
        preview_win_height = long_edge + int(10*sf)
    preview_win.geometry(f"{preview_win_width}x{preview_win_height}")  # 设置窗口大小
    preview_win.resizable(False, False)

    # 显示预览
    label = ttk.Label(preview_win, image=preview, anchor="center")
    label.pack(padx=int(5*sf), pady=int(5*sf))
    label.image = preview  # 保持引用，防止被垃圾回收
    label.bind("<Double-1>", lambda event: open_file(file_path=file_path))  # 双击打开文件

    # 用于关闭预览窗口的label
    close_label = ttk.Label(preview_win, text="✕", style="Close.TLabel")
    close_label.place(relx=1.0, x=int(-5*sf), y=int(5*sf), anchor="ne")  # 右上角
    close_label.bind("<Button-1>", close_preview_window)  # 绑定点击事件

    root.update_idletasks()  # 刷新主窗口状态
    # 预览窗口出现主窗口左侧
    x = root.winfo_rootx() - preview_win_width - int(1*sf)  # 特意与主窗口左侧边框留出1px间距，不紧贴边框更好看一些
    y = root.winfo_rooty() + window_height + int(18*sf)
    # 指定预览窗口显示的位置
    preview_win.geometry(f"+{x}+{y}")

def close_preview_window(event=None):
    # 关闭预览窗口
    global preview_win
    if preview_win and preview_win.winfo_exists():
        preview_win.destroy()
        preview_win = None

def on_tree_select(event):
    """当选中某个搜索结果时，如果是 PDF 文件，则在独立窗口显示预览"""
    global preview_win, results_tree, preview_check, preview_var, last_file
    
    selected_item = results_tree.selection()
    if search_hint.cget("text") != LANGUAGES[current_language]['esc_return']:
        search_hint.config(text=LANGUAGES[current_language]['esc_return'])
    if not selected_item:
        return
    
    # 如果复选框存在且未勾选，关闭预览显示并返回
    if 'preview_check' in globals() and preview_check and not preview_var.get():
        hide_warning_message()  # 隐藏信息
        if preview_win and preview_win.winfo_exists():
            preview_win.destroy()
            preview_win = None
        return
    file_path = results_tree.item(selected_item, "values")[2]  # 获取文件路径
    if file_path is None or file_path == "":
        hide_warning_message()  # 如果没有pdf，隐藏信息
        if preview_win and preview_win.winfo_exists():
            preview_win.destroy()  # 销毁旧窗口
        return
    elif not os.path.exists(file_path):
        show_warning_message(f"{LANGUAGES[current_language]['file_not_found']}: {file_path}", "red")
        return

    # 如果点击的是同一个文件，不重复生成缩略图
    if last_file == file_path:
        if preview_win and preview_win.winfo_exists():
            return
    else:
        last_file = file_path
        if preview_win and preview_win.winfo_exists():
            preview_win.destroy()  # 销毁旧窗口

    if file_path.lower().endswith(".pdf"):
        threading.Thread(target=lambda: generate_preview_thread(file_path), daemon=True).start()
    else:
        if preview_win and preview_win.winfo_exists():
            preview_win.destroy()  # 关闭预览窗口

def on_main_window_move(event):
    """当主窗口移动时，让 preview_win 也随之移动"""
    global preview_win
    if preview_win and preview_win.winfo_exists():
        # 计算新的位置
        x = root.winfo_rootx() - preview_win.winfo_width() - int(1*sf)
        y = root.winfo_rooty() + window_height + int(18*sf)
        preview_win.geometry(f"+{x}+{y}")

def on_focus_in(event):
    # 如果焦点在主窗口，调出预览窗口
    global preview_win
    if root.state() != "iconic" and root.state() != "withdrawn":
        if preview_win and preview_win.winfo_exists():
            if preview_win.state() == "withdrawn":  # 当窗口被隐藏时重新显示
                preview_win.deiconify()
            preview_win.lift()  # 提升预览窗口到最前

def on_window_state_change(event=None):
    # 当窗口状态改变时，隐藏或显示预览窗口
    global preview_win
    if root.state() == "iconic" or root.state() == "withdrawn":  # 窗口最小化或隐藏
        if preview_win and preview_win.winfo_exists():
            preview_win.withdraw()  # 隐藏预览窗口
    else:  # 窗口恢复
        if preview_win and preview_win.winfo_exists():
            preview_win.deiconify()  # 显示预览窗口
            preview_win.lift()  # 提升预览窗口到最前

def show_about():
    """自定义关于信息的窗口"""
    entry_focus()  # 保持焦点在输入框
    if hasattr(root, 'about_win') and root.about_win.winfo_exists():
        # 如果窗口已经打开，就返回
        root.about_win.lift()  # 如果存在，提升窗口到最前
        return

    # 创建自定义关于窗口
    about_win = tk.Toplevel(root)
    root.about_win = about_win  # 将窗口绑定到 root 的属性上
    about_win.withdraw()  # 先隐藏窗口
    about_win.attributes("-topmost", True)
    about_win.title(LANGUAGES[current_language]['about'])
    about_win_width = int(375*sf)
    about_win_height = int(285*sf)
    about_win.geometry(f"{about_win_width}x{about_win_height}")
    about_win.resizable(False, False)

    # 窗口关闭时重置标志位
    def on_close():
        del root.about_win  # 删除引用
        about_win.destroy()

    about_win.protocol("WM_DELETE_WINDOW", on_close)

    # 窗口位置，跟随主窗口居中显示，不考虑Treeview高度
    root.update_idletasks()  # 刷新主窗口状态
    about_win.update_idletasks()
    position_right = int(root.winfo_x() + root.winfo_width()/2 - about_win_width/2)
    position_down = int(root.winfo_y() + window_height - about_win_height + int(6*sf))
    about_win.geometry(f"+{position_right}+{position_down}")
    about_win.deiconify() # 显示窗口
    
    # 设置窗口图标（复用主窗口图标）
    about_win.iconphoto(True, icon)
    
    # 主容器框架
    main_frame = tk.Frame(about_win)
    main_frame.pack(fill=tk.BOTH, expand=True, padx=int(10*sf), pady=int(10*sf))

    # 左侧图标区域
    icon_frame = tk.Frame(main_frame, width=100)
    icon_frame.pack(side=tk.LEFT, fill=tk.Y, padx=int(15*sf), pady=int(20*sf))
    
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
        f"Drawing Finder - Version {ver}",
        f"{LANGUAGES[current_language]['about_text_1']}\n",
        f"{LANGUAGES[current_language]['about_text_2']}\n",
        f"{LANGUAGES[current_language]['about_text_3']}",
    ]

    # 文本标签
    for i, text in enumerate(about_text):
        if i == 0:
            label = ttk.Label(text_frame, text=text, font=("Segoe UI", 10, "bold"), anchor="w")
        else:
            label = ttk.Label(text_frame, text=text, font=("Segoe UI", 9), anchor="w")
        label.pack(anchor="w", fill=tk.X)

        if i == 0:
            # 版本号之后插入占位 frame（用于显示更新信息）
            update_frame = tk.Frame(text_frame, height=int(24*sf))
            update_frame.pack(anchor="w", fill='x')
            update_frame.pack_propagate(False)  # 固定高度占位

    # 后台线程检查更新
    def update_info(about_win, update_frame):
        global active_threads
        # 获取当前线程并添加到活动线程集合中
        thread = threading.current_thread()
        active_threads.add(thread)
        try:
            new_available, latest_ver, download_url = fetch_update_thread()
            if new_available == 1:
                # 有新版本可用
                # 显示更新信息前先判断about窗口是否还存在，如果用户已经关闭了，就不显示
                if about_win.winfo_exists() and update_frame.winfo_exists():
                    about_win.after(0, lambda: show_update_label(update_frame, latest_ver, download_url))            
            elif new_available == 0:
                # 无需更新版本
                if about_win.winfo_exists() and update_frame.winfo_exists():
                    about_win.after(0, lambda: show_update_label(update_frame, None, None))
            elif new_available == -1:
                # 更新检查失败
                if about_win.winfo_exists() and update_frame.winfo_exists():
                    about_win.after(0, lambda: show_update_label(update_frame, None, download_url))
        
        except Exception as e:
            print(f"Update check failed: {e}")
            
        finally:
            active_threads.discard(thread)  # 线程结束后移除

    threading.Thread(target=lambda: update_info(about_win, update_frame), daemon=True).start()

    # 邮箱按钮和地址
    email_frame = tk.Frame(text_frame)
    email_frame.pack(anchor="w")

    email_label = ttk.Label(email_frame, text=LANGUAGES[current_language]['email'], foreground="blue", cursor="hand2", font=("Segoe UI", 9, "underline"))
    email_label.pack(side=tk.LEFT, padx=0)
    email_label.bind("<Button-1>", lambda event: send_email())

    email_label = ttk.Label(email_frame, text=": wtweitang@hotmail.com", font=("Segoe UI", 9))
    email_label.pack(side=tk.LEFT)

    # 设置OK按钮的样式，主要想增加按钮高度，以便于不移动鼠标即可点击按钮关闭窗口
    style.configure("AboutOK.TButton", font=("Segoe UI", 9), padding=(5, 5))
    ok_button = ttk.Button(about_win, text="OK", style="AboutOK.TButton", command=on_close)
    ok_button.pack(padx=int(20*sf), pady=(0, int(20*sf)), side=tk.RIGHT)
    ok_button.focus()

def send_email():
    """打开默认邮件客户端发送邮件"""
    import webbrowser
    try:
        webbrowser.open(f"mailto:wtweitang@hotmail.com?subject=Drawing%20Finder%20Feedback%20v{ver}")
    except Exception as e:
        messagebox.showerror(LANGUAGES[current_language]['error'], f"{LANGUAGES[current_language]['failed_email']}: {e}")

def check_for_updates():
    """检查更新，返回最新版本号和下载链接"""
    try:
        with urllib.request.urlopen(release_url, timeout=10) as response:
            data = json.load(response)
            latest_ver = data["tag_name"].lstrip("v")  # 获取最新版本号，如"v1.3.9" -> "1.3.9"
            # 获取匿名跳转链接（点击后 GitHub 会重定向到 CDN加速地址）
            url = data["assets"][0]["url"]  # 获取资源链接
            headers = {"Accept": "application/octet-stream"}  # 设置 Accept 头部
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req) as dl_response:
                # 最终跳转的真实下载地址
                download_url = dl_response.geturl()
            return latest_ver, download_url
    except Exception as e:
        print(f"Update check failed: {e}")
        return None, "failed"  # GitHub频繁访问API会导致无响应，通过给download_url回传"failed"来返回错误

def show_update_label(parent, latest_ver, download_url):
    """显示版本提示信息"""
    if latest_ver:
        # 有新版本
        update_label = ttk.Label(parent, text=f"{LANGUAGES[current_language]['update_download']} v{latest_ver}", foreground="blue", cursor="hand2", font=("Segoe UI", 8, "italic underline"))
        update_label.bind("<Button-1>", lambda e: webbrowser.open(download_url))
    else:
        if download_url == "failed":
            # 更新检查失败，提示稍后再试
            update_label = ttk.Label(parent, text=f"{LANGUAGES[current_language]['update_check_failed']}", foreground="red", font=("Segoe UI", 8, "italic"))
        else:
            # 已经是最新版本
            update_label = ttk.Label(parent, text=f"{LANGUAGES[current_language]['already_latest']}", foreground="green", font=("Segoe UI", 8, "italic"))

    update_label.pack(anchor="w", fill='x')

def fetch_update_thread():
    """后台线程检查版本更新"""
    global active_threads
    # 获取当前线程并添加到活动线程集合中
    thread = threading.current_thread()
    active_threads.add(thread)

    new_available = 0  # 默认没有新版本
    try:
        latest_ver, download_url = check_for_updates()
        if latest_ver:
            # 比较版本号
            get_ver_parts = list(map(int, latest_ver.split('.')))
            cur_ver_parts = list(map(int, ver.split('.')))
            
            # 对较短的版本号补充 0，以确保两者长度相同
            length = max(len(get_ver_parts), len(cur_ver_parts))
            get_ver_parts.extend([0] * (length - len(get_ver_parts)))
            cur_ver_parts.extend([0] * (length - len(cur_ver_parts)))
            
            # 逐部分比较
            for v1, v2 in zip(get_ver_parts, cur_ver_parts):
                if v1 > v2:
                    # 如果最新版本号大于当前版本号，表示有更新
                    new_available = 1
                    break
                elif v1 < v2:
                    # 当前版本比最新版本号大，说明已经是更新版本，无需更新
                    new_available = 0
                    break
            else:
                # 检查完版本号三个字段都相同，说明正在使用的已经是最新版本，无需更新
                new_available = 0
        else:
            if download_url == "failed":
                # 更新检查失败，回传-1值做后续处理
                new_available = -1

        return new_available, latest_ver, download_url
    
    except Exception as e:
        print(f"Update check failed: {e}")
        
    finally:
        active_threads.discard(thread)  # 线程结束后移除

def change_about_symbol_color():
    """检查是否有新版本，如果有则改变about符号的颜色为蓝色"""
    global active_threads
    # 获取当前线程并添加到活动线程集合中
    thread = threading.current_thread()
    active_threads.add(thread)
    try:
        new_available, _, _  = fetch_update_thread()  # 获取新版本信息
        if new_available == 1:
            root.after(0, lambda: about_label.config(foreground="dodgerblue"))  # 有新版本，变为蓝色

    except Exception as e:
        print(f"Update check failed: {e}")
        
    finally:
        active_threads.discard(thread)  # 线程结束后移除

def reset_window():
    """恢复主窗口到初始状态，停止搜索进程，清空缓存"""
    global result_frame, results_tree, window_expanded, shortcut_frame, last_query, preview_win, refresh_cache_click_count, refresh_cache_click_first_time, last_input

    last_input = ""
    entry_focus()  # 保持焦点在输入框

    # 触发停止事件
    stop_event.set()

    # 等待所有线程结束
    for thread in list(active_threads):
        thread.join(timeout=0.5)  # 最多等 0.5 秒
    
    # 清除所有线程引用
    active_threads.clear()
    
    # 重置停止事件，以便下一次搜索可以正常启动
    stop_event.clear()

    # 清空目录缓存, 重置cache label颜色, 隐藏刷新按钮（与窗口背景同色）
    # 暂时注释掉，reset时不清除缓存
    #directory_cache.clear()
    #cache_label.config(foreground="lightgray")
    #refresh_cache_label.config(foreground="#F0F0F0")

    # 重置缓存刷新点击次数和计时
    refresh_cache_click_count = 0
    refresh_cache_click_first_time = None

    # 清除 partname 数据
    if 'search_partname' in globals() and hasattr(search_partname, 'data'):
        del search_partname.data

    # 隐藏 preview_check
    preview_check.forget()

    # 关闭预览窗口
    if preview_win and preview_win.winfo_exists():
        preview_win.destroy()

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
        expand_btn.config(text=f"{LANGUAGES[current_language]['quick']}   ❯❯")  # 改为 "❯❯"
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

    entry_focus()  # 保持焦点在输入框
    if window_expanded:
        # 如果已有搜索结果，保持显示搜索结果
        if result_frame:
            height = root.winfo_height()  # 获取当前窗口高度
            root.geometry(f"{window_width}x{height}")
        else:
            # 收缩窗口，隐藏快捷按钮框架
            root.geometry(f"{window_width}x{window_height}")  # 恢复到原始大小
        expand_btn.config(text=f"{LANGUAGES[current_language]['quick']}   ❯❯")  # 改为 "❯❯"
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
        expand_btn.config(text=f"{LANGUAGES[current_language]['quick']}   ❮❮")  # 改为 "❮❮"

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
    entry_focus()  # 保持焦点在输入框
    is_checked = topmost_var.get()
    root.attributes("-topmost", is_checked)
    # 如果预览窗口存在，也设置其置顶
    if preview_win and preview_win.winfo_exists():
        preview_win.attributes("-topmost", is_checked)

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
    context_menu.add_command(label=LANGUAGES[current_language]['copy'], command=copy_text)
    context_menu.add_command(label=LANGUAGES[current_language]['cut'], command=cut_text)
    context_menu.add_command(label=LANGUAGES[current_language]['paste'], command=paste_text)

    # 绑定右键事件
    def show_context_menu(event):
        context_menu.tk_popup(event.x_root, event.y_root)

    # 将右键单击事件绑定到 Entry 小部件
    entry_widget.bind("<Button-3>", show_context_menu)

def open_mini_window():
    # 打开 mini 窗口
    # 隐藏预览窗口
    if preview_win and preview_win.winfo_exists():
        preview_win.withdraw()
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
            clear_label_mini.place(in_=mini_entry, relx=1.0, rely=0.5, anchor='e', x=int(-3*sf))  # 显示 X
        else:
            clear_label_mini.place_forget()  # 空内容时，隐藏 X`

    def on_focus_in(event):
        # 当窗口获得焦点时，设置透明度为 1
        mini_win.attributes('-alpha', 1)

    def on_focus_out(event):
        # 当窗口失去焦点时，设置透明度为 0.5
        mini_win.attributes('-alpha', 0.5)

    # 创建 mini 窗口
    mini_win = tk.Toplevel(root)
    mini_win.withdraw()  # 先隐藏窗口
    mini_win.title("Drawing Finder")
    mini_win_width = int(230*sf)
    mini_win_height = int(35*sf)
    mini_win.geometry(f"{mini_win_width}x{mini_win_height}")
    mini_win.attributes("-topmost", True) # 窗口置顶
    mini_win.attributes('-alpha', 1)  # 设置初始窗口透明度
    mini_win.resizable(False, False)

    # 绑定mini窗口焦点事件
    mini_win.bind('<FocusIn>', on_focus_in)  # 窗口获得焦点时触发
    mini_win.bind('<FocusOut>', on_focus_out)  # 窗口失去焦点时触发
    mini_win.bind("<Alt-m>", lambda event: on_close())

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
    mini_entry = ttk.Entry(mini_frame, font=("Consolas", 12), width=13)
    mini_entry.pack(side="left", pady=0, padx=int(5*sf))
    create_entry_context_menu(mini_entry)
    mini_entry.focus()

    # 定义 mini 窗口的搜索操作
    def on_search_mini(event=None):
        mini_entry.focus() # 焦点回位到输入框

        fill_entry_from_clipboard(widget=mini_entry)  # 从剪贴板填充输入框内容
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
    mini_entry.bind("<FocusIn>", update_label_color_mini)  # 当 mini Entry 获取焦点时检查内容

    # 绑定点击事件：点击 Label 的 X 就清空 Entry
    clear_label_mini.bind("<Button-1>", clear_mini_entry)

    # 添加搜索按钮
    search_btn_mini = ttk.Button(mini_frame, text=LANGUAGES[current_language]['search'], width=10, style="All.TButton", command=on_search_mini)
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
    if preview_win and preview_win.winfo_exists():
        preview_win.deiconify()  # 显示预览窗口
        preview_win.lift()
    entry.focus_set()  # 设置焦点到输入框

def on_root_close():
    # 关闭主窗口时清除所有未完成的线程
    global preview_win
    stop_event.set()  # 发送退出信号

    # 强制终止所有子线程。实际上这段代码可以不写，
    # 因为所有线程都已经设置为守护线程daemon=True，会在主线程退出时自动结束，
    # 加了这段代码是为了确保没有遗漏的非守护线程
    for thread in threading.enumerate():
        if thread is not threading.main_thread():
            thread.join(timeout=0.1)  # 等待 0.1 秒

    # 关闭预览窗口
    if preview_win and preview_win.winfo_exists():
        preview_win.destroy()
    root.destroy()  # 关闭窗口

def clear_entry(event=None):
    # 清空输入框内容
    global last_input
    entry.focus()
    entry.delete(0, tk.END)  # 清空输入框内容
    last_input = None
    clear_label.place_forget()  # 清空后，隐藏 X
    search_hint.place_forget()  # 隐藏回车搜索提示

def show_entry_label(event=None):
    # 判断输入框内容是否为空
    if entry.get():
        clear_label.place(in_=entry, relx=1.0, rely=0.5, anchor='e', x=int(-3*sf))  # 显示 X 删除按钮
        if len(entry.get()) < 14:
            # 小于14个字符时显示回车搜索提示，超过时隐藏
            search_hint.place(in_=entry, relx=1.0, rely=0.5, anchor="e", x=int(-20*sf))  # 显示回车搜索提示
        else:
            search_hint.place_forget()
    else:
        clear_label.place_forget()  # 空内容时，隐藏 X
        search_hint.place_forget()  # 隐藏回车搜索提示

def debounce(func, delay=200):
    """装饰器函数，用于防抖"""
    def wrapper(*args, **kwargs):
        if hasattr(wrapper, 'after_id'):
            root.after_cancel(wrapper.after_id)
        wrapper.after_id = root.after(delay, lambda: func(*args, **kwargs))
    return wrapper

def detect_system_language():
    """检测系统语言"""
    try:
        lang, _ = locale.getlocale()
        # Windows 11里会得到English_Canada，French_Canada这样的格式
        if lang:
            return lang.split('_')[0][:2].lower() # 判断前两位小写，兼容旧版本的系统
        else:
            return "en"  # 如果无法获取语言，使用en作为默认值
    except:
        return "en"  # 如果获取语言失败，使用en作为默认值

def switch_language(event=None):
    """切换语言"""
    global current_language, previous_language
    if current_language == "en":
        current_language = "fr"
        previous_language = "en"
        lang_label.config(text="En")
        update_texts()
    elif current_language == "fr":
        current_language = "en"
        previous_language = "fr"
        lang_label.config(text="Fr")
        update_texts()

def update_texts():
    """更新主窗口控件的文本"""
    global default_parts_path, changed_parts_path
    prompt_label.config(text=LANGUAGES[current_language]['input'])
    search_btn.config(text=LANGUAGES[current_language]['search_pdf'])
    lucky_btn.config(text=LANGUAGES[current_language]['lucky'])
    search_3d_btn.config(text=LANGUAGES[current_language]['3d'])
    search_cache_btn.config(text=LANGUAGES[current_language]['vault'])
    search_partname_btn.config(text=LANGUAGES[current_language]['partname'])
    reset_btn.config(text=LANGUAGES[current_language]['reset'])
    if expand_btn.cget("text").endswith('❯❯'):
        expand_btn.config(text=f"{LANGUAGES[current_language]['quick']}   ❯❯")
    else:
        expand_btn.config(text=f"{LANGUAGES[current_language]['quick']}   ❮❮")
    if directory_label.cget("text").startswith(LANGUAGES[previous_language]['default_parts_dir']):
        directory_label.config(text=f"{LANGUAGES[current_language]['default_parts_dir']} {default_parts_path}")
    else:
        directory_label.config(text=f"{LANGUAGES[current_language]['parts_dir']} {changed_parts_path}")
    change_label.config(text=LANGUAGES[current_language]['change'])
    default_label.config(text=LANGUAGES[current_language]['default'])
    preview_check.config(text=LANGUAGES[current_language]['preview'])
    create_entry_context_menu(entry) # 更新主窗口输入框的右键菜单语言
    if search_hint.cget("text") == LANGUAGES["en"]['enter_search'] or search_hint.cget("text") == LANGUAGES["fr"]['enter_search']:
        search_hint.config(text=LANGUAGES[current_language]['enter_search'])
    elif search_hint.cget("text") == LANGUAGES["en"]['esc_return'] or search_hint.cget("text") == LANGUAGES["fr"]['esc_return']:
        search_hint.config(text=LANGUAGES[current_language]['esc_return'])
    show_warning_message(LANGUAGES[current_language]['cpoied_part_number'], "blue")  # 更新从剪贴板读取的提示

def entry_focus():
    # 焦点回到输入框，并重置回车搜索的提示
    entry.focus()
    if search_hint.cget("text") != LANGUAGES[current_language]['enter_search']:
        search_hint.config(text=LANGUAGES[current_language]['enter_search'])

# 创建主窗口
try:
    # 适配系统缩放比例
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    ScaleFactor=ctypes.windll.shcore.GetScaleFactorForDevice(0)
    sf = ScaleFactor/100
    tk_sf = sf*(96/72)
    
    # 获取系统的默认语言
    system_lang = detect_system_language()
    if system_lang in ['fr', 'en']:
        current_language = system_lang
    else:
        current_language = "en"  # 默认英语
        
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
    root.title("Drawing Finder")
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
    style.configure("RefreshCache.TLabel", font=("Segoe UI", 12))
    style.configure("Tooltip.TLabel", background="#ffffe0")
    style.configure("Clear.TLabel", background="white")
    style.configure("Enter.TLabel", background="white", font=("Segoe UI", 9))
    style.configure("Preview.TCheckbutton", font=("Segoe UI", 9))
    style.configure("Close.TLabel", foreground="red", background="white", font=("Segoe UI", 9))
    style.configure("Lang.TLabel", font=("Consolas", 9), background="lightblue")

    # 第一行标签的框架
    label_frame = ttk.Frame(root)
    label_frame.pack(pady=(int(15*sf), int(5*sf)), anchor="w")
    label_frame.pack_propagate(False)  # 禁止自动调整大小
    label_frame.config(width=int(345*sf), height=int(23*sf))  # 设置frame大小

    # 标签放在第一行
    prompt_label = ttk.Label(label_frame, text=LANGUAGES[current_language]['input'], font=("Segoe UI", 9), anchor="w")
    prompt_label.pack(side=tk.LEFT, padx=(int(20*sf), 0))

    # 添加置顶选项
    # 创建一个 IntVar 绑定复选框的状态（0 未选中，1 选中）
    topmost_var = tk.IntVar()

    # 创建复选框，用于控制窗口置顶
    checkbox = ttk.Checkbutton(label_frame, text="📌", variable=topmost_var, style="Top.TCheckbutton", command=toggle_topmost)
    checkbox.pack(side=tk.RIGHT, padx=int(10*sf))
    Tooltip(checkbox, lambda: LANGUAGES[current_language]['tip_top'], delay=500)
    def toggle_topmost_hotkey():
        # 快捷键无法直接改变checkbox状态，需要手动切换
        topmost_var.set(not topmost_var.get())
        toggle_topmost()
    root.bind("<Alt-t>", lambda event: toggle_topmost_hotkey())

    # 添加切换mini窗口的按钮
    mini_search_label = ttk.Label(label_frame, text="🍀", font=("Segoe UI", 10), cursor="hand2")
    mini_search_label.pack(side=tk.RIGHT, padx=int(5*sf))
    mini_search_label.bind("<Button-1>", lambda event: open_mini_window())
    Tooltip(mini_search_label, lambda: LANGUAGES[current_language]['tip_mini'], delay=500)
    root.bind("<Alt-m>", lambda event: open_mini_window())

     # 按钮宽度，输入框宽度，Label宽度和位置，无法根据缩放比例在布局内进行同比例调整，所以指定具体值
    if ScaleFactor == 100:
        btn_width = 21
        entry_width = 25
        parts_dir_width = 33
        parts_y_position = int(5*sf-2)
    elif ScaleFactor == 125:
        btn_width = 20
        entry_width = 25
        parts_dir_width = 37
        parts_y_position = int(5*sf)
    elif ScaleFactor == 150:
        btn_width = 19
        entry_width = 25
        parts_dir_width = 35
        parts_y_position = int(5*sf+2)
    elif ScaleFactor == 175:
        btn_width = 21
        entry_width = 26
        parts_dir_width = 36
        parts_y_position = int(5*sf+4)
    elif ScaleFactor == 200:
        btn_width = 20
        entry_width = 25
        parts_dir_width = 38
        parts_y_position = int(5*sf+6)
    else:
        btn_width = 20
        entry_width = 25
        parts_dir_width = 35
        # 对于大于200的缩放比例，进行特殊处理，使最后一行始终位于窗口底部
        parts_y_position = int(5*sf+(14*sf-22))

    # 创建输入框框架
    entry_frame = ttk.Frame(root)
    entry_frame.pack(pady=0, anchor="w")
    entry_frame.pack_propagate(False)  # 禁止自动调整大小
    entry_frame.config(width=int(345*sf), height=int(58*sf))  # 设置frame大小
    entry = ttk.Entry(entry_frame, width=entry_width, font=("Consolas", 16))
    entry.pack(padx=int(20*sf), pady=(int(5*sf), int(4*sf)), anchor="w")
    create_entry_context_menu(entry)
    entry.focus()
    entry.bind("<Return>", lambda event: search_pdf_files())
    entry.bind("<Escape>", lambda event: (entry.selection_range(0, tk.END), entry.focus_set()))
    entry.bind("<Button-1>", show_search_history)  # 点击输入框时显示历史记录

    # 清除 Entry 内容的Label，初始不显示
    clear_label = ttk.Label(entry_frame, text="✕", font=('Segoe UI', 10), foreground='red', style="Clear.TLabel", cursor="arrow")
    # 回车搜索的提示，初始不显示
    search_hint = ttk.Label(entry_frame, text=LANGUAGES[current_language]['enter_search'], foreground="gray", style="Enter.TLabel")
    search_anim = SearchAnimation(search_hint.master)  # 动画效果

    # 监听 Entry 内容变化来更新 Label 的颜色， 同时实时更新匹配历史
    last_input = None
    def key_release(event=None):
        global last_input
        current_input = entry.get().lower()
        if current_input != last_input:
            # 键盘事件后对比输入框的内容，如果产生变化，就显示历史记录，避免使用快捷键时也触发显示
            show_search_history(event)
            show_entry_label(event)
            last_input = current_input

    entry.bind("<KeyRelease>", key_release)
    entry.bind("<FocusIn>", show_entry_label)  # 当 Entry 获取焦点时根据内容决定是否显示回车搜索和清除label
    root.bind("<Alt-c>", lambda event: clear_entry())  # 绑定 Alt+C 快捷键清空输入框

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
    search_btn = ttk.Button(button_frame, text=LANGUAGES[current_language]['search_pdf'], width=btn_width, style="All.TButton", command=search_pdf_files)
    search_btn.grid(row=0, column=0, padx=(int(5*sf), int(10*sf)), pady=int(8*sf))
    Tooltip(search_btn, lambda: LANGUAGES[current_language]['tip_search_pdf'], delay=500)
    root.bind("<Alt-p>", lambda event: search_pdf_files())

    # I'm Feeling Lucky 按钮
    lucky_btn = ttk.Button(button_frame, text=LANGUAGES[current_language]['lucky'], style="All.TButton", width=btn_width, command=feeling_lucky)
    lucky_btn.grid(row=0, column=1, padx=(int(10*sf), int(5*sf)), pady=int(8*sf))
    Tooltip(lucky_btn, lambda: LANGUAGES[current_language]['tip_lucky'], delay=500)
    root.bind("<Alt-l>", lambda event: feeling_lucky())

    # Search 3D Drawing 按钮
    search_3d_btn = ttk.Button(button_frame, text=LANGUAGES[current_language]['3d'], style="All.TButton", width=btn_width, command=search_3d_files)
    search_3d_btn.grid(row=1, column=0, padx=(int(5*sf), int(10*sf)), pady=int(8*sf))
    Tooltip(search_3d_btn, lambda: LANGUAGES[current_language]['tip_3d'], delay=500)
    root.bind("<Alt-d>", lambda event: search_3d_files())

    # Search Vault Cache 按钮
    search_cache_btn = ttk.Button(button_frame, text=LANGUAGES[current_language]['vault'], style="All.TButton", width=btn_width, command=search_vault_cache)
    search_cache_btn.grid(row=1, column=1, padx=(int(10*sf), int(5*sf)), pady=int(8*sf))
    Tooltip(search_cache_btn, lambda: LANGUAGES[current_language]['tip_vault'], delay=500)
    root.bind("<Alt-v>", lambda event: search_vault_cache())

    # Search Part Name 按钮
    search_partname_btn = ttk.Button(button_frame, text=LANGUAGES[current_language]['partname'], width=btn_width, style="All.TButton", command=search_partname)
    search_partname_btn.grid(row=2, column=0, padx=(int(5*sf), int(10*sf)), pady=int(8*sf))
    Tooltip(search_partname_btn, lambda: LANGUAGES[current_language]['tip_partname'], delay=500)
    root.bind("<Alt-n>", lambda event: search_partname())

    # 扩展按钮
    expand_btn = ttk.Button(button_frame, text=f"{LANGUAGES[current_language]['quick']}   ❯❯", width=btn_width, style="All.TButton", command=toggle_window_size)
    expand_btn.grid(row=2, column=1, padx=(int(10*sf), int(5*sf)), pady=int(8*sf))
    Tooltip(expand_btn, lambda: LANGUAGES[current_language]['tip_quick'], delay=500)
    root.bind("<Alt-q>", lambda event: toggle_window_size())

    # 显示默认目录及更改功能
    directory_frame = tk.Frame(root)
    directory_frame.pack(anchor="w", padx=int(14*sf), pady=parts_y_position, fill="x")
    directory_label = ttk.Label(directory_frame, text=f"{LANGUAGES[current_language]['default_parts_dir']} {default_parts_path}", font=("Segoe UI", 8), width=parts_dir_width, anchor="w")
    directory_label.pack(side=tk.LEFT)
    Tooltip(directory_label, lambda: directory_label.cget("text"), delay=500)

    # Change 按钮
    change_label = ttk.Label(directory_frame, text=LANGUAGES[current_language]['change'], style="Change.TLabel", cursor="hand2")
    change_label.pack(side=tk.LEFT, padx=(0, int(8*sf)))
    Tooltip(change_label, lambda: LANGUAGES[current_language]['tip_change'], delay=500)
    change_label.bind("<Button-1>", lambda event: update_directory())

    # Default 按钮
    default_label = ttk.Label(directory_frame, text=LANGUAGES[current_language]['default'], style="Change.TLabel", cursor="hand2")
    default_label.pack(side=tk.LEFT, padx=0)
    Tooltip(default_label, lambda: LANGUAGES[current_language]['tip_default'], delay=500)
    default_label.bind("<Button-1>", lambda event: reset_to_default_directory())

    # About 按钮
    about_frame = ttk.Frame(root)
    about_frame.pack(anchor="w", padx=0, pady=0)
    about_frame.pack_propagate(False)  # 禁止自动调整大小
    about_frame.config(width=int(345*sf), height=int(35*sf))  # 设置frame大小
    about_label = ttk.Label(about_frame, text="ⓘ", style="About.TLabel", cursor="hand2")
    about_label.pack(side=tk.RIGHT, padx=int(10*sf), pady=(int(3*sf), int(4*sf)))
    Tooltip(about_label, lambda: LANGUAGES[current_language]['tip_about'], delay=500)
    about_label.bind("<Button-1>", lambda event: show_about())
    root.bind("<Alt-a>", lambda event: show_about())

    # 添加刷新缓存标志
    # 先设置与窗口背景同色隐藏刷新标志，等有缓存完成后再显示
    refresh_cache_label = ttk.Label(about_frame, text="⟳", style="RefreshCache.TLabel", foreground="#F0F0F0")
    refresh_cache_label.pack(side=tk.RIGHT, padx=0, pady=int(4*sf))
    # 只有当refresh_cache_label是非隐藏的状态（非"#F0F0F0"颜色），才显示Tooltip
    def get_refresh_tooltip():
        current_fg = str(refresh_cache_label.cget("foreground")).lower()  # 获取当前字体颜色，转为小写以统一比较
        if current_fg == "#f0f0f0":
            return ""  # 颜色为 #F0F0F0 时返回空字符串，不显示 Tooltip
        return LANGUAGES[current_language]['tip_refresh']  # 其他颜色时显示提示
    Tooltip(refresh_cache_label, lambda: get_refresh_tooltip(), delay=500)
    refresh_cache_label.bind("<Button-1>", on_refresh_cache_click)

    # 显示缓存状态, 灰色无缓存，绿色缓存已完成，红色正在缓存
    cache_label = ttk.Label(about_frame, text="●", style="Cache.TLabel")
    cache_label.pack(side=tk.RIGHT, padx=0, pady=int(4*sf))
    Tooltip(cache_label, get_cache_str, delay=500)
    cache_label.bind("<Button-1>", on_refresh_cache_click)  # 绑定点击刷新缓存函数，在点击该处时也能刷新缓存

    # 语言设置label
    if current_language == "fr":
        lang_label = ttk.Label(about_frame, text="En", style="Lang.TLabel", cursor="hand2")
    else:
        lang_label = ttk.Label(about_frame, text="Fr", style="Lang.TLabel", cursor="hand2")
    lang_label.pack(side=tk.RIGHT, padx=int(14*sf), pady=(int(5*sf), int(4*sf)))
    Tooltip(lang_label, lambda: LANGUAGES[current_language]['language'], delay=500)
    lang_label.bind("<Button-1>", switch_language)  # 点击切换语言
    root.bind("<Alt-s>", lambda event: switch_language())

    # 创建用于preview check的frame，用来占位
    preview_frame = ttk.Frame(about_frame)
    preview_frame.pack(side=tk.LEFT, padx=0, pady=0)
    preview_frame.pack_propagate(False)  # 禁止自动调整大小
    preview_frame.config(width=int(120*sf), height=int(25*sf))  # 设置frame大小
    # 添加 preview_check 复选框
    preview_var = tk.BooleanVar(value=True)  # 默认选中
    preview_check = ttk.Checkbutton(
        preview_frame, text=LANGUAGES[current_language]['preview'], variable=preview_var, style="Preview.TCheckbutton", 
        command=lambda: (on_tree_select(None), entry_focus())
    )
    Tooltip(preview_check, lambda: LANGUAGES[current_language]['show_preview'], delay=500)

    # Reset 按钮
    reset_btn = ttk.Button(about_frame, text=LANGUAGES[current_language]['reset'], width=10, style="All.TButton", command=reset_window)
    reset_btn.pack(side=tk.LEFT, padx=(int(15*sf),0), pady=(int(4*sf)))
    Tooltip(reset_btn, lambda: LANGUAGES[current_language]['tip_reset'], delay=500)
    root.bind("<Alt-r>", lambda event: reset_window())

    # 显示可以直接读取剪贴版的提示
    root.after(100, lambda: show_warning_message(LANGUAGES[current_language]['cpoied_part_number'], "blue"))
    # 2秒后台线程根据版本信息改变about符号颜色
    root.after(2000, lambda: threading.Thread(target=change_about_symbol_color, daemon=True).start())  
    # 运行主循环
    root.mainloop()
except Exception as e:
    print(f"An error occurred: {e}")
    messagebox.showerror("Error", f"An error occurred: {e}")