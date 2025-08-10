import os
import winreg
import ctypes
from PIL import Image, ImageTk
import tkinter as tk
import win32con
import win32gui
import win32ui
from win32com.client import Dispatch

# 多語言字典
LANG = {
    'zh': {
        'title': '搜索自動開機程序 ',
        'refresh': '刷新自動開機程序',
        'toggle_lang': '切換語言',
        'no_icon': '（無抓取到圖示）',
        'Author':'作者:Koala'
    },
    'en': {
        'title': 'Search Startup Programs',
        'refresh': 'Refresh Startup List',
        'toggle_lang': 'Switch Language',
        'no_icon': '(No got Icon)',
        'Author':'Author:Koala'
    }
}
current_lang = 'zh'  # 預設中文

def get_registry_autorun(root, path):
    items = []
    try:
        with winreg.OpenKey(root, path) as key:
            i = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(key, i)
                    items.append((name, value))
                    i += 1
                except OSError:
                    break
    except FileNotFoundError:
        pass
    return items

def get_startup_folder_autorun():
    items = []
    startup_path = os.path.expandvars(r'%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup')
    try:
        for f in os.listdir(startup_path):
            fullpath = os.path.join(startup_path, f)
            items.append((os.path.splitext(f)[0], fullpath))
    except FileNotFoundError:
        pass
    return items

def hicon_to_pil(hicon):
    try:
        info = win32gui.GetIconInfo(hicon)
        hbmColor = info[4]
        hbmMask  = info[3]
        if not hbmColor:
            return None
        bmpinfo = win32gui.GetObject(hbmColor)
        width, height = bmpinfo.bmWidth, bmpinfo.bmHeight
        bmp = win32ui.CreateBitmapFromHandle(hbmColor)
        bmpstr = bmp.GetBitmapBits(True)
        img = Image.frombuffer('RGBA', (width, height), bmpstr, 'raw', 'BGRA', 0, 1)
        win32gui.DeleteObject(hbmColor)
        win32gui.DeleteObject(hbmMask)
        return img
    except Exception as e:
        print(f'icon抽取失敗: {e}')
        return None

def extract_icon(exe_path, large=False):
    if not os.path.isfile(exe_path):
        return None
    large_icons = (ctypes.c_void_p * 1)()
    small_icons = (ctypes.c_void_p * 1)()
    count = ctypes.windll.shell32.ExtractIconExW(exe_path, 0, large_icons, small_icons, 1)
    if count == 0:
        return None
    hicon = large_icons[0] if large else small_icons[0]
    img = hicon_to_pil(hicon)
    if img:
        img = img.resize((24,24), Image.LANCZOS)
    return img

def resolve_lnk_path(lnk_path):
    try:
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)
        target = shortcut.Targetpath
        return target
    except Exception as e:
        print(f"無法解析捷徑 {lnk_path}: {e}")
        return None

def find_exe_path(exe_name):
    search_dirs = [
        os.environ.get("ProgramFiles", r"C:\\Program Files"),
        os.environ.get("ProgramFiles(x86)", r"C:\\Program Files (x86)"),
        os.environ.get("SystemRoot", r"C:\\Windows"),
    ]
    for base_dir in search_dirs:
        for root, dirs, files in os.walk(base_dir):
            if exe_name in files:
                return os.path.join(root, exe_name)
    return None

def parse_exe_from_command(cmd):
    # 先優先找引號包住的路徑
    if '"' in cmd:
        segs = [seg for seg in cmd.split('"') if seg.strip()]
        for seg in segs:
            if ".exe" in seg.lower():
                ix = seg.lower().find('.exe')
                path = seg[:ix+4].strip()
                if os.path.isabs(path):
                    return path
    # 沒引號的話抓第一個 .exe 可疑片段
    parts = cmd.strip().split(' ')
    for part in parts:
        if part.lower().endswith('.exe'):
            return part
    # 退而求其次，抓第一個片段
    return parts[0]

root = tk.Tk()
root.title(LANG[current_lang]['title'])
root.geometry("1100x650")
frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True)
canvas = tk.Canvas(frame)
scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
scrollable_frame = tk.Frame(canvas)
scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")
    )
)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)
canvas.pack(side="left", fill=tk.BOTH, expand=True)
scrollbar.pack(side="right", fill=tk.Y)
icon_refs = []  # 防止圖片被垃圾回收

def refresh_list():
    for widget in scrollable_frame.winfo_children():
        widget.destroy()
    icon_refs.clear()
    items = []
    items += get_registry_autorun(winreg.HKEY_CURRENT_USER, r"Software\\Microsoft\\Windows\\CurrentVersion\\Run")
    items += get_registry_autorun(winreg.HKEY_LOCAL_MACHINE, r"Software\\Microsoft\\Windows\\CurrentVersion\\Run")
    items += get_startup_folder_autorun()
    for name, path in items:
        frm = tk.Frame(scrollable_frame)
        exe_path_raw = parse_exe_from_command(path)
        icon_img = None
        # .lnk 捷徑解析
        if exe_path_raw.lower().endswith('.lnk') and os.path.exists(exe_path_raw):
            target_path = resolve_lnk_path(exe_path_raw)
            if target_path and target_path.lower().endswith('.exe') and os.path.exists(target_path):
                img = extract_icon(target_path)
                icon_img = ImageTk.PhotoImage(img) if img else None
        # .exe 路徑存在直接抽 icon，沒絕對路徑就搜尋系統常見資料夾
        elif exe_path_raw.lower().endswith('.exe'):
            if not os.path.isabs(exe_path_raw):
                exe_only = os.path.basename(exe_path_raw)
                found_path = find_exe_path(exe_only)
                if found_path:
                    img = extract_icon(found_path)
                    icon_img = ImageTk.PhotoImage(img) if img else None
            elif os.path.exists(exe_path_raw):
                img = extract_icon(exe_path_raw)
                icon_img = ImageTk.PhotoImage(img) if img else None
        if icon_img:
            icon_lbl = tk.Label(frm, image=icon_img)
            icon_lbl.pack(side=tk.LEFT)
            icon_refs.append(icon_img)
        else:
            icon_lbl = tk.Label(frm, text=LANG[current_lang]['no_icon'], width=10, fg='gray')
            icon_lbl.pack(side=tk.LEFT)
        txt_lbl = tk.Label(frm, text=f"{name} : {path}", anchor="w", font=("Consolas", 11))
        txt_lbl.pack(side=tk.LEFT)
        frm.pack(fill=tk.X, padx=2, pady=2)

def toggle_language():
    global current_lang
    current_lang = 'en' if current_lang == 'zh' else 'zh'
    root.title(LANG[current_lang]['title'])
    btn_refresh.config(text=LANG[current_lang]['refresh'])
    btn_toggle_lang.config(text=LANG[current_lang]['toggle_lang'])
    author.config(text=LANG[current_lang]['Author'])
    refresh_list()

btn_refresh = tk.Button(root, text=LANG[current_lang]['refresh'], font=("Microsoft JhengHei", 12), command=refresh_list)
btn_refresh.pack(pady=8)

btn_toggle_lang = tk.Button(root, text=LANG[current_lang]['toggle_lang'], font=("Microsoft JhengHei", 10), command=toggle_language)
btn_toggle_lang.pack(pady=4)
author = tk.Label(root, text=LANG[current_lang]['Author'], font=("Microsoft JhengHei", 10))
author.pack()
refresh_list()
root.mainloop()
