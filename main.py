import tkinter as tk
from tkinter import simpledialog, messagebox
import random
import winreg
from datetime import datetime
import os

# 注册表路径
REGISTRY_PATH = r"Software\StudentIDDraw"

# 默认设置
DEFAULT_ALLOW_REPEAT = False
DEFAULT_MIN_ID = 1
DEFAULT_MAX_ID = 45
PASSWORD = "Admin@123"

selected_ids = []
is_running = False
job = None


# 注册表工具函数
def set_registry_value(value_name, value, value_type=winreg.REG_SZ):
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH) as reg_key:
            winreg.SetValueEx(reg_key, value_name, 0, value_type, value)
    except WindowsError as e:
        print(f"Error setting registry value: {e}")


def get_registry_value(value_name, default_value):
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, winreg.KEY_READ
        ) as reg_key:
            value, _ = winreg.QueryValueEx(reg_key, value_name)
            return value
    except WindowsError:
        return default_value


def delete_registry_value(value_name):
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, winreg.KEY_WRITE
        ) as reg_key:
            winreg.DeleteValue(reg_key, value_name)
    except WindowsError as e:
        print(f"Error deleting registry value: {e}")


# 加载已抽取的学号
def load_selected_ids():
    global selected_ids
    selected_ids_str = get_registry_value("SelectedIDs", "")
    if selected_ids_str:
        selected_ids = []
        for item in selected_ids_str.split(","):
            parts = item.split("_")
            if len(parts) == 2:
                id, timestamp = parts
                timestamp = datetime.fromtimestamp(int(timestamp)).strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
                selected_ids.append((id, timestamp))
            else:
                print(f"Error parsing item: {item}, resetting to NULL")
                delete_registry_value("SelectedIDs")
                selected_ids = []
    else:
        selected_ids = []
    return selected_ids


# 保存已抽取的学号
def save_selected_ids():
    global selected_ids
    selected_ids_str = ",".join(
        [
            f"{id}_{int(datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S').timestamp())}"
            for id, timestamp in selected_ids
        ]
    )
    set_registry_value("SelectedIDs", selected_ids_str)


# 加载设置
def load_settings():
    allow_repeat = (
        get_registry_value("AllowRepeat", str(DEFAULT_ALLOW_REPEAT)) == "True"
    )
    min_id = int(get_registry_value("MinID", str(DEFAULT_MIN_ID)))
    max_id = int(get_registry_value("MaxID", str(DEFAULT_MAX_ID)))
    return allow_repeat, min_id, max_id


# 保存设置
def save_settings(allow_repeat, min_id, max_id):
    set_registry_value("AllowRepeat", str(allow_repeat))
    set_registry_value("MinID", str(min_id))
    set_registry_value("MaxID", str(max_id))


# 生成随机的学号
def generate_random_id(selected_ids, allow_repeat, min_id, max_id):
    if not allow_repeat:
        all_ids = set(range(min_id, max_id + 1))
        selected_id_numbers = [int(id) for id, _ in selected_ids]
        remaining_ids = list(all_ids - set(selected_id_numbers))
        if not remaining_ids:
            return None  # 如果所有学号都已抽取完毕，则返回 None
        return random.choice(remaining_ids)
    else:
        return random.randint(min_id, max_id)

# 更新历史记录列表
def update_history_list():
    global history_list
    if history_list:
        history_list.delete(0, tk.END)
        for id, timestamp in selected_ids:
            history_list.insert(tk.END, f"{id} ({timestamp})")

# 每隔 5 毫秒更新一次显示的随机学号
def update_number(allow_repeat, min_id, max_id):
    global job
    if is_running:
        random_id = generate_random_id(selected_ids, allow_repeat, min_id, max_id)
        label.config(text=str(random_id))
        job = root.after(5, update_number, allow_repeat, min_id, max_id)

# 抽取学号按钮点击事件处理函数
def draw_student_id():
    global is_running, job, selected_ids
    allow_repeat, min_id, max_id = load_settings()

    if is_running:
        # 停止跳动，保存最终学号
        root.after_cancel(job)
        id = label.cget("text")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        selected_ids.append((str(id), timestamp))
        save_selected_ids()
        if settings_window:
            update_history_list()
        button.config(text="开始")
        is_running = False
    else:
        # 开始跳动
        is_running = True
        button.config(text="停止")
        update_number(allow_repeat, min_id, max_id)


# 设置按钮点击事件处理函数
def open_settings():
    global settings_window
    global history_list

    if settings_window:
        settings_window.lift()
        return

    settings_window = tk.Toplevel(root)
    settings_window.title("设置")

    allow_repeat, min_id, max_id = load_settings()

    # 使用 grid 管理器来创建一个自适应大小的界面
    settings_window.grid_columnconfigure(1, weight=1)
    settings_window.grid_rowconfigure(4, weight=1)

    load_selected_ids()

    history_frame = tk.Frame(settings_window)
    history_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
    history_frame.grid_columnconfigure(0, weight=1)
    history_frame.grid_rowconfigure(0, weight=1)

    history_list = tk.Listbox(history_frame, selectmode=tk.SINGLE)
    history_list.grid(row=0, column=0, sticky="nsew")

    scrollbar = tk.Scrollbar(
        history_frame, orient="vertical", command=history_list.yview
    )
    scrollbar.grid(row=0, column=1, sticky="ns")
    history_list.config(yscrollcommand=scrollbar.set)

    update_history_list()

    def add_id():
        new_id = simpledialog.askstring("新增学号", "请输入新的学号:")
        if new_id:
            if new_id.isdigit():
                new_id = int(new_id)
                if min_id <= new_id <= max_id:
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    selected_ids.append((str(new_id), timestamp))
                    update_history_list()
                else:
                    messagebox.showerror("错误", f"学号必须在 {min_id} 和 {max_id} 之间")
            else:
                messagebox.showerror("错误", "请输入有效的学号")

    def delete_id():
        selected = history_list.curselection()
        if selected:
            index = selected[0]
            selected_ids.pop(index)
            update_history_list()
            if index < len(selected_ids):
                history_list.select_set(index)
            elif selected_ids:
                history_list.select_set(index - 1)
        else:
            messagebox.showerror("错误", "请选择一个学号进行删除")

    def edit_id():
        selected = history_list.curselection()
        if selected:
            index = selected[0]
            current_id, _ = selected_ids[index]
            new_id = simpledialog.askstring(
                "修改学号", "请输入新的学号:", initialvalue=str(current_id)
            )
            if new_id:
                if new_id.isdigit():
                    new_id = int(new_id)
                    if min_id <= new_id <= max_id:
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        selected_ids[index] = (str(new_id), timestamp)
                        update_history_list()
                        history_list.select_set(index)
                    else:
                        messagebox.showerror("错误", f"学号必须在 {min_id} 和 {max_id} 之间")
                else:
                    messagebox.showerror("错误", "请输入有效的学号")
        else:
            messagebox.showerror("错误", "请选择一个学号进行修改")

    def clear_all_ids():
        global selected_ids
        selected_ids = []
        update_history_list()

    tk.Button(settings_window, text="新增记录", font=("宋体", 10), command=add_id).grid(
        row=2, column=0, padx=5, pady=5, sticky="ew"
    )
    tk.Button(settings_window, text="删除记录", font=("宋体", 10), command=delete_id).grid(
        row=2, column=1, padx=5, pady=5, sticky="ew"
    )
    tk.Button(settings_window, text="修改记录", font=("宋体", 10), command=edit_id).grid(
        row=3, column=0, padx=5, pady=5, sticky="ew"
    )
    tk.Button(
        settings_window, text="清空记录", font=("宋体", 10), command=clear_all_ids
    ).grid(row=3, column=1, padx=5, pady=5, sticky="ew")

    allow_repeat_var = tk.BooleanVar(value=allow_repeat)
    tk.Checkbutton(
        settings_window, text="允许重复抽号", font=("宋体", 10), variable=allow_repeat_var
    ).grid(row=5, column=0, padx=5, pady=5, sticky="w")

    tk.Label(settings_window, text="学号范围:", font=("宋体", 10)).grid(
        row=6, column=0, padx=5, pady=5, sticky="e"
    )
    min_id_var = tk.StringVar(value=min_id)
    max_id_var = tk.StringVar(value=max_id)
    tk.Entry(settings_window, textvariable=min_id_var).grid(
        row=6, column=1, padx=5, pady=5, sticky="ew"
    )
    tk.Entry(settings_window, textvariable=max_id_var).grid(
        row=7, column=1, padx=5, pady=5, sticky="ew"
    )

    def save_changes():
        entered_password = simpledialog.askstring("密码", "请输入密码:", show="*")
        if entered_password == PASSWORD:
            save_settings(
                allow_repeat_var.get(), int(min_id_var.get()), int(max_id_var.get())
            )
            save_selected_ids()
            messagebox.showinfo("提示", "设置已保存")
            close_settings_window()
        else:
            messagebox.showerror("错误", "密码错误")

    tk.Button(settings_window, text="保存", font=("宋体", 10), command=save_changes).grid(
        row=8, column=0, columnspan=2, padx=5, pady=20, sticky="ew"
    )

    # 自动调整窗口大小
    settings_window.attributes("-topmost", 1)
    settings_window.update_idletasks()
    settings_window.minsize(
        settings_window.winfo_reqwidth(), settings_window.winfo_reqheight()
    )
    settings_window.resizable(False, False)  # 让设置窗口不可缩放

    settings_window.protocol("WM_DELETE_WINDOW", close_settings_window)


def close_settings_window():
    global settings_window
    settings_window.destroy()
    settings_window = None


# 初始化全局变量
selected_ids = load_selected_ids()
settings_window = None
history_list = None

# 创建主窗口
root = tk.Tk()
root.title("随机抽号机")
root.attributes("-topmost", 1)
icon_path = os.path.join("icon.png")
icon_image = tk.PhotoImage(file=icon_path)
root.iconphoto(True, icon_image)

# 创建标签，用于显示抽取的学号
label = tk.Label(root, text="", font=("宋体", 50))
label.pack(pady=20, padx=20)

# 创建抽取学号按钮
button = tk.Button(root, text="抽取学号", font=("宋体", 12), command=draw_student_id)
button.pack(pady=10)

# 创建设置按钮
settings_button = tk.Button(root, text="设置", font=("宋体", 12), command=open_settings)
settings_button.pack(pady=10)

# 创建多行标签，用于显示信息，内容居中显示
info_label = tk.Label(
    root,
    text="v1.0 2024/06/07\nCopyright 2024 earthjasonlin\n保留所有权利。\n本程序仅作为交流学习使用",
    font=("宋体", 10),
    justify="center",
    anchor="center",
)
info_label.pack(pady=10, padx=20, fill="x")

# 自动调整窗口大小
root.update_idletasks()
root.minsize(root.winfo_reqwidth(), root.winfo_reqheight())
root.resizable(False, False)

root.mainloop()
