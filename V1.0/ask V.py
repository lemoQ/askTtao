import os
import tkinter as tk
from tkinter import ttk, messagebox

import win32api
import win32gui
import win32con
import time
import logging
import pyautogui
import win32process
import psutil

# 配置日志记录
logging.basicConfig(filename='log.txt', level=logging.INFO,
                    format='%(asctime)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

# 定义全局变量 log_text
log_text = None

def minimize_all_windows():
    """最小化所有可见窗口"""
    def enum_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
    win32gui.EnumWindows(enum_callback, None)

def log_message(message):
    """记录日志并显示在日志窗口中"""
    global log_text
    logging.info(message)
    if log_text:
        log_text.insert(tk.END, f"{message}\n")
        log_text.see(tk.END)


def get_process_name(hwnd):
    """获取窗口对应的进程名"""
    try:
        _, process_id = win32process.GetWindowThreadProcessId(hwnd)
        return psutil.Process(process_id).name().lower()
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        return ""


# 改进后的窗口查找函数
def get_window_handles():
    handles = []

    def callback(hwnd, ctx):
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd):
            name = win32gui.GetWindowText(hwnd)
            process_name = get_process_name(hwnd)

            # 同时检查窗口标题和进程名
            if ("线" in name.lower() or "py" in name.lower()) or \
                    ("edge" in process_name or "game" in process_name):  # 添加游戏进程名关键字
                handles.append((hwnd, name))

    win32gui.EnumWindows(callback, None)
    return handles

def get_target(name):
    window_list = get_window_handles()
    target_hwnd = None
    for hwnd, title in window_list:
        if name in title:
            target_hwnd = hwnd
            log_message(f"找到包含 '{name}' 的窗口: {title}")
            break
    return target_hwnd

def update_treeview():
    try:
        # 清空当前的 Treeview 内容
        for item in tree.get_children():
            tree.delete(item)
        # 获取窗口句柄和标题
        window_list = get_window_handles()
        for hwnd, title in window_list:
            # 模拟其他列的数据
            team_number = "1"
            is_captain = "是"
            is_offset = "否"
            auto_battle_mode = "开启"
            tree.insert("", "end", values=(hwnd, title, team_number, is_captain, is_offset, auto_battle_mode))
        log_message("成功更新窗口句柄信息")
    except Exception as e:
        log_message(f"更新窗口句柄信息时出错: {e}")
    # 每隔 5 秒更新一次
    root.after(5000, update_treeview)

def get_py():
    target_hwnd = get_target("py")
    if target_hwnd:
        try:

            # 展开要操作的窗口
            win32gui.ShowWindow(target_hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(target_hwnd)
            log_message(f"展开pycharm成功")
        except Exception as e:
            log_message(f"操作包含 'pycharm' 的窗口句柄的进程时出错: {e}")
        finally:
            # 显示pycharm窗口
            win32gui.ShowWindow(win32gui.GetForegroundWindow(), win32con.SW_RESTORE)
    else:
        log_message("未找到标题包含 'pycharm' 的窗口")
def test_button_clicked():
    log_message("测试按钮 - 使用 SendInput 模拟输入")
    target_hwnd = get_target("桂圆")

    if target_hwnd:
        try:
            # 可选：激活窗口（可能导致窗口显示）
            # activate_window(target_hwnd)

            # 确保窗口有焦点（这一步可能导致窗口显示，根据需要调整）
            win32gui.SetForegroundWindow(target_hwnd)

            # 模拟 Alt+E
            log_message("发送 Alt+E")
            pyautogui.hotkey('alt', 'e')
            time.sleep(0.5)
            log_message("发送 Ctrl+B")
            pyautogui.hotkey('ctrl', 'b')
            time.sleep(0.5)
            log_message("已向窗口发送 Ctrl+B 和 Alt+E 组合键")
        except Exception as e:
            log_message(f"操作窗口时出错: {e}")
    else:
        log_message("未找到目标窗口")


def yjzd():
    log_message("一键组队")
    target_hwnd = get_target("桂圆")
    win32gui.SetForegroundWindow(target_hwnd)
    # 图像识别和点击功能
    image_path = "D:/CTTQ/lshPy/askTaoPy/V1.0/png/zd.png"
    if not os.path.exists(image_path):
        log_message(f"错误：图像文件不存在 - {image_path}")
        return

    location = pyautogui.locateOnScreen(image_path, confidence=0.8)
    log_message("点击组队")
    # 确保窗口有焦点（这一步可能导致窗口显示，根据需要调整）
    win32gui.SetForegroundWindow(target_hwnd)
    pyautogui.move(location)
    time.sleep(1)
    pyautogui.click(location)
    log_message("组队完成")


def jsdw():
    log_message("解散队伍")
    target_hwnd = get_target("桂圆")
    win32gui.SetForegroundWindow(target_hwnd)
    pyautogui.hotkey('ctrl', 'b')
    pyautogui.hotkey('alt', 't')
    time.sleep(0.5)
    # 图像识别和点击功能
    image_path = "D:/CTTQ/lshPy/askTaoPy/V1.0/png/ld.png"
    if not os.path.exists(image_path):
        log_message(f"错误：图像文件不存在 - {image_path}")
        return

    location = pyautogui.locateOnScreen(image_path, confidence=0.8)
    log_message("点击离队")
    # 确保窗口有焦点（这一步可能导致窗口显示，根据需要调整）
    win32gui.SetForegroundWindow(target_hwnd)
    pyautogui.moveTo(location)
    time.sleep(1)
    pyautogui.click(location)
    log_message("离队完成")
    # 这里可以添加保存窗口的具体逻辑
    # messagebox.showinfo("提示", "窗口已保存")


def clear_data():
    log_message("清空数据")
    # 这里可以添加清空数据的具体逻辑
    for item in tree.get_children():
        tree.delete(item)
    messagebox.showinfo("提示", "数据已清空")


def delete_selected_row():
    selected_item = tree.selection()
    if selected_item:
        log_message("删除选中行")
        tree.delete(selected_item)
        messagebox.showinfo("提示", "选中行已删除")
    else:
        log_message("未选中任何行，无法删除")
        messagebox.showwarning("警告", "请先选中要删除的行")


def clear_log():
    log_message("清空日志")
    # 清空日志文件
    with open('log.txt', 'w') as f:
        f.write('')
    # 清空日志显示区域
    log_text.delete(1.0, tk.END)
    messagebox.showinfo("提示", "日志已清空")


def get_authorization_code():
    log_message("获取授权码")
    # 这里可以添加获取授权码的具体逻辑
    messagebox.showinfo("提示", "授权码获取功能待实现")


def show_version_update_content():
    log_message("版本更新内容")
    # 这里可以添加显示版本更新内容的具体逻辑
    messagebox.showinfo("提示", "版本更新内容待实现")


def start_auto_elder():
    log_message("自动长老")
    # 这里可以添加自动长老的具体逻辑
    messagebox.showinfo("提示", "自动长老功能启动")


def start_auto_elite():
    log_message("自动精英")
    # 这里可以添加自动精英的具体逻辑
    messagebox.showinfo("提示", "自动精英功能启动")


def start_auto_normal():
    log_message("自动普通")
    # 这里可以添加自动普通的具体逻辑
    messagebox.showinfo("提示", "自动普通功能启动")


def start_auto_five_pulse():
    log_message("自动五脉")
    # 这里可以添加自动五脉的具体逻辑
    messagebox.showinfo("提示", "自动五脉功能启动")


def start_auto_five_pulse_2():
    log_message("自动五脉2")
    # 这里可以添加自动五脉2的具体逻辑
    messagebox.showinfo("提示", "自动五脉2功能启动")


def start_repeat_brushing():
    log_message("刷道模式: 重复刷道")
    # 这里可以添加重复刷道的具体逻辑
    messagebox.showinfo("提示", "重复刷道模式启动")


def start_brushing():
    log_message("开始刷道")
    # 这里可以添加开始刷道的具体逻辑
    messagebox.showinfo("提示", "刷道开始")


def start_reward_activity():
    log_message("活动: 悬赏1星")
    # 这里可以添加悬赏1星活动的具体逻辑
    messagebox.showinfo("提示", "悬赏1星活动启动")


def start_activity():
    log_message("开始活动")
    # 这里可以添加开始活动的具体逻辑
    messagebox.showinfo("提示", "活动开始")


def start_auto_immortal():
    log_message("自动仙人")
    # 这里可以添加自动仙人的具体逻辑
    messagebox.showinfo("提示", "自动仙人功能启动")


def start_auto_mountain_repair():
    log_message("自动跑修山")
    # 这里可以添加自动跑修山的具体逻辑
    messagebox.showinfo("提示", "自动跑修山功能启动")


def start_auto_nine_yao():
    log_message("自动跑九耀")
    # 这里可以添加自动跑九耀的具体逻辑
    messagebox.showinfo("提示", "自动跑九耀功能启动")


def start_big_fly_nine_yao():
    log_message("大飞九耀")
    # 这里可以添加大飞九耀的具体逻辑
    messagebox.showinfo("提示", "大飞九耀功能启动")


def start_auto_immortal_search():
    log_message("自动跑寻仙")
    # 这里可以添加自动跑寻仙的具体逻辑
    messagebox.showinfo("提示", "自动跑寻仙功能启动")


def start_auto_ten_absolutes():
    log_message("自动跑十绝")
    # 这里可以添加自动跑十绝的具体逻辑
    messagebox.showinfo("提示", "自动跑十绝功能启动")


def start_gang_boss():
    selected_boss = boss_var.get()
    log_message(f"开始boss - {selected_boss}")
    # 这里可以添加开始帮派boss的具体逻辑
    messagebox.showinfo("提示", f"{selected_boss} 帮派boss活动启动")


def start_next_day_auto_task():
    log_message("次日自动任务")
    # 这里可以添加次日自动任务的具体逻辑
    messagebox.showinfo("提示", "次日自动任务启动")


def start_auto_task():
    log_message("自动任务")
    # 这里可以添加自动任务的具体逻辑
    messagebox.showinfo("提示", "自动任务启动")


def start_auto_welfare():
    log_message("自动领取福利")
    # 这里可以添加自动领取福利的具体逻辑
    messagebox.showinfo("提示", "自动领取福利功能启动")


def start_auto_life_record():
    log_message("自动浮生录")
    # 这里可以添加自动浮生录的具体逻辑
    messagebox.showinfo("提示", "自动浮生录功能启动")


def start_auto_round_supplement():
    log_message("自动补回合")
    # 这里可以添加自动补回合的具体逻辑
    messagebox.showinfo("提示", "自动补回合功能启动")


def start_auto_single_lottery():
    log_message("自动单次求签")
    # 这里可以添加自动单次求签的具体逻辑
    messagebox.showinfo("提示", "自动单次求签功能启动")


def start_auto_timed_lottery():
    log_message("自动定时求签")
    # 这里可以添加自动定时求签的具体逻辑
    messagebox.showinfo("提示", "自动定时求签功能启动")


root = tk.Tk()
root.title("菜狗V1.0")

# 顶部表格区域
table_frame = tk.Frame(root)
table_frame.pack(pady=5, fill=tk.X)

columns = ("句柄", "游戏名称", "队伍编号", "是否队长", "是否偏移", "自动战斗模式")
tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=10)
for col in columns:
    tree.heading(col, text=col)
tree.pack(fill=tk.BOTH, expand=True)

# 开始更新 Treeview
update_treeview()

# 基础功能区域
base_frame = tk.Frame(root)
base_frame.pack(pady=5, fill=tk.X)

# 掉线监测复选框
check_var1 = tk.BooleanVar()
tk.Checkbutton(base_frame, text="掉线监测", variable=check_var1).pack(side=tk.LEFT, padx=2)

# 自动补状态及单选框
auto_state_var = tk.StringVar()
tk.Checkbutton(base_frame, text="自动补状态", variable=check_var1).pack(side=tk.LEFT, padx=2)
tk.Radiobutton(base_frame, text="每次补自动", variable=auto_state_var, value="each").pack(side=tk.LEFT, padx=2)
tk.Radiobutton(base_frame, text="每轮补自动", variable=auto_state_var, value="each_round").pack(side=tk.LEFT, padx=2)
tk.Radiobutton(base_frame, text="不补自动补状态", variable=auto_state_var, value="none").pack(side=tk.LEFT, padx=2)

# 其他复选框
check_vars = []
for txt in ["使用特八"]:
    var = tk.BooleanVar()
    check_vars.append(var)
    tk.Checkbutton(base_frame, text=txt, variable=var).pack(side=tk.LEFT, padx=2)

# 操作按钮区域
button_frame = tk.Frame(root)
button_frame.pack(pady=5, fill=tk.X)
button_functions = {
    "一键组队": yjzd,
    "解散队伍": jsdw,
    "清空数据": clear_data,
    "删除选中行": delete_selected_row,
    "清空日志": clear_log,
    "测试按钮": test_button_clicked,
    "获取授权码": get_authorization_code,
    "版本更新内容": show_version_update_content
}
for btn_txt in button_functions.keys():
    tk.Button(button_frame, text=btn_txt, width=10, command=button_functions[btn_txt]).pack(side=tk.LEFT, padx=2)

# 队长功能区域（第一组）
captain_frame1 = tk.Frame(root)
captain_frame1.pack(pady=5, fill=tk.X)
captain_functions_1 = {
    "自动长老": start_auto_elder,
    "自动精英": start_auto_elite,
    "自动普通": start_auto_normal,
    "自动五脉": start_auto_five_pulse,
    "自动五脉2": start_auto_five_pulse_2,
    "刷道模式: 重复刷道": start_repeat_brushing,
    "开始刷道": start_brushing,
    "活动: 悬赏1星": start_reward_activity,
    "开始活动": start_activity
}
for btn_txt in captain_functions_1.keys():
    tk.Button(captain_frame1, text=btn_txt, width=12, command=captain_functions_1[btn_txt]).pack(side=tk.LEFT, padx=2)

# 队长功能区域（第二组）
captain_frame2 = tk.Frame(root)
captain_frame2.pack(pady=5, fill=tk.X)
team_var = tk.StringVar()
captain_functions_2 = {
    "自动仙人": start_auto_immortal,
    "自动跑修山": start_auto_mountain_repair,
    "自动跑九耀": start_auto_nine_yao,
    "大飞九耀": start_big_fly_nine_yao,
    "自动跑寻仙": start_auto_immortal_search,
    "自动跑十绝": start_auto_ten_absolutes
}
for btn_txt in captain_functions_2.keys():
    tk.Button(captain_frame2, text=btn_txt, width=12, command=captain_functions_2[btn_txt]).pack(side=tk.LEFT, padx=2)
for btn_txt in ["队伍1", "队伍2", "所有队伍"]:
    tk.Radiobutton(captain_frame2, text=btn_txt, variable=team_var, value=btn_txt).pack(side=tk.LEFT, padx=2)

# 帮派boss区域
gang_frame = tk.Frame(root)
gang_frame.pack(pady=5, fill=tk.X)
boss_var = tk.StringVar()
tk.OptionMenu(gang_frame, boss_var, "1星", "2星").pack(side=tk.LEFT, padx=2)
tk.Button(gang_frame, text="开始boss", width=10, command=start_gang_boss).pack(side=tk.LEFT, padx=2)

# 个人功能区域
personal_frame = tk.Frame(root)
personal_frame.pack(pady=5, fill=tk.X)
personal_functions = {
    "次日自动任务": start_next_day_auto_task,
    "自动任务": start_auto_task,
    "自动领取福利": start_auto_welfare,
    "自动浮生录": start_auto_life_record,
    "自动补回合": start_auto_round_supplement,
    "自动单次求签": start_auto_single_lottery,
    "自动定时求签": start_auto_timed_lottery
}
for btn_txt in personal_functions.keys():
    tk.Button(personal_frame, text=btn_txt, width=12, command=personal_functions[btn_txt]).pack(side=tk.LEFT, padx=2)

# 日志显示区域
log_frame = tk.Frame(root)
log_frame.pack(pady=5, fill=tk.BOTH, expand=True)

log_text = tk.Text(log_frame)
log_text.pack(fill=tk.BOTH, expand=True)

root.mainloop()

