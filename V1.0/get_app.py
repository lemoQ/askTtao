import win32gui

def get_window_handles():
    """获取所有可见窗口的句柄和标题"""
    handles = []
    def enum_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            handles.append((hwnd, win32gui.GetWindowText(hwnd)))
    win32gui.EnumWindows(enum_callback, None)
    return handles

if __name__ == "__main__":
    window_list = get_window_handles()
    print("句柄\t\t\t标题")
    print("-" * 50)
    for hwnd, title in window_list:
        # 若只要特定标题（如包含“十九线”），可添加条件：if "十九线" in title:
        print(f"{hwnd}\t{title}")