import pyautogui
import win32gui
import win32ui
import win32con


def capture_edge_window():
    def find_edge_window():
        def callback(hwnd, param):
            window_text = win32gui.GetWindowText(hwnd)
            if "Edge" in window_text:
                param.append(hwnd)
            return True

        hwnd_list = []
        win32gui.EnumWindows(callback, hwnd_list)
        if hwnd_list:
            print(f"找到 Edge 窗口，句柄: {hwnd_list[0]}")
            return hwnd_list[0]
        print("未找到 Edge 窗口")
        return 0

    edge_handle = find_edge_window()
    if edge_handle == 0:
        print("未找到 Edge 窗口，返回默认值")
        return None, 0, 0
    if not win32gui.IsWindow(edge_handle):
        print("找到的 Edge 窗口句柄无效")
        return None, 0, 0

    # 检查窗口是否被最小化或隐藏
    if win32gui.IsIconic(edge_handle) or not win32gui.IsWindowVisible(edge_handle):
        print("Edge 窗口被最小化或隐藏")
        return None, 0, 0

    try:
        # 获取窗口客户区坐标
        client_left, client_top, client_right, client_bottom = win32gui.GetClientRect(edge_handle)
        # 获取窗口左上角的屏幕坐标
        left, top = win32gui.ClientToScreen(edge_handle, (client_left, client_top))
        width = client_right - client_left
        height = client_bottom - client_top

        print(f"Edge 窗口坐标: left={left}, top={top}, width={width}, height={height}")

        hwndDC = win32gui.GetWindowDC(edge_handle)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)

        saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

        signedIntsArray = saveBitMap.GetBitmapBits(True)
        img = pyautogui.screenshot(region=(left, top, width, height))
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(edge_handle, hwndDC)

        return img, left, top
    except Exception as e:
        print(f"捕获 Edge 窗口截图时出错: {e}")
        return None, 0, 0


def find_image_on_screen(template_path, screenshot, left, top):
    try:
        region = (left, top, screenshot.width, screenshot.height)
        location = pyautogui.locateOnScreen(template_path, region=region, confidence=0.9)
        if location is not None:
            center_x = location.left + location.width // 2
            center_y = location.top + location.height // 2
            return (center_x, center_y)
        else:
            print("未找到匹配的图片")
            return None
    except pyautogui.ImageNotFoundException:
        print(f"在指定区域 {region} 内未找到模板图片 {template_path}")
        return None
    except Exception as e:
        print(f"查找模板图片时出错: {e}")
        return None


def main():
    def find_edge_window():
        def callback(hwnd, param):
            window_text = win32gui.GetWindowText(hwnd)
            if "Edge" in window_text:
                param.append(hwnd)
            return True

        hwnd_list = []
        win32gui.EnumWindows(callback, hwnd_list)
        if hwnd_list:
            return hwnd_list[0]
        return 0

    edge_handle = find_edge_window()
    if edge_handle == 0:
        print("未找到 Edge 窗口")
        return

    screenshot, left, top = capture_edge_window()
    if screenshot is None:
        return

    template_path = 'C://Users//Administrator//Desktop//pyImg//edge//baidu.png'

    location = find_image_on_screen(template_path, screenshot, left, top)
    if location is not None:
        global_x = left + location[0]
        global_y = top + location[1]
        pyautogui.moveTo(global_x, global_y)
        pyautogui.click()


if __name__ == "__main__":
    main()
