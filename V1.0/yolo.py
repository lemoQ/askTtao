import pyautogui
import cv2
import numpy as np
import win32gui
import win32ui
import win32con
from ultralytics import YOLO


def capture_edge_window():
    def find_edge_window():
        def callback(hwnd, param):
            window_text = win32gui.GetWindowText(hwnd)
            if "Edge" in window_text:  # 模糊匹配，只要窗口标题包含 "Edge" 即可
                param.append(hwnd)
            return True

        hwnd_list = []
        win32gui.EnumWindows(callback, hwnd_list)
        if hwnd_list:
            return hwnd_list[0]
        return 0

    try:
        # 获取 Edge 窗口句柄
        edge_handle = find_edge_window()
        if edge_handle == 0:
            print("未找到 Edge 窗口")
            return None

        # 获取窗口位置和大小
        left, top, right, bottom = win32gui.GetWindowRect(edge_handle)
        width = right - left
        height = bottom - top

        # 获取窗口设备上下文
        hwndDC = win32gui.GetWindowDC(edge_handle)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        # 创建位图对象
        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
        saveDC.SelectObject(saveBitMap)

        # 拷贝屏幕内容到位图
        saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)

        # 转换为 numpy 数组
        signedIntsArray = saveBitMap.GetBitmapBits(True)
        img = np.frombuffer(signedIntsArray, dtype='uint8')
        img.shape = (height, width, 4)

        # 释放资源
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(edge_handle, hwndDC)

        return cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    except Exception as e:
        print(f"捕获 Edge 窗口截图时出错: {e}")
        return None


def find_image_on_screen(model, screenshot):
    try:
        # 使用 YOLOv8 模型进行目标检测
        results = model(screenshot)

        if len(results[0].boxes) == 0:
            print("未找到匹配的图片")
            return None

        # 获取第一个检测到的目标的边界框
        box = results[0].boxes[0]
        x1, y1, x2, y2 = box.xyxy[0].cpu().numpy().astype(int)

        # 计算中心点位置
        center_x = (x1 + x2) // 2
        center_y = (y1 + y2) // 2

        return (center_x, center_y)
    except Exception as e:
        print(f"查找模板图片时出错: {e}")
        return None


def main():
    def find_edge_window():
        def callback(hwnd, param):
            window_text = win32gui.GetWindowText(hwnd)
            if "Edge" in window_text:  # 模糊匹配，只要窗口标题包含 "Edge" 即可
                param.append(hwnd)
            return True

        hwnd_list = []
        win32gui.EnumWindows(callback, hwnd_list)
        if hwnd_list:
            return hwnd_list[0]
        return 0

    # 获取 Edge 窗口句柄和位置
    edge_handle = find_edge_window()
    if edge_handle == 0:
        print("未找到 Edge 窗口")
        return
    left, top, _, _ = win32gui.GetWindowRect(edge_handle)

    # 捕获 Edge 窗口截图
    screenshot = capture_edge_window()
    if screenshot is None:
        return

    # 加载 YOLOv8 模型
    model = YOLO('yolov8n.pt')

    # 在截图中查找模板图片
    location = find_image_on_screen(model, screenshot)
    if location is not None:
        # 计算全局坐标
        global_x = left + location[0]
        global_y = top + location[1]

        # 移动鼠标到该位置
        pyautogui.moveTo(global_x, global_y)


if __name__ == "__main__":
    main()

