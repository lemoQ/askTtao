import cv2
import numpy as np
import pyautogui
from PIL import ImageGrab
import win32gui


def find_image_position_in_window(hwnd, image_path, threshold=0.8):
    """
    在指定窗口中查找图像的位置

    参数:
    hwnd: 窗口句柄
    image_path: 要查找的图像文件路径
    threshold: 匹配阈值，范围0-1，值越高匹配越严格，默认0.8

    返回:
    如果找到匹配图像，返回其中心坐标(x, y)和相似度值
    如果未找到，返回None, None, None
    """
    try:
        # 获取窗口位置和大小
        window_rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = window_rect
        width = right - left
        height = bottom - top

        # 如果窗口太小，可能已最小化或不可用
        if width <= 10 or height <= 10:
            return None, None, None

        # 截取窗口内容
        screenshot = ImageGrab.grab(window_rect)
        screenshot_np = np.array(screenshot)
        screenshot_cv = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2BGR)

        # 读取模板图像
        template = cv2.imread(image_path, cv2.IMREAD_COLOR)
        if template is None:
            raise FileNotFoundError(f"无法加载图像: {image_path}")

        # 确保模板图像尺寸小于或等于屏幕截图
        if template.shape[0] > screenshot_cv.shape[0] or template.shape[1] > screenshot_cv.shape[1]:
            return None, None, None

        # 执行模板匹配
        result = cv2.matchTemplate(screenshot_cv, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

        # 如果相似度达到阈值，计算匹配区域的中心坐标
        if max_val >= threshold:
            h, w, _ = template.shape
            center_x = max_loc[0] + w // 2
            center_y = max_loc[1] + h // 2

            # 将坐标转换为屏幕坐标
            screen_x = left + center_x
            screen_y = top + center_y

            return screen_x, screen_y, max_val

        return None, None, None

    except Exception as e:
        print(f"图像识别过程中出错: {e}")
        return None, None, None