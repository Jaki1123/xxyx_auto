import subprocess
import time
# 窗口
import pygetwindow as gw
# 鼠标自动化
import pyautogui
# 图片匹配
import cv2
import numpy as np
import win32gui
import win32api
import win32con
import win32ui
from PIL import Image
import ctypes
from ctypes.wintypes import RECT


# 获取某个窗口的截图
def capture_background_window(window_title):
    """
    对后台窗口截图（即使窗口未激活/最小化）

    :param window_title: 需要截图的窗口标题
    :return: 截取的 PIL 图片
    """
    # 获取窗口句柄
    windows = gw.getWindowsWithTitle(window_title)
    if not windows:
        print(f"未找到窗口: {window_title}")
        return None

    hwnd = windows[0]._hWnd  # 窗口句柄

    # 获取窗口尺寸
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top

    # 获取窗口设备上下文（DC）
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    # 创建兼容的位图（用于存储窗口图像）
    save_bitmap = win32ui.CreateBitmap()
    save_bitmap.CreateCompatibleBitmap(mfc_dc, width, height)

    # 复制窗口内容到内存中
    save_dc.SelectObject(save_bitmap)
    save_dc.BitBlt((0, 0), (width, height), mfc_dc, (0, 0), win32con.SRCCOPY)

    # 将位图转换为 PIL 图像
    bmp_info = save_bitmap.GetInfo()
    bmp_str = save_bitmap.GetBitmapBits(True)
    img = Image.frombuffer("RGB", (bmp_info["bmWidth"], bmp_info["bmHeight"]), bmp_str, "raw", "BGRX", 0, 1)

    # 释放资源
    win32gui.DeleteObject(save_bitmap.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)
    print("以获取到窗口截图")
    print(img)
    return img


def find_picture_bysift(image_path, region):
    """
    在指定区域内查找目标图片，并移动鼠标到匹配位置

    :param image_path: 目标图片路径
    :param region: 搜索区域 (x, y, width, height)
    :return: 匹配位置 (x, y)，未找到返回 None
    """
    x, y, w, h = region

    # 截取屏幕特定区域
    screenshot = capture_background_window("MuMu模拟器12")
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    # 读取目标图片
    template = cv2.imread(image_path, cv2.IMREAD_COLOR)
    if template is None:
        print("无法加载目标图片，请检查路径")
        return None

    # 初始化 SIFT 关键点检测器
    sift = cv2.SIFT_create()

    # 计算关键点 & 描述符
    kp1, des1 = sift.detectAndCompute(template, None)
    kp2, des2 = sift.detectAndCompute(screenshot, None)

    if des1 is None or des2 is None:
        print("未检测到足够的关键点")
        return None

    # 使用 FLANN 进行特征匹配
    FLANN_INDEX_KDTREE = 1
    index_params = dict(algorithm=FLANN_INDEX_KDTREE, trees=5)
    search_params = dict(checks=50)

    flann = cv2.FlannBasedMatcher(index_params, search_params)
    matches = flann.knnMatch(des1, des2, k=2)

    # Lowe’s ratio test 过滤误匹配
    good_matches = [m for m, n in matches if m.distance < 0.7 * n.distance]

    if len(good_matches) < 10:  # 至少 10 个匹配点
        print("匹配点太少，未找到目标")
        return None

    # 计算匹配点的重心（提高计算精度）
    center_x = int(np.mean([kp2[m.trainIdx].pt[0] for m in good_matches]))
    center_y = int(np.mean([kp2[m.trainIdx].pt[1] for m in good_matches]))

    # 移动鼠标到目标位置
    # pyautogui.moveTo(center_x, center_y, duration=0.5)

    print(f"找到目标: ({center_x}, {center_y})")
    return center_x, center_y


if __name__ == '__main__':
    print("MuMu模拟器12未开启，现在开始启动MuMu模拟器")
    # 打开Mumu模拟器
    # subprocess.Popen(r"D:\Program Files\Netease\MuMu Player 12\shell\MuMuPlayer.exe")
    # time.sleep(5)
    xxyx = gw.getWindowsWithTitle("MuMu模拟器12")
    hwnd = xxyx[0]._hWnd  # 获取窗口句柄

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top

    center_x, center_y = find_picture_bysift(r"./img\mumu_xxyx_img.png",(left, top, right, bottom))
    # print("找到图片对应窗口位置："+center_x+"，"+center_y)

    click_x = left + center_x
    click_y = top + center_y
    pyautogui.moveTo(1504, 1288)

    print(click_x | (click_y << 16))
    print(click_y | (click_x << 16))

    # # 发送鼠标点击消息（后台点击）
    # win32gui.SetWindowPos(xxyx._hWnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    # win32gui.SetForegroundWindow(hwnd)  # 激活窗口
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, (click_y << 16) | click_x )
    win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, None,(click_y << 16) | click_x)

    # qweqwe = gw.getAllWindows()
    # for qwe in qweqwe:
    #     print(qwe.title)
    # win = gw.getWindowsWithTitle("pythonProject2 – houtaimoshi.py")
    # hwnd = win[0]._hWnd  # 获取窗口句柄
    # win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, (5 << 16) | 5)
    # win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, None, (5 << 16) | 5)

def click_with_SendInput(x, y):
    """
    使用 SendInput 模拟真实鼠标点击（适用于 DirectX 游戏）
    """
    ctypes.windll.user32.SetCursorPos(x, y)  # 移动鼠标
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)  # 按下鼠标
    time.sleep(0.05)  # 短暂停顿，模拟真实点击
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # 释放鼠标

# 示例：点击屏幕上的 (500, 300)
click_with_SendInput(click_x, click_y)