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
import win32con

# 设置全局参数，游戏框大小
xxyx_weight,xxyx_hight = 904, 1667

def find_and_move_to_image(image_path, region):
    """
    在指定区域内查找图片，并移动鼠标到该位置

    :param image_path: 需要查找的图片路径
    :param region: 指定搜索区域 (x, y, width, height)
    :return: 目标图片在 Windows 屏幕中的 (x, y) 坐标，如果未找到则返回 None
    """
    # 截取屏幕指定区域 (x, y, width, height)
    x, y, w, h = region
    screenshot = pyautogui.screenshot(region=(x, y, w, h))

    # 转换截图为 OpenCV 格式 (RGB → BGR)
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

    # 读取模板图片
    template = cv2.imread(image_path, cv2.IMREAD_COLOR)
    if template is None:
        print("无法加载模板图片，请检查路径")
        return None

    # 获取模板尺寸
    template_h, template_w, _ = template.shape

    # 使用 OpenCV 进行模板匹配
    result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)

    # 获取匹配位置
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

    # 设定匹配阈值
    threshold = 0.8  # 你可以调整这个值
    if max_val >= threshold:
        top_left = max_loc
        center_x = x + top_left[0] + template_w // 2
        center_y = y + top_left[1] + template_h // 2

        # 移动鼠标到目标位置
        pyautogui.click(center_x, center_y)

        print(f"找到目标，移动鼠标到: ({center_x}, {center_y})，并单击")
        return center_x, center_y
    else:
        print("未找到目标图片")
        return None

def find_and_click_sift(image_path, region):
    """
    在指定区域内查找目标图片，并移动鼠标到匹配位置

    :param image_path: 目标图片路径
    :param region: 搜索区域 (x, y, width, height)
    :return: 匹配位置 (x, y)，未找到返回 None
    """
    x, y, w, h = region

    # 截取屏幕特定区域
    screenshot = pyautogui.screenshot(region=(x, y, w, h))
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
    center_x = int(np.mean([kp2[m.trainIdx].pt[0] for m in good_matches])) + x
    center_y = int(np.mean([kp2[m.trainIdx].pt[1] for m in good_matches])) + y

    # 移动鼠标到目标位置
    pyautogui.click(center_x, center_y)

    print(f"找到目标，移动鼠标到: ({center_x}, {center_y})")
    return center_x, center_y


def find_and_move_sift(image_path, region):
    """
    在指定区域内查找目标图片，并移动鼠标到匹配位置

    :param image_path: 目标图片路径
    :param region: 搜索区域 (x, y, width, height)
    :return: 匹配位置 (x, y)，未找到返回 None
    """
    x, y, w, h = region

    # 截取屏幕特定区域
    screenshot = pyautogui.screenshot(region=(x, y, w, h))
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
    center_x = int(np.mean([kp2[m.trainIdx].pt[0] for m in good_matches])) + x
    center_y = int(np.mean([kp2[m.trainIdx].pt[1] for m in good_matches])) + y

    # 移动鼠标到目标位置
    # pyautogui.moveTo(center_x, center_y, duration=0.5)

    print(f"找到目标: ({center_x}, {center_y})")
    return center_x, center_y

def init_xxyx_inwindows():
    """
    从MuMu模拟器初始化小小英雄
    :return:
    """
    # Step1:初始化小小英雄
    windows = gw.getAllTitles()
    xxyx = None
    if gw.getWindowsWithTitle("MuMu模拟器12"):
        print("MuMu模拟器12已开启")
        xxyx = gw.getWindowsWithTitle("MuMu模拟器12")[0]
        xxyx.resizeTo(1600, 1000)
        # 设置窗口为顶置窗口
        win32gui.SetWindowPos(xxyx._hWnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        print("初始化游戏配置")
        find_and_click_sift(r"./img\mumu_xxyx_img.png", (xxyx.left, xxyx.top, xxyx.width, xxyx.height))
        xxyx.resizeTo(xxyx_weight,xxyx_hight)
        time.sleep(0.5)
        xxyx.moveTo(0, 0)
    else:
        print("MuMu模拟器12未开启，现在开始启动MuMu模拟器")
        # 打开Mumu模拟器
        subprocess.Popen(r"D:\Program Files\Netease\MuMu Player 12\shell\MuMuPlayer.exe")
        print("打开Mumu模拟器")
        time.sleep(0.5) ##提供程序响应时间
        while True:
            if gw.getWindowsWithTitle("MuMu模拟器12"):
                print("找到MuMu模拟器进程，开始调整大小至1600:1000")
                xxyx = gw.getWindowsWithTitle("MuMu模拟器12")[0]
                win32gui.SetWindowPos(xxyx._hWnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                # 设置窗口大小（宽, 高）
                xxyx.resizeTo(1600, 1000)
                # 设置窗口位置（左上角的 X, Y 坐标）
                xxyx.moveTo(0, 0)
                break
            else:
                print("没找到MuMu模拟器进程,3s后继续尝试")
                time.sleep(3)

        while True:
            find_result = find_and_click_sift(r"./img\mumu_xxyx_img.png", (0, 0, 1600, 1000))
            if find_result is not None:
                time.sleep(0.5)
                break
            else:
                time.sleep(6)
        # 移动窗口：
        xxyx.moveTo(0, 0)
        width, height = xxyx.width, xxyx.height
        return width, height


# 初始化菜单栏
def init_menu():
    while True:
        find_return_button = find_and_move_sift(r"./img\return_button.png",
                                               (0, 0, xxyx_weight, xxyx_hight))
        if find_return_button is not None:
            print("未在初始菜单栏，有返回按钮,至初始菜单栏")
            pyautogui.click(find_return_button)
            time.sleep(0.5)


        find_close_button = find_and_move_sift(r"./img\close_button.png",
                                      (0, 0,xxyx_weight,xxyx_hight))
        if find_close_button is not None:
            print("未在初始菜单栏，有关闭按钮,至初始菜单栏")
            pyautogui.click(find_close_button)
            time.sleep(0.5)
        else:
            print("以回归至初始菜单栏")
            break
    return True
def wujingshilian():
    init_menu()
    #     Part 2 开启无尽试炼
    click_fuben_button = find_and_click_sift(r"./img\fuben_button.png",
                                             (0, 0, xxyx_weight, xxyx_hight))
    time.sleep(0.3)
    click_teshufuben_button = find_and_click_sift(r"./img\teshufuben.png",
                                                  (0, 0, xxyx_weight, xxyx_hight))
    time.sleep(0.3)
    click_in_button = find_and_click_sift(r"./img\in_button.png",
                                          (0, 0, xxyx_weight, xxyx_hight))
    time.sleep(0.3)
    click_tiaozhan_button = find_and_click_sift(r"./img\tiaozhan.png",
                                                (0, 0, xxyx_weight, xxyx_hight))
    time.sleep(1.5)
    while True:
        find_re_tiaozhan = find_and_move_sift(r"./img\re_tiaozhan.png",
                                              (0, 0, xxyx_weight, xxyx_hight))
        if find_re_tiaozhan is not None:
            pyautogui.click(find_re_tiaozhan)
            time.sleep(3)
        else:
            time.sleep(2)

if __name__ == '__main__':
    # 初始化开启小小英雄。优化项：忽略图片大小进行匹配
    init_xxyx_inwindows()
    # pyautogui.moveTo(394, 45)
#     Part 1 初始化菜单栏
#     init_menu()
#     Part 2 开启无尽试炼
#     wujingshilian()
    # Part 3 组队副本
    # result1 = find_and_move_sift(r"./img\zuduifuben.png",(0, 0, xxyx_weight, xxyx_hight))
    # pyautogui.moveTo(442, 810)