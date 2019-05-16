import win32gui
import win32api
import win32con
import time
import msvcrt

classname = "Win32Window0"
titlename = "阴阳师-网易游戏"
# 获取句柄
hwnd = win32gui.FindWindow(classname, titlename)

# 获取窗口左上角和右下角坐标

print("目标窗口大小:" + str(win32gui.GetWindowRect(hwnd)))

def click_position(hwd, x_position, y_position):
    """
    鼠标左键点击指定坐标
    :param hwd:
    :param x_position:
    :param y_position:
    :param sleep:
    :return:
    """
    # 将两个16位的值连接成一个32位的地址坐标

    long_position = win32api.MAKELONG(x_position, y_position)
    # long_position = 0x000D09D0
    # win32api.SendMessage(hwnd, win32con.MOUSEEVENTF_LEFTDOWN, win32con.MOUSEEVENTF_LEFTUP, long_position)
    # print("点击位置为：" + str(long_position))
    # win32gui.PostMessage(hwd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
    # 点击左键
    win32api.PostMessage(hwd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position)
    win32api.PostMessage(hwd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position)
    time.sleep(0.1)


# while True:
#     if ord(msvcrt.getch()) in [68, 100]:
#         break
#     print("Press 'D' to exit...")
while True:
    right, top, left, bottom = win32gui.GetWindowRect(hwnd)
    # click_position(hwnd, int((left - right) * 0.873), int((bottom - top) * 0.77))  # 组队开始78
    # time.sleep(0.4)
    # click_position(hwnd, int((left - right) * 0.873), int((bottom - top) * 0.77))  # 组队开始
    # time.sleep(0.2)
    click_position(hwnd, int((left - right) * 0.11), int((bottom - top) * 0.363))  # 左边组队开始
    time.sleep(0.5)
    click_position(hwnd, int((left - right) * 0.928), int((bottom - top) * 0.705))  # 组队开始
    time.sleep(0.5)
    # click_position(hwnd, int((left - right) * 0.5), int((bottom - top) * 0.15))  # 点怪0.73，0.282
    click_position(hwnd, int((left - right) * 0.73), int((bottom - top) * 0.2))  # 点怪0.73，0.282
    time.sleep(0.5)
