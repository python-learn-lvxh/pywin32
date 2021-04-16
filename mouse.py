# -*- coding: UTF-8 -*-
# Python version: 3.4.0
# 参考 https://blog.csdn.net/andyliulin/article/details/78355448?utm_term=python%E5%86%99%E6%8C%82%E6%9C%BA%E8%84%9A%E6%9C%AC&utm_medium=distribute.pc_aggpage_search_result.none-task-blog-2~all~sobaiduweb~default-1-78355448&spm=3001.4430
import win32api


def LeftClick(x, y):  # 鼠标左键点击屏幕上的坐标(x, y)
    win32api.SetCursorPos((x, y))  # 鼠标定位到坐标(x, y)
    # 注意：不同的屏幕分辨率会影响到鼠标的定位，有需求的请用百分比换算
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)  # 鼠标左键按下
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)  # 鼠标左键弹起
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN + win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)    # 测试


def PressOnce(x):  # 模拟键盘输入一个按键的值，键码: x
    win32api.keybd_event(x, 0, 0, 0)


# 测试
LeftClick(30, 30)  # 我的电脑？
PressOnce(13)  # Enter
PressOnce(9)  # TAB