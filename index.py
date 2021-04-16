# 获取所有活动窗口的标题
# ! /usr/bin/env python
#  -*- coding: utf-8 -*-
import time
import win32gui
import win32api
import win32con

"""
该类实现了:查找(定位)句柄信息，菜单信息
"""


class HandleMenu(object):

    def __init__(self, cls_name=None, title=None):

        self.handle = win32gui.FindWindow(cls_name, title)
        self.window_list = []

    def call_back(self, sub_handle, sub_handles):  # Edit 20de0
        """遍历子窗体"""

        title = win32gui.GetWindowText(sub_handle)
        cls_name = win32gui.GetClassName(sub_handle)
        print(title, '+', cls_name)

        position = win32gui.GetWindowRect(sub_handle)
        aim_point = round(position[0] + (position[2] - position[0]) / 2), round(
            position[1] + (position[3] - position[1]) / 2)
        win32api.SetCursorPos(aim_point)
        time.sleep(1)

        # 鼠标点击
        # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        # time.sleep(0.05)
        # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
        # time.sleep(0.05)
        # ComboBox - ---------
        # Edit - ---------
        if cls_name == 'ComboBox':
            win32gui.SendMessage(sub_handle, win32con.WM_SETTEXT, None, '902723')
            time.sleep(1)
        sub_handles.append({'cls_name': cls_name, 'title': title})

        return True

    def get_sub_handles(self):
        """通过父句柄获取子句柄"""

        sub_handles = []
        win32gui.EnumChildWindows(self.handle, self.call_back, sub_handles)
        print(sub_handles)
        return sub_handles

    def get_menu_text(self, menu, idx):
        import win32gui_struct
        mii, extra = win32gui_struct.EmptyMENUITEMINFO()  # 新建一个win32gui的空的结构体mii
        win32gui.GetMenuItemInfo(menu, idx, True, mii)  # 将子菜单内容获取到mii
        ftype, fstate, wid, hsubmenu, hbmpchecked, hbmpunchecked,
        dwitemdata, text, hbmpitem = win32gui_struct.UnpackMENUITEMINFO(mii)  # 解包mii
        return text

    def get_menu(self):
        """menu操作（记事本）"""

        menu = win32gui.GetMenu(self.handle)
        menu1 = win32gui.GetSubMenu(menu, 0)  # 第几个菜单 0-第一个
        cmd_ID = win32gui.GetMenuItemID(menu1, 3)  # 第几个子菜单
        win32gui.PostMessage(self.handle, win32con.WM_COMMAND, cmd_ID, 0)
        menu_text1 = [self.get_menu_text(menu, i) for i in range(5)]
        menu_text2 = [self.get_menu_text(menu1, i) for i in range(9)]

        print(menu_text1)
        print(menu_text2)

    def get_all_window_info(self, hwnd, nouse):

        # 去掉下面这句就所有都输出了，但是我不需要那么多
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
            # 设置为最前窗口
            win32gui.SetForegroundWindow(hwnd)
            # 获取某个句柄的标题和类名
            title = win32gui.GetWindowText(hwnd)
            cls_name = win32gui.GetClassName(hwnd)
            d = {'类名': cls_name, '标题': title}
            info = win32gui.GetWindowRect(hwnd)
            aim_point = round(info[0] + (info[2] - info[0]) / 2), round(info[1] + (info[3] - info[1]) / 2)
            win32api.SetCursorPos(aim_point)
            time.sleep(2)
            self.window_list.append(d)

    def get_all_windows(self):
        """获取所有活动窗口的类名、标题"""

        win32gui.EnumWindows(self.get_all_window_info, 0)

        return self.window_list


if __name__ == '__main__':
    # 1.通过父句柄获取子句柄
    # hm=HandleMenu(title='另存为')
    # hm.get_sub_handles()
    # 2.menu操作
    # hm=HandleMenu(title='aa - 记事本')
    # hm.get_menu()
    # 3.获取所有活动窗口的类名、标题
    hm = HandleMenu()
    hm.get_all_windows()

---
import win32api
import win32gui
import win32con
import win32print
import time

# 1 获取句柄
# 1.1 通过坐标获取窗口句柄
handle = win32gui.WindowFromPoint(win32api.GetCursorPos())  # (259, 185)
# 1.2 获取最前窗口句柄
handle = win32gui.GetForegroundWindow()
# 1.3 通过类名或查标题找窗口
handle = win32gui.FindWindow('cls_name', "title")
# 1.4 找子窗体
sub_handle = win32gui.FindWindowEx(handle, None, 'Edit', None)  # 子窗口类名叫“Edit”

# 句柄操作
title = win32gui.GetWindowText(handle)
cls_name = win32gui.GetClassName(handle)
print({'类名': cls_name, '标题': title})
# 获取窗口位置
info = win32gui.GetWindowRect(handle)
# 设置为最前窗口
win32gui.SetForegroundWindow(handle)

# 2.按键-看键盘码
# 获取鼠标当前位置的坐标
cursor_pos = win32api.GetCursorPos()
# 将鼠标移动到坐标处
win32api.SetCursorPos((200, 200))
# 回车
win32api.keybd_event(13, 0, win32con.KEYEVENTF_EXTENDEDKEY, 0)
win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)
# 左单键击
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
# 右键单击
win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP | win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
# 鼠标左键按下-放开
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
# 鼠标右键按下-放开
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
# TAB键
win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
# 快捷键Alt+F
win32api.keybd_event(18, 0, 0, 0)  # Alt
win32api.keybd_event(70, 0, 0, 0)  # F
win32api.keybd_event(70, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)

# 3.Message
win = win32gui.FindWindow('Notepad', None)
tid = win32gui.FindWindowEx(win, None, 'Edit', None)
# 输入文本
win32gui.SendMessage(tid, win32con.WM_SETTEXT, None, '你好hello word!')
# 确定
win32gui.SendMessage(handle, win32con.WM_COMMAND, 1, btnhld)
# 关闭窗口
win32gui.PostMessage(win32gui.FindWindow('Notepad', None), win32con.WM_CLOSE, 0, 0)
# 回车
win32gui.PostMessage(tid, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)  # 插入一个回车符,post 没有返回值，执行完马上返回
win32gui.PostMessage(tid, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
print("%x" % win)
print("%x" % tid)
# 选项框
res = win32api.MessageBox(None, "Hello Pywin32", "pywin32", win32con.MB_YESNOCANCEL)
print(res)
win32api.MessageBox(0, "Hello PYwin32", "MessageBox", win32con.MB_OK | win32con.MB_ICONWARNING)  # 加警告标志

# 4.截屏
from PIL import ImageGrab

# 利用PIL截屏
im = ImageGrab.grab()
im.save('aa.jpg')

# 5.文件的读写
import win32file, win32api, win32con
import os


def SimpleFileDemo():
    testName = os.path.join(win32api.GetTempPath(), "opt_win_file.txt")

    if os.path.exists(testName):
        os.unlink(testName)  # os.unlink() 方法用于删除文件,如果文件是一个目录则返回一个错误。
    # 写
    handle = win32file.CreateFile(testName,
                                  win32file.GENERIC_WRITE,
                                  0,
                                  None,
                                  win32con.CREATE_NEW,
                                  0,
                                  None)
    test_data = "Hello0there".encode("ascii")
    win32file.WriteFile(handle, test_data)
    handle.Close()
    # 读
    handle = win32file.CreateFile(testName, win32file.GENERIC_READ, 0, None, win32con.OPEN_EXISTING, 0, None)
    rc, data = win32file.ReadFile(handle, 1024)
    handle.Close()  # 此处也可使用win32file.CloseHandle(handle)来关闭句柄
    if data == test_data:
        print("Successfully wrote and read a file")
    else:
        raise Exception("Got different data back???")
    os.unlink(testName)


if __name__ == '__main__':
    SimpleFileDemo()
    # print(win32api.GetTempPath())  # 获取临时文件夹路径

# 6. ShellExecute
win32api.ShellExecute(None, "open", "C:Test.txt", None, None, SW_SHOWNORMAL)  # 打开C:Test.txt 文件
win32api.ShellExecute(None, "open", "http:#www.google.com", None, None, SW_SHOWNORMAL)  # 打开网页www.google.com
win32api.ShellExecute(None, "explore", "D:C++", None, None, SW_SHOWNORMAL)  # 打开目录D:C++
win32api.ShellExecute(None, "print", "C:Test.txt", None, None, SW_HIDE)  # 打印文件C:Test.txt
win32api.ShellExecute(None,
                      "open", "mailto:", None, None, SW_SHOWNORMAL)  # 打开邮箱
win32api.ShellExecute(None, "open", "calc.exe", None, None, SW_SHOWNORMAL)  # 调用计算器
win32api.ShellExecute(None, "open", "NOTEPAD.EXE", None, None, SW_SHOWNORMAL)  # 调用记事本
#  ShellExecute 不支持定向输出

# 打印--只能打印txt word excel 不能打印图片及pdf
from win32con import SW_HIDE, SW_SHOWNORMAL

for fn in ['2.txt', '3.txt']:
    print(fn)
    res = win32api.ShellExecute(0,  # 指定父窗口句柄

                                'print',  # 指定动作, 譬如: open、print、edit、explore、find

                                fn,  # 指定要打开的文件或程序

                                win32print.GetDefaultPrinter(),
                                # 给要打开的程序指定参数;GetDefaultPrinter　　取得默认打印机名称 <type 'str'>,GetDefaultPrinterW　　取得默认打印机名称 <type 'unicode'>

                                "./downloads/",  # 目录路径

                                SW_SHOWNORMAL)  # 打开选项,SW_HIDE = 0; {隐藏},SW_SHOWNORMAL = 1; {用最近的大小和位置显示, 激活}
    print(res)  # 返回值大于32表示执行成功,返回值小于32表示执行错误


# 打印 -pdf
def print_pdf(pdf_file_name):
    """
    静默打印pdf
    :param pdf_file_name:
    :return:
    """
    # GSPRINT_PATH = resource_path + 'GSPRINTgsprint'
    GHOSTSCRIPT_PATH = resource_path + 'GHOSTSCRIPTbingswin32c'  # gswin32c.exe
    currentprinter = config.printerName  # "printerName":"FUJI XEROX ApeosPort-VI C3370""
    currentprinter = win32print.GetDefaultPrinter()
    arg = '-dPrinted '
    '-dBATCH '
    '-dNOPAUSE '
    '-dNOSAFER '
    '-dFitPage '
    '-dNORANGEPAGESIZE '
    '-q '
    '-dNumCopies=1 '
    '-sDEVICE=mswinpr2 '
    '-sOutputFile="spool'
    + currentprinter + " " +
    pdf_file_name


log.info(arg)
win32api.ShellExecute(
    0,
    'open',
    GHOSTSCRIPT_PATH,
    arg,
    ".",
    0
)
# os.remove(pdf_file_name)