import pyautogui
import pygetwindow as gw
import time

# 获取Chrome窗口
chrome_windows = [window for window in gw.getWindowsWithTitle('Chrome') if 'Chrome' in window.title]

if not chrome_windows:
    print("未找到Chrome窗口")
else:
    chrome_window = chrome_windows[0]
    chrome_window.activate()  # 激活Chrome窗口

    # 等待几秒以确保窗口已激活
    #time.sleep(2)

    # 截图并查找按钮位置
    button_location = pyautogui.locateOnScreen('D:\\小软件\\TouTiao\\Pictures\\发稿100.png')

    # 点击按钮
    if button_location:
        pyautogui.click(button_location)
    else:
        print("未找到按钮")
