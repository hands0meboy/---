import os
import time
import json
import shutil
import threading
import subprocess
import wx
import wx.adv
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

CONFIG_FILE = 'config.json'
ICON_FILE = 'ico.ico'  # 确保图标文件与脚本在同一目录下
APP_TITLE = "autoZipper"

class Watcher:
    def __init__(self, folder_to_watch, extraction_path):
        self.folder_to_watch = folder_to_watch
        self.extraction_path = extraction_path
        self.observer = Observer()

    def run(self):
        event_handler = Handler(self.extraction_path)
        self.observer.schedule(event_handler, self.folder_to_watch, recursive=False)
        self.observer.start()
        self.observer.join()  # 确保线程保持运行状态

    def stop(self):
        self.observer.stop()
        self.observer.join()

class Handler(FileSystemEventHandler):
    def __init__(self, extraction_path):
        self.extraction_path = extraction_path

    def on_created(self, event):
        if not event.is_directory and event.event_type == 'created':
            file_path = event.src_path
            if file_path.endswith(('.zip', '.rar', '.7z', '.tar.gz', '.tar')):
                if self.is_download_complete(file_path):
                    try:
                        self.extract_file(file_path, self.extraction_path)
                    except Exception as e:
                        pass

    def is_download_complete(self, file_path):
        previous_size = -1
        while True:
            current_size = os.path.getsize(file_path)
            if current_size == previous_size:
                return True
            previous_size = current_size
            time.sleep(1)
            
    def extract_file(self, file_path, extraction_path):
        try:
            # 确保路径中没有特殊字符并使用绝对路径
            file_path = os.path.abspath(file_path)
            extraction_path = os.path.abspath(extraction_path)

            # 检查文件和文件夹是否存在
            if not os.path.isfile(file_path):
                return
            if not os.path.isdir(extraction_path):
                os.makedirs(extraction_path)

            # 调用 Bandizip 命令行工具 bz.exe 解压文件
            bz_path = "bz.exe"  # 确保此路径是正确的
            command = f'bz x -o:"{extraction_path}" -target:auto "{file_path}"'
            result = subprocess.run(command, shell=True, capture_output=True, text=True, encoding='utf-8')

        except Exception as e:
            pass

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super(MyFrame, self).__init__(parent, title=title, size=(500, 160), style=wx.DEFAULT_FRAME_STYLE & ~(wx.MAXIMIZE_BOX | wx.RESIZE_BORDER))
        self.panel = wx.Panel(self)
        self.folder_to_watch = None
        self.extraction_path = None
        self.watcher_thread = None
        self.watcher = None
        self.create_widgets()
        self.load_config()
        self.Bind(wx.EVT_CLOSE, self.on_minimize)

        self.tray_icon = TrayIcon(self)
        self.Bind(wx.EVT_ICONIZE, self.on_iconify)

        self.SetIcon(wx.Icon(ICON_FILE))  # 设置应用图标
        self.SetTitle(APP_TITLE)  # 仅设置一次标题

        self.Center()  # 将窗口居中

    def create_widgets(self):
        vbox = wx.BoxSizer(wx.VERTICAL)
        
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        self.watch_btn = wx.Button(self.panel, label="选择监控文件夹")
        self.watch_btn.Bind(wx.EVT_BUTTON, self.set_watch_folder)
        hbox1.Add(self.watch_btn, flag=wx.EXPAND|wx.ALL, border=5)
        
        self.watch_folder_txt = wx.TextCtrl(self.panel, style=wx.TE_READONLY | wx.TE_NOHIDESEL)
        self.watch_folder_txt.SetBackgroundColour(self.panel.GetBackgroundColour())
        self.watch_folder_txt.Bind(wx.EVT_SET_FOCUS, self.disable_focus)
        hbox1.Add(self.watch_folder_txt, proportion=1, flag=wx.EXPAND|wx.ALL, border=5)
        vbox.Add(hbox1, flag=wx.EXPAND|wx.ALL, border=5)
        
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        self.extract_btn = wx.Button(self.panel, label="选择解压文件夹")
        self.extract_btn.Bind(wx.EVT_BUTTON, self.set_extract_folder)
        hbox2.Add(self.extract_btn, flag=wx.EXPAND|wx.ALL, border=5)
        
        self.extract_folder_txt = wx.TextCtrl(self.panel, style=wx.TE_READONLY | wx.TE_NOHIDESEL)
        self.extract_folder_txt.SetBackgroundColour(self.panel.GetBackgroundColour())
        self.extract_folder_txt.Bind(wx.EVT_SET_FOCUS, self.disable_focus)
        hbox2.Add(self.extract_folder_txt, proportion=1, flag=wx.EXPAND|wx.ALL, border=5)
        vbox.Add(hbox2, flag=wx.EXPAND|wx.ALL, border=5)
        
        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        self.start_stop_btn = wx.Button(self.panel, label="开始监控", size=(100, 30))
        self.start_stop_btn.Bind(wx.EVT_BUTTON, self.toggle_watching)
        hbox3.Add(self.start_stop_btn, flag=wx.ALIGN_CENTER|wx.ALL, border=5)
        vbox.Add(hbox3, flag=wx.ALIGN_CENTER|wx.ALL, border=5)

        self.panel.SetSizer(vbox)

    def disable_focus(self, event):
        self.panel.SetFocus()

    def set_watch_folder(self, event):
        with wx.DirDialog(self, "选择监控文件夹", style=wx.DD_DEFAULT_STYLE) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                self.folder_to_watch = dialog.GetPath()
                self.watch_folder_txt.SetValue(self.folder_to_watch)
                self.save_config()

    def set_extract_folder(self, event):
        with wx.DirDialog(self, "选择解压文件夹", style=wx.DD_DEFAULT_STYLE) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                self.extraction_path = dialog.GetPath()
                self.extract_folder_txt.SetValue(self.extraction_path)
                self.save_config()

    def toggle_watching(self, event):
        if not self.check_bz_installed():
            wx.MessageBox("请先安装Bandizip：https://www.bandisoft.com/", "错误", wx.OK | wx.ICON_ERROR)
            return

        if self.watcher and self.watcher_thread and self.watcher_thread.is_alive():
            self.stop_watching()
        else:
            self.start_watching()

    def check_bz_installed(self):
        # 检查bz.exe是否在系统路径中
        return shutil.which("bz.exe") is not None

    def validate_paths(self):
        if not os.path.isdir(self.folder_to_watch):
            wx.MessageBox("监控文件夹路径无效，请重新选择。", "错误", wx.OK | wx.ICON_ERROR)
            return False
        if not os.path.isdir(self.extraction_path):
            wx.MessageBox("解压文件夹路径无效，请重新选择。", "错误", wx.OK | wx.ICON_ERROR)
            return False
        return True

    def start_watching(self):
        if not self.folder_to_watch or not self.extraction_path:
            wx.MessageBox("请先选择监控文件夹和解压文件夹", "警告", wx.OK | wx.ICON_WARNING)
            return
        if not self.validate_paths():
            return
        self.watcher = Watcher(self.folder_to_watch, self.extraction_path)
        self.watcher_thread = threading.Thread(target=self.watcher.run)
        self.watcher_thread.daemon = True  # 使用daemon属性代替setDaemon
        self.watcher_thread.start()
        self.start_stop_btn.SetLabel("停止监控")
        self.watch_btn.Disable()
        self.extract_btn.Disable()

    def stop_watching(self):
        if self.watcher:
            self.watcher.stop()
            self.watcher_thread.join()  # 等待线程结束
            self.watcher = None
            self.watcher_thread = None
        self.start_stop_btn.SetLabel("开始监控")
        self.watch_btn.Enable()
        self.extract_btn.Enable()

    def save_config(self):
        config = {
            'folder_to_watch': self.folder_to_watch,
            'extraction_path': self.extraction_path
        }
        with open(CONFIG_FILE, 'w') as config_file:
            json.dump(config, config_file)

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as config_file:
                config = json.load(config_file)
                self.folder_to_watch = config.get('folder_to_watch')
                self.extraction_path = config.get('extraction_path')
                self.watch_folder_txt.SetValue(self.folder_to_watch or "")
                self.extract_folder_txt.SetValue(self.extraction_path or "")

    def on_minimize(self, event):
        self.Hide()

    def on_iconify(self, event):
        if self.IsIconized():
            self.Hide()

    def on_quit(self, event):
        if self.watcher:
            self.watcher.stop()
        self.tray_icon.Destroy()
        self.Destroy()

class TrayIcon(wx.adv.TaskBarIcon):
    TBMENU_RESTORE = wx.NewIdRef()
    TBMENU_CLOSE   = wx.NewIdRef()

    def __init__(self, frame):
        self.frame = frame
        super(TrayIcon, self).__init__()
        self.SetIcon(wx.Icon(ICON_FILE), "autoZipper")
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DOWN, self.on_taskbar_left_click)
        self.Bind(wx.adv.EVT_TASKBAR_RIGHT_DOWN, self.on_taskbar_right_click)

    def on_taskbar_left_click(self, event):
        if not self.frame.IsShown():
            self.frame.Show()
            self.frame.Restore()

    def on_taskbar_right_click(self, event):
        menu = wx.Menu()
        restore_item = wx.MenuItem(menu, self.TBMENU_RESTORE, "打开")
        close_item = wx.MenuItem(menu, self.TBMENU_CLOSE, "关闭")
        menu.Append(restore_item)
        menu.Append(close_item)
        self.Bind(wx.EVT_MENU, self.on_restore, id=self.TBMENU_RESTORE)
        self.Bind(wx.EVT_MENU, self.on_close, id=self.TBMENU_CLOSE)
        self.PopupMenu(menu)
        menu.Destroy()

    def on_restore(self, event):
        self.frame.Show()
        self.frame.Restore()

    def on_close(self, event):
        self.frame.Close()
        wx.CallAfter(self.frame.on_quit, event)

    def ShowBalloon(self, title, msg):
        pass  # 不显示最小化到系统托盘的气泡提示

class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame(None, title=APP_TITLE)
        frame.Show()
        frame.SetSize((350, 180))  # 设置窗口大小
        frame.SetMinSize((350, 180))  # 设置窗口最小大小
        frame.SetMaxSize((350, 180))  # 设置窗口最大大小
        return True

if __name__ == "__main__":
    app = MyApp(False)
    app.MainLoop()