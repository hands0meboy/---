import os
import time
import sys
import pythoncom
import json
import threading
import logging
import wx
import wx.adv
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32com.client as win32

CONFIG_FILE = 'config.json'
ICON_FILE = 'ico.ico'  # 确保图标文件与脚本在同一目录下
APP_TITLE = "doc2docx"

#logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logging.disable

class Watcher:
    def __init__(self, folder_to_watch):
        self.folder_to_watch = folder_to_watch
        self.observer = Observer()
        self._stop_event = threading.Event()

    def run(self):
        logging.debug(f"开始监控文件夹: {self.folder_to_watch}")
        event_handler = Handler()
        self.observer.schedule(event_handler, self.folder_to_watch, recursive=True)
        self.observer.start()
        """try:
            while not self._stop_event.is_set():
                logging.debug("监控线程运行中...")
                time.sleep(5)
        except KeyboardInterrupt:
            self.stop() """
        self.observer.join()
    def stop(self):
        logging.debug("停止监控")
        self._stop_event.set()
        self.observer.stop()
        self.observer.join()  # 添加这行代码等待观察者线程结束

class Handler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()

    def on_created(self, event):
        logging.debug(f"检测到新文件: {event.src_path}")
        if not event.is_directory and event.event_type == 'created':
            file_path = event.src_path
            if file_path.endswith('.doc'):
                logging.debug(f"检测到 .doc 文件: {file_path}")
                if self.is_file_ready(file_path):
                    try:
                        logging.debug(f"开始转换文件: {file_path}")
                        self.convert_doc_to_docx(file_path)
                    except Exception as e:
                        logging.debug(f"文件转换失败: {file_path}, 错误: {e}")
                else:
                    logging.debug(f"文件未准备好: {file_path}")

    def is_file_ready(self, file_path):
        logging.debug(f"检查文件是否准备好: {file_path}")
        if os.path.basename(file_path).startswith('~$'):
            logging.debug("文件是临时文件，忽略")
            return False

        previous_size = -1
        while True:
            try:
                current_size = os.path.getsize(file_path)
            except FileNotFoundError:
                logging.debug("文件未找到，可能已被移动或删除")
                return False

            if current_size == previous_size:
                logging.debug("文件已准备好")
                return True
            previous_size = current_size
            logging.debug(f"文件大小变化中: {current_size}")
            time.sleep(1)

    def convert_doc_to_docx(self, file_path):
        try:
            pythoncom.CoInitialize()  # 初始化COM库
            file_path = os.path.abspath(file_path)
            docx_path = file_path.replace(".doc", ".docx")
            logging.debug(f"转换文件: {file_path} 到 {docx_path}")

            word = win32.Dispatch("Word.Application")
            doc = word.Documents.Open(file_path)
            doc.SaveAs(docx_path, FileFormat=12)
            doc.Close()
            word.Quit()

            logging.debug(f"文件转换成功: {docx_path}")

        except Exception as e:
            logging.debug(f"文件转换过程中出现错误: {e}")

        finally:
            pythoncom.CoUninitialize()  # 取消初始化COM库

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super(MyFrame, self).__init__(parent, title=title, size=(500, 160), style=wx.DEFAULT_FRAME_STYLE & ~(wx.MAXIMIZE_BOX | wx.RESIZE_BORDER))
        self.panel = wx.Panel(self)
        self.folder_to_watch = None
        self.watcher_thread = None
        self.watcher = None
        self.create_widgets()
        self.load_config()  # 确保在初始化时加载配置并启动监控
        self.Bind(wx.EVT_CLOSE, self.on_minimize)

        self.tray_icon = TrayIcon(self)
        self.Bind(wx.EVT_ICONIZE, self.on_iconify)

        self.SetIcon(wx.Icon(ICON_FILE))  # 设置应用图标
        self.SetTitle(APP_TITLE)  # 仅设置一次标题

        self.Center()  # 将窗口居中

    def create_widgets(self):
        vbox = wx.BoxSizer(wx.VERTICAL)
    
        # 添加伸展空间
        vbox.AddStretchSpacer()
        
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)
        self.watch_btn = wx.Button(self.panel, label="选择监控文件夹")
        self.watch_btn.Bind(wx.EVT_BUTTON, self.set_watch_folder)
        hbox1.Add(self.watch_btn, flag=wx.ALL, border=5)
        
        self.watch_folder_txt = wx.TextCtrl(self.panel, style=wx.TE_READONLY | wx.TE_NOHIDESEL)
        self.watch_folder_txt.SetBackgroundColour(self.panel.GetBackgroundColour())
        self.watch_folder_txt.Bind(wx.EVT_SET_FOCUS, self.disable_focus)
        hbox1.Add(self.watch_folder_txt, proportion=1, flag=wx.EXPAND|wx.ALL, border=5)
        vbox.Add(hbox1, flag=wx.EXPAND|wx.ALL, border=5)
        
        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        self.start_stop_btn = wx.Button(self.panel, label="开始监控", size=(100, 30))
        self.start_stop_btn.Bind(wx.EVT_BUTTON, self.toggle_watching)
        hbox3.Add(self.start_stop_btn, flag=wx.ALL, border=5)
        vbox.Add(hbox3, flag=wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, border=5)

        # 添加伸展空间
        vbox.AddStretchSpacer()

        self.panel.SetSizer(vbox)
        vbox.Layout()

    def disable_focus(self, event):
        self.panel.SetFocus()

    def set_watch_folder(self, event):
        with wx.DirDialog(self, "选择监控文件夹", style=wx.DD_DEFAULT_STYLE) as dialog:
            if dialog.ShowModal() == wx.ID_OK:
                self.folder_to_watch = dialog.GetPath()
                self.watch_folder_txt.SetValue(self.folder_to_watch)
                self.save_config()
                logging.debug(f"已选择新监控文件夹路径: {self.folder_to_watch}")

    def toggle_watching(self, event):
        if self.watcher and self.watcher_thread and self.watcher_thread.is_alive():
            self.stop_watching()
            logging.debug("监控已停止")
        else:
            self.start_watching()
            logging.debug("监控已启动")

    def validate_paths(self):
        logging.debug(f"验证路径: {self.folder_to_watch}")
        if not os.path.isdir(self.folder_to_watch):
            wx.MessageBox("监控文件夹路径无效，请重新选择。", "错误", wx.OK | wx.ICON_ERROR)
            logging.debug("监控文件夹路径无效")
            return False
        return True

    def start_watching(self):
        logging.debug("尝试开始监控...")
        if not self.folder_to_watch:
            wx.MessageBox("请先选择监控文件夹", "警告", wx.OK | wx.ICON_WARNING)
            logging.debug("监控文件夹未选择，无法开始监控")
            return
        if not self.validate_paths():
            logging.debug("监控文件夹路径无效")
            return
        logging.debug(f"监控文件夹: {self.folder_to_watch}")
        self.watcher = Watcher(self.folder_to_watch)
        self.watcher_thread = threading.Thread(target=self.watcher.run)
        self.watcher_thread.daemon = True  # 使用daemon属性代替setDaemon
        self.watcher_thread.start()
        self.start_stop_btn.SetLabel("停止监控")
        self.watch_btn.Disable()
        logging.debug("监控已开始")

    def stop_watching(self):
        if self.watcher:
            self.watcher.stop()
            self.watcher_thread.join()  # 等待线程结束
            self.watcher = None
            self.watcher_thread = None
        self.start_stop_btn.SetLabel("开始监控")
        self.watch_btn.Enable()

    def save_config(self):
        config = {
            'folder_to_watch': self.folder_to_watch
        }
        with open(CONFIG_FILE, 'w') as config_file:
            json.dump(config, config_file)

    def load_config(self):
        logging.debug("加载配置文件.")
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as config_file:
                config = json.load(config_file)
                self.folder_to_watch = config.get('folder_to_watch')
                self.watch_folder_txt.SetValue(self.folder_to_watch or "")
                logging.debug(f"已加载监控文件夹路径: {self.folder_to_watch}")

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
        self.SetIcon(wx.Icon(ICON_FILE), "doc2docx")
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
