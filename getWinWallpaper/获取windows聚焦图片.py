import os
import shutil
import datetime
import getpass
from tkinter import messagebox

# 定义原始文件夹和目标文件夹的路径
src_folder = f'C:\\Users\\{getpass.getuser()}\\AppData\\Roaming\\Microsoft\\Windows\\Themes\\CachedFiles'
dest_folder = ''
desk_letter = ''

if not os.path.exists(src_folder):
    messagebox.showinfo(
        title='提示',
        message=f'没打开聚焦功能\n'
                f'请先在个性化中打开聚焦功能'
    )
else:
    # 定义盘符优先级顺序
    drive_order = ['F:', 'D:', 'C:']
    # 遍历盘符列表，判断是否存在
    for drive in drive_order:
        if os.path.exists(drive):
            desk_letter = drive
            break
    else:
        pass

    if desk_letter == 'C:':
        dest_folder = f'C:\\Users\\{getpass.getuser()}\\Pictures\\聚焦图片'
    else:
        dest_folder = f'{desk_letter}\\聚焦图片'

    if not os.path.exists(dest_folder):
        os.makedirs(dest_folder)
    else:
        pass

    # 获取当前日期
    now = datetime.datetime.now()
    date_str = now.strftime('%Y-%m-%d')

    # 拼接文件夹路径
    new_folder = os.path.join(dest_folder, date_str)

    # 判断文件夹是否存在，不存在则创建
    if not os.path.exists(new_folder):
        os.mkdir(new_folder)
    else:
        pass

    # 修改目标文件夹路径为新的文件夹路径
    new_dest_folder = new_folder

    # 获取原始文件夹内所有文件的文件名列表
    file_names = os.listdir(src_folder)

    # 获取前一天的文件夹路径与文件
    yesterday = now - datetime.timedelta(days=1)
    yesterday_str = yesterday.strftime('%Y-%m-%d')
    yesterday_folder = os.path.join(dest_folder, yesterday_str)
    if os.path.exists(yesterday_folder):
        yesterday_file_names = os.listdir(yesterday_folder)

    # 遍历文件名列表，将每个文件拷贝到目标文件夹
    for file_name in file_names:
        src_path = os.path.join(src_folder, file_name)
        dest_path = os.path.join(new_dest_folder, file_name + '.png')

        if os.path.exists(yesterday_folder):
            if not file_name + '.png' in yesterday_file_names:
                shutil.copy(src_path, dest_path)
            else:
                continue
        else:
            shutil.copy(src_path, dest_path)

    messagebox.showinfo(
        title='提示',
        message=f'获取成功\n'
                f'存放地址：{dest_folder}'
    )
