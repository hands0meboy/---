import os
import sys
import shutil

def delete_files_in_directory(exe_path):
    # 获取exe所在的目录
    directory = os.path.dirname(os.path.abspath(exe_path))
    
    # 遍历目录下的所有文件和子目录
    for root, dirs, files in os.walk(directory):
        for item in files + dirs:
            # 构建完整路径
            full_path = os.path.join(root, item)
            # 检查当前项是否为exe文件本身，如果不是则删除
            if full_path != os.path.abspath(exe_path):
                try:
                    if os.path.isfile(full_path):
                        # 移除只读属性
                        os.chmod(full_path, 0o777)
                        os.remove(full_path)  # 删除文件
                    else:
                        # 递归删除目录，并移除只读属性
                        for sub_root, sub_dirs, sub_files in os.walk(full_path, topdown=False):
                            for name in sub_files:
                                file_path = os.path.join(sub_root, name)
                                os.chmod(file_path, 0o777)
                                os.remove(file_path)
                            for name in sub_dirs:
                                subdirectory_path = os.path.join(sub_root, name)
                                os.rmdir(subdirectory_path)
                        os.rmdir(full_path)  # 删除根目录
                except PermissionError as e:
                    print(f"PermissionError: Cannot delete {full_path} - {e}")
                except Exception as e:
                    print(f"Error: Cannot delete {full_path} - {e}")

if __name__ == "__main__":
    delete_files_in_directory(sys.executable)
