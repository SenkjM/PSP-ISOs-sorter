#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PSP ISO文件排序工具
功能：扫描指定文件夹的ISO文件，按创建时间排序，支持拖拽调整顺序，并修改文件创建时间
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime
import platform
import subprocess
from pathlib import Path
import time


class DragDropListbox(tk.Listbox):
    """支持拖拽排序的Listbox"""
    
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.drag_start_index = None
        self.bind('<Button-1>', self.on_click)
        self.bind('<B1-Motion>', self.on_drag)
        self.bind('<ButtonRelease-1>', self.on_drop)
        
    def on_click(self, event):
        """鼠标点击事件"""
        self.drag_start_index = self.nearest(event.y)
        
    def on_drag(self, event):
        """拖拽事件"""
        if self.drag_start_index is not None:
            current_index = self.nearest(event.y)
            if current_index != self.drag_start_index:
                # 高亮显示当前位置
                self.selection_clear(0, tk.END)
                self.selection_set(current_index)
                
    def on_drop(self, event):
        """释放事件"""
        if self.drag_start_index is not None:
            drop_index = self.nearest(event.y)
            if drop_index != self.drag_start_index:
                # 移动项目
                item = self.get(self.drag_start_index)
                self.delete(self.drag_start_index)
                self.insert(drop_index, item)
                self.selection_clear(0, tk.END)
                self.selection_set(drop_index)
                
                # 通知父窗口更新数据
                if hasattr(self.master, 'on_list_reorder'):
                    self.master.on_list_reorder(self.drag_start_index, drop_index)
                    
        self.drag_start_index = None


class PSPISOSorter:
    """PSP ISO文件排序工具主类"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("PSP ISO文件排序工具")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # 数据存储
        self.current_folder = tk.StringVar()
        self.iso_files = []  # 存储[(文件路径, 创建时间), ...]
        
        self.setup_ui()
        self.scan_folder()
        
    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 文件夹选择区域
        folder_frame = ttk.LabelFrame(main_frame, text="文件夹选择", padding="5")
        folder_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        folder_frame.columnconfigure(1, weight=1)
        
        ttk.Label(folder_frame, text="当前文件夹:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.current_folder, state="readonly")
        self.folder_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(folder_frame, text="浏览...", command=self.browse_folder).grid(row=0, column=2)
        ttk.Button(folder_frame, text="刷新", command=self.scan_folder).grid(row=0, column=3, padx=(5, 0))
        
        # 信息显示区域
        info_frame = ttk.Frame(main_frame)
        info_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.info_label = ttk.Label(info_frame, text="请选择包含PSP ISO文件的文件夹")
        self.info_label.grid(row=0, column=0, sticky=tk.W)
        
        # 文件列表区域
        list_frame = ttk.LabelFrame(main_frame, text="ISO文件列表 (拖拽调整顺序)", padding="5")
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 创建支持拖拽的列表框
        self.file_listbox = DragDropListbox(list_frame, height=15)
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(10, 0))
        
        ttk.Button(button_frame, text="完成调整", command=self.apply_changes, 
                  style="Accent.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="重置顺序", command=self.reset_order).pack(side=tk.LEFT)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def browse_folder(self):
        """浏览文件夹"""
        folder = filedialog.askdirectory(title="选择包含PSP ISO文件的文件夹")
        if folder:
            self.current_folder.set(folder)
            self.scan_folder()
            
    def scan_folder(self):
        """扫描文件夹中的ISO文件"""
        folder = self.current_folder.get()
        if not folder or not os.path.exists(folder):
            self.iso_files = []
            self.update_file_list()
            self.info_label.config(text="请选择有效的文件夹")
            return
            
        try:
            self.status_var.set("正在扫描文件夹...")
            self.root.update()
            
            # 查找所有ISO文件
            iso_files = []
            for file in os.listdir(folder):
                if file.lower().endswith('.iso'):
                    file_path = os.path.join(folder, file)
                    if os.path.isfile(file_path):
                        # 获取文件创建时间
                        creation_time = self.get_creation_time(file_path)
                        iso_files.append((file_path, creation_time))
            
            # 按创建时间降序排序（最新的在前）
            iso_files.sort(key=lambda x: x[1], reverse=True)
            
            self.iso_files = iso_files
            self.update_file_list()
            
            count = len(iso_files)
            self.info_label.config(text=f"找到 {count} 个ISO文件，按创建时间排序（最新在上）")
            self.status_var.set(f"找到 {count} 个ISO文件")
            
        except Exception as e:
            messagebox.showerror("错误", f"扫描文件夹时出错：{str(e)}")
            self.status_var.set("扫描失败")
            
    def get_creation_time(self, file_path):
        """获取文件创建时间"""
        try:
            if platform.system() == 'Windows':
                # Windows系统获取创建时间
                stat = os.stat(file_path)
                return stat.st_ctime
            else:
                # 其他系统使用修改时间
                stat = os.stat(file_path)
                return stat.st_mtime
        except:
            return 0
            
    def update_file_list(self):
        """更新文件列表显示"""
        self.file_listbox.delete(0, tk.END)
        
        for i, (file_path, creation_time) in enumerate(self.iso_files):
            file_name = os.path.basename(file_path)
            time_str = datetime.datetime.fromtimestamp(creation_time).strftime("%Y-%m-%d %H:%M:%S")
            display_text = f"{i+1:2d}. {file_name} ({time_str})"
            self.file_listbox.insert(tk.END, display_text)
            
    def on_list_reorder(self, from_index, to_index):
        """列表重新排序时的回调"""
        if 0 <= from_index < len(self.iso_files) and 0 <= to_index < len(self.iso_files):
            # 移动数据
            item = self.iso_files.pop(from_index)
            self.iso_files.insert(to_index, item)
            # 更新显示
            self.update_file_list()
            self.status_var.set("顺序已调整，点击'完成调整'应用更改")
            
    def reset_order(self):
        """重置为原始顺序（按创建时间排序）"""
        if not self.iso_files:
            return
            
        # 重新按创建时间排序
        self.iso_files.sort(key=lambda x: x[1], reverse=True)
        self.update_file_list()
        self.status_var.set("已重置为按创建时间排序")
        
    def apply_changes(self):
        """应用更改，修改文件创建时间"""
        if not self.iso_files:
            messagebox.showwarning("警告", "没有可处理的文件")
            return
            
        # 确认对话框
        result = messagebox.askyesno(
            "确认操作", 
            f"即将修改 {len(self.iso_files)} 个文件的创建时间。\n"
            "此操作不可撤销，是否继续？",
            icon="warning"
        )
        
        if not result:
            return
            
        try:
            self.status_var.set("正在修改文件时间...")
            self.root.update()
            
            # 获取当前时间作为基准
            base_time = time.time()
            success_count = 0
            
            # 按当前列表顺序，从最新时间开始递减修改
            for i, (file_path, _) in enumerate(self.iso_files):
                # 每个文件间隔1秒
                new_time = base_time - i
                
                try:
                    self.set_file_time(file_path, new_time)
                    success_count += 1
                    
                    # 更新进度
                    progress = f"正在处理 {i+1}/{len(self.iso_files)}: {os.path.basename(file_path)}"
                    self.status_var.set(progress)
                    self.root.update()
                    
                except Exception as e:
                    print(f"修改文件 {file_path} 的时间失败: {e}")
                    
            # 重新扫描以显示更新后的时间
            self.scan_folder()
            
            messagebox.showinfo("完成", f"成功修改了 {success_count} 个文件的创建时间")
            self.status_var.set(f"完成：成功修改 {success_count} 个文件")
            
        except Exception as e:
            messagebox.showerror("错误", f"修改文件时间时出错：{str(e)}")
            self.status_var.set("修改失败")
            
    def set_file_time(self, file_path, timestamp):
        """设置文件的创建时间和修改时间"""
        try:
            # 设置访问时间和修改时间
            os.utime(file_path, (timestamp, timestamp))
            
            # Windows系统额外设置创建时间
            if platform.system() == 'Windows':
                try:
                    import win32file
                    import win32con
                    from pywintypes import Time
                    
                    # 转换时间格式
                    dt = datetime.datetime.fromtimestamp(timestamp)
                    win_time = Time(dt)
                    
                    # 打开文件句柄
                    handle = win32file.CreateFile(
                        file_path,
                        win32con.GENERIC_WRITE,
                        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                        None,
                        win32con.OPEN_EXISTING,
                        0,
                        None
                    )
                    
                    # 设置文件时间
                    win32file.SetFileTime(handle, win_time, win_time, win_time)
                    win32file.CloseHandle(handle)
                    
                except ImportError:
                    # 如果没有pywin32，使用PowerShell命令
                    self.set_creation_time_powershell(file_path, timestamp)
                    
        except Exception as e:
            raise Exception(f"设置文件时间失败: {e}")
            
    def set_creation_time_powershell(self, file_path, timestamp):
        """使用PowerShell设置Windows文件创建时间"""
        try:
            dt = datetime.datetime.fromtimestamp(timestamp)
            time_str = dt.strftime("%m/%d/%Y %H:%M:%S")
            
            cmd = f'(Get-Item "{file_path}").CreationTime = "{time_str}"'
            subprocess.run(['powershell', '-Command', cmd], 
                          check=True, capture_output=True, text=True)
        except:
            pass  # 如果PowerShell命令失败，继续执行


def main():
    """主函数"""
    root = tk.Tk()
    app = PSPISOSorter(root)
    
    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap(default='icon.ico')
    except:
        pass
        
    # 启动主循环
    root.mainloop()


if __name__ == "__main__":
    main()
