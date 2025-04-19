# -*- coding: utf-8 -*-
# @Time    : 2025/4/18 12:43
# @File    : main.py
# @Author  : 枫蒲晚霞 ObBack
# @Email   : 1989397949@qq.com

import time
import pystray
import tkinter as tk
import threading
import ctypes
import sys
import hashlib
import win32event
import win32api
import random
import os
from tkinter import messagebox
from tkinter import simpledialog
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from PIL import Image
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

class AudioLevel:
    def __init__(self, root_window):  # 初始化程序
        # 窗口
        self.root = root_window
        self.root.protocol('WM_DELETE_WINDOW', self.root.withdraw())
        self.root.withdraw()  # 隐藏主窗口
        self.root.title("Audio Level")
        
        # 变量
        self.random_number = random.randint(90, 100)
        self.admin_password = hashlib.sha256(("91" + str(self.random_number) + "91").encode()).hexdigest() # 密码
        self.audio_size = self.random_number / 100  # 音量大小

        self.audio_control()# 音频控制
        self.create_tray_icon()# 托盘

        # 音量线程
        self.running = True
        self.volume_thread = threading.Thread(target=self.adjust_volume_loop)
        self.volume_thread.daemon = True
        self.volume_thread.start()

    def audio_control(self):  # 音频控制
        print("音频控制...")
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(interface, POINTER(IAudioEndpointVolume))
        self.original_volume = self.volume.GetMasterVolumeLevelScalar()

    def password_verification(self): # 输入密码
        temp_root = tk.Toplevel(self.root)
        temp_root.withdraw()
        temp_root.grab_set()  # 强制焦点

        input_pwd = None
        try:
            input_pwd = simpledialog.askstring(
                "密码验证",
                "请输入管理员密码:",
                show='*',
                parent=temp_root
            )
        finally:
            temp_root.grab_release()
            temp_root.destroy()

        if input_pwd is None:
            return False
        if hashlib.sha256(input_pwd.encode()).hexdigest() == self.admin_password:
            return True
        messagebox.showerror("错误", "密码错误，请重新输入！", parent=self.root)
        return False

    def safe_exit(self):  # 安全退出
        print("退出...")
        if self.password_verification():
            if messagebox.askyesno("确认", "确定要退出吗？", parent=self.root):
                self.running = False
                self.volume_thread.join()  # 等待线程结束
                self.tray_icon.stop()
                self.restore_volume()
                self.root.quit()  # 正确退出主循环
                self.root.destroy()  # 最后销毁窗口
        
    def restore_volume(self):  # 恢复音量
        if hasattr(self, 'original_volume'):
            self.volume.SetMasterVolumeLevelScalar(self.original_volume, None)
            print(f"\n音量已恢复至 {self.original_volume*100:.0f}%")

    def adjust_volume_loop(self):  # 控音量线程
        print(f"音量已保存: {self.original_volume*100:.0f}%")
        print(f"音量已调整至{self.audio_size*100:.0f}%...")  # 修改提示信息
        try:
            while self.running:
                self.volume.SetMasterVolumeLevelScalar(self.audio_size, None)
                time.sleep(0.1)
        finally:
            self.restore_volume()

    def create_tray_icon(self):  # 托盘
        print("创建托盘...")
        if os.path.exists("icon.ico"):
            image = Image.open("icon.ico")
        else:
            image = Image.new('RGB', (64, 64), 'blue')
        menu = pystray.Menu(
            pystray.MenuItem('退出', lambda: self.root.after(0, self.safe_exit))
        )
        self.tray_icon = pystray.Icon("audio_minimizer", image, "音量控制器", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def no_kill(self):  # 防杀进程
        print("防杀...")

if __name__ == "__main__":
    # 检测管理员
    if ctypes.windll.shell32.IsUserAnAdmin():
        # 互斥锁
        mutex = win32event.CreateMutex(None, False, "AudioLevelMinimizer_Mutex")
        if win32api.GetLastError() == 183:
            messagebox.showerror("错误", "程序已在运行中！")
            exit(0)
        sys.mutex = mutex
        # 启动
        root = tk.Tk()
        app = AudioLevel(root)
        root.mainloop()
    else:
        # 管理员
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{__file__}"', None, 1)
        time.sleep(1)  # 等待
        sys.exit(0)