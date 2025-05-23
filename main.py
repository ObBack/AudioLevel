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
import win32event
import math
import win32api
import os
from tkinter import messagebox
from tkinter import simpledialog
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from PIL import Image
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

class AudioLevelSetter:
    def __init__(self, root_window):  # 初始化程序
        # 窗口
        self.root = root_window
        self.root.protocol('WM_DELETE_WINDOW', self.root.withdraw())
        self.root.withdraw()
        self.root.geometry("300x100+-500+-300")
        self.root.title("Audio Level Setter - 音量设置器")  # 标题
        self.set_audio_size(bypass_password=True)
        self.audio_control()
        self.audio_size = self.original_volume  # 音量大小
        self.password_set() # 密码
        self.create_tray_icon() # 托盘
        self.running = True  # 运行状态
        self.volume_thread = threading.Thread(target=self.adjust_volume_loop)
        self.volume_thread.daemon = True
        self.volume_thread.start()

    def password_set(self): # 根据音量设置密码
        self.password = str(math.floor(self.volume.GetMasterVolumeLevelScalar() * 100) * 2)

    def audio_control(self):  # 音频控制
        print("音频控制...")
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(interface, POINTER(IAudioEndpointVolume))
        self.original_volume = self.volume.GetMasterVolumeLevelScalar()

    def password_detection(self): # 密码检测
        temp_root = tk.Toplevel(self.root)
        temp_root.withdraw()
        temp_root.grab_set()  # 强制焦点

        input_pwd = None
        try:
            input_pwd = simpledialog.askstring(
                "密码验证 - Password Verification",
                "请输入密码: \nPlease enter password:",
                show='*',
                parent=temp_root
            )
        finally:
            temp_root.grab_release()
            temp_root.destroy()

        if input_pwd is None:
            return False
        if str(input_pwd) == self.password:
            return True
        messagebox.showerror("错误 - Error", "密码错误，请重新输入！ \nPassword error, please try again!", parent=self.root)
        return False

    def safe_exit(self):  # 安全退出
        print("退出...")
        if self.password_detection():
            if messagebox.askyesno("确认 - Confirm", "确定要退出吗？ \nAre you sure to exit?", parent=self.root):
                self.running = False
                self.volume_thread.join() # 等待线程结束
                self.tray_icon.stop()
                self.restore_volume()
                self.root.quit()  # 退出
                self.root.destroy()  # 最后销毁窗口
        
    def restore_volume(self):  # 恢复音量
        if hasattr(self, 'original_volume'):
            self.volume.SetMasterVolumeLevelScalar(self.original_volume, None)
            print(f"\n音量已恢复至 {self.original_volume*100:.0f}%")

    def adjust_volume_loop(self):  # 控音量
        print(f"音量已保存: {self.original_volume*100:.0f}%")
        print(f"音量已调整至{self.audio_size*100:.0f}%...")
        try:
            while self.running:
                self.volume.SetMasterVolumeLevelScalar(self.audio_size, None)
                time.sleep(0.1)
        except Exception as e:
            print(f"异常: {e}")
        finally:
            self.restore_volume()

    def create_tray_icon(self):  # 托盘
        print("创建托盘...")
        if os.path.exists("icon.ico"):
            image = Image.open("icon.ico")
        else:
            image = Image.new('RGB', (64, 64), 'black')
        menu = pystray.Menu(
            pystray.MenuItem('退出 - Exit', lambda: self.root.after(0, self.safe_exit)), 
            pystray.MenuItem('设置音量 - Set Volume', lambda: self.root.after(0, self.set_audio_size))
        )
        self.tray_icon = pystray.Icon("Audio Level Setter", image, "Audio Level Setter", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def set_audio_size(self, bypass_password=False):  # 调整设置音量
        if bypass_password or self.password_detection():
            settings_window = tk.Toplevel(self.root)
            settings_window.resizable(False, False)
            settings_window.geometry("400x200+600+400")
            settings_window.title("Audio Level Set - 音量设置")
            
            label = tk.Label(settings_window, text="更改音量为：\nset volume to:\n(0-100)")
            label.pack(pady=10)
            
            entry = tk.Entry(settings_window)
            entry.pack()
            
            def apply_changes():
                try:
                    new_size = int(entry.get())
                    if 0 <= new_size <= 100:
                        self.audio_size = new_size / 100.0
                        print(f"音量已更新至 {self.audio_size * 100:.0f}%")
                        self.password = str(new_size *2)
                        print(f"密码已更新: {self.password}")
                        messagebox.showinfo("成功 - Success", "设置已保存！ \nSetting saved!", parent=self.root)
                        settings_window.destroy()
                    else:
                        messagebox.showerror("错误 - Error", "请输入0到100之间的整数! \nPlease enter an integer between 0 and 100.")
                except ValueError:
                    messagebox.showerror("错误 - Error", "错误！\nError!", parent=self.root)
            
            button = tk.Button(settings_window, text="确认 - Confirm", command=apply_changes)
            button.pack(pady=10)
        else:
            messagebox.showerror("错误 - Error", "!", parent=self.root)

if __name__ == "__main__":
    # 检测管理员
    if ctypes.windll.shell32.IsUserAnAdmin():
        # 互斥锁
        mutex = win32event.CreateMutex(None, False, "AudioLevelSetter_Mutex")
        if win32api.GetLastError() == 183:
            messagebox.showerror("错误 - Error", "程序已在运行！ \nProgram is already running!")
            exit(0)
        sys.mutex = mutex
        # 启动
        root = tk.Tk()
        app = AudioLevelSetter(root)
        root.mainloop()
    else:
        # 管理员获取
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{__file__}"', None, 1)
        time.sleep(1)  # 等待
        sys.exit(0)