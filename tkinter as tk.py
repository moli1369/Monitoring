import tkinter as tk
from tkinter import ttk, messagebox
import win32gui
import win32con
import win32api
import time
import threading
from datetime import datetime, timedelta
import sys

class FullscreenWindowManager:
    def __init__(self, root):
        self.root = root
        self.root.title("مدیریت نمایش تمام صفحه")
        self.root.geometry("800x600")
        
        # متغیرهای برنامه
        self.available_windows = []
        self.selected_windows = []
        self.is_running = False
        self.rotation_thread = None
        self.current_window_index = 0
        self.next_change_time = None
        
        # ایجاد رابط کاربری
        self.setup_ui()
        
        # بارگیری اولیه پنجره‌ها
        self.refresh_windows_list()
        
        # شروع به روزرسانی وضعیت
        self.update_status()

    def setup_ui(self):
        # فریم اصلی
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # لیست پنجره‌های موجود
        ttk.Label(main_frame, text="پنجره‌های موجود:").pack(anchor='w')
        self.windows_listbox = tk.Listbox(main_frame, selectmode=tk.MULTIPLE, height=15, font=('Tahoma', 10))
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.windows_listbox.yview)
        self.windows_listbox.configure(yscrollcommand=scrollbar.set)
        self.windows_listbox.pack(side="left", fill=tk.BOTH, expand=True, pady=10)
        scrollbar.pack(side="right", fill="y")
        
        # دکمه‌های کنترل
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="بارگیری مجدد لیست", command=self.refresh_windows_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="اضافه کردن به انتخابی‌ها", command=self.add_to_selected).pack(side=tk.LEFT, padx=5)
        
        # لیست پنجره‌های انتخابی
        ttk.Label(main_frame, text="پنجره‌های انتخابی:").pack(anchor='w')
        self.selected_listbox = tk.Listbox(main_frame, height=8, font=('Tahoma', 10))
        self.selected_listbox.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # کنترل‌های چرخش
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=15)
        
        ttk.Label(control_frame, text="مدت نمایش هر پنجره (دقیقه):").pack(side=tk.LEFT)
        self.interval_var = tk.StringVar(value="1")
        ttk.Entry(control_frame, textvariable=self.interval_var, width=5).pack(side=tk.LEFT, padx=5)
        
        self.start_button = ttk.Button(control_frame, text="شروع نمایش", command=self.toggle_rotation)
        self.start_button.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(control_frame, text="حذف از انتخابی‌ها", command=self.remove_from_selected).pack(side=tk.LEFT, padx=5)
        
        # نوار وضعیت
        self.status_frame = ttk.Frame(main_frame, relief=tk.SUNKEN)
        self.status_frame.pack(fill=tk.X, pady=(15,0))
        
        self.current_window_label = ttk.Label(self.status_frame, text="پنجره فعلی: -")
        self.current_window_label.pack(side=tk.LEFT, padx=10)
        
        self.time_remaining_label = ttk.Label(self.status_frame, text="زمان باقیمانده: -")
        self.time_remaining_label.pack(side=tk.RIGHT, padx=10)
    
    def refresh_windows_list(self):
        """بارگیری لیست پنجره‌های موجود"""
        self.available_windows = []
        self.windows_listbox.delete(0, tk.END)
        
        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title.strip():
                    self.available_windows.append((hwnd, title))
        
        win32gui.EnumWindows(enum_handler, None)
        
        for hwnd, title in self.available_windows:
            self.windows_listbox.insert(tk.END, f"{title[:70]}... (ID: {hwnd})" if len(title) > 70 else f"{title} (ID: {hwnd})")
    
    def add_to_selected(self):
        """اضافه کردن پنجره‌های انتخاب شده به لیست انتخابی"""
        selections = self.windows_listbox.curselection()
        for index in selections:
            if index < len(self.available_windows):
                hwnd, title = self.available_windows[index]
                if (hwnd, title) not in self.selected_windows:
                    self.selected_windows.append((hwnd, title))
                    self.selected_listbox.insert(tk.END, title)
    
    def remove_from_selected(self):
        """حذف پنجره‌های انتخاب شده از لیست انتخابی"""
        selections = self.selected_listbox.curselection()
        for index in reversed(selections):
            if index < len(self.selected_windows):
                self.selected_windows.pop(index)
                self.selected_listbox.delete(index)
    
    def toggle_rotation(self):
        """شروع یا توقف چرخش پنجره‌ها"""
        if not self.is_running:
            self.start_rotation()
        else:
            self.stop_rotation()
    
    def start_rotation(self):
        """شروع چرخش پنجره‌ها"""
        if not self.selected_windows:
            messagebox.showwarning("هشدار", "لطفاً حداقل یک پنجره انتخاب کنید")
            return
        
        try:
            interval = float(self.interval_var.get())
            if interval <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("خطا", "لطفاً یک عدد معتبر برای فاصله زمانی وارد کنید")
            return
        
        self.is_running = True
        self.current_window_index = 0
        self.start_button.config(text="توقف نمایش")
        
        # غیرفعال کردن کنترل‌ها در حین اجرا
        self.windows_listbox.config(state=tk.DISABLED)
        self.selected_listbox.config(state=tk.DISABLED)
        
        # Minimize کردن پنجره اصلی (نه مخفی کردن)
        self.root.iconify()  # این خط پنجره را به تسک بار می‌فرستد
        
        self.rotation_thread = threading.Thread(
            target=self.rotate_windows,
            args=(interval,),
            daemon=True
        )
        self.rotation_thread.start()
    
    def stop_rotation(self):
        """توقف چرخش پنجره‌ها"""
        self.is_running = False
        self.start_button.config(text="شروع نمایش")
        self.next_change_time = None
        
        # فعال کردن مجدد کنترل‌ها
        self.windows_listbox.config(state=tk.NORMAL)
        self.selected_listbox.config(state=tk.NORMAL)
        
        # بازگرداندن پنجره اصلی از حالت Minimize
        self.root.deiconify()
        self.root.lift()  # پنجره را به جلو بیاور
        
        self.update_status_labels("-", "-")
    
    def set_fullscreen(self, hwnd):
        """تنظیم پنجره به حالت تمام صفحه"""
        try:
            # ابتدا پنجره را به حالت عادی بازگردان
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            
            # سپس به حالت تمام صفحه برو
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            
            # تنظیم موقعیت و اندازه پنجره برای تمام صفحه
            screen_width = win32api.GetSystemMetrics(0)
            screen_height = win32api.GetSystemMetrics(1)
            win32gui.SetWindowPos(
                hwnd,
                win32con.HWND_TOP,
                0, 0,
                screen_width, screen_height,
                win32con.SWP_SHOWWINDOW
            )
            
            # فعال کردن پنجره
            win32gui.SetForegroundWindow(hwnd)
            return True
        except Exception as e:
            print(f"خطا در تنظیم تمام صفحه: {e}")
            return False
    
    def rotate_windows(self, interval_minutes):
        """چرخش بین پنجره‌های انتخابی"""
        interval_seconds = interval_minutes * 60
        
        while self.is_running and self.selected_windows:
            try:
                hwnd, title = self.selected_windows[self.current_window_index]
                
                # نمایش پنجره به صورت تمام صفحه
                if self.set_fullscreen(hwnd):
                    # محاسبه زمان تغییر بعدی
                    self.next_change_time = datetime.now() + timedelta(seconds=interval_seconds)
                    self.update_status_labels(title, interval_seconds)
                    
                    # انتظار برای مدت زمان مشخص
                    start_time = time.time()
                    while self.is_running and (time.time() - start_time) < interval_seconds:
                        time.sleep(0.5)
                        remaining = max(0, interval_seconds - (time.time() - start_time))
                        self.update_status_labels(title, remaining)
                    
                    # رفتن به پنجره بعدی
                    if self.is_running:
                        self.current_window_index = (self.current_window_index + 1) % len(self.selected_windows)
                else:
                    # اگر پنجره تمام صفحه نشد، به پنجره بعدی برو
                    self.current_window_index = (self.current_window_index + 1) % len(self.selected_windows)
                    time.sleep(1)
            
            except Exception as e:
                print(f"خطا در چرخش پنجره‌ها: {e}")
                time.sleep(1)
                continue
        
        # پس از توقف، پنجره اصلی را بازگردان
        self.root.after(0, lambda: [self.root.deiconify(), self.root.lift()])
        self.stop_rotation()
    
    def update_status_labels(self, current_window, remaining_time):
        """به روزرسانی برچسب‌های وضعیت"""
        if isinstance(remaining_time, (int, float)):
            mins, secs = divmod(int(remaining_time), 60)
            time_str = f"{mins:02d}:{secs:02d}"
        else:
            time_str = str(remaining_time)
        
        def update():
            self.current_window_label.config(text=f"پنجره فعلی: {current_window[:50]}..." if len(current_window) > 50 else f"پنجره فعلی: {current_window}")
            self.time_remaining_label.config(text=f"زمان باقیمانده: {time_str}")
        
        self.root.after(0, update)
    
    def update_status(self):
        """به روزرسانی دوره‌ای وضعیت"""
        if self.is_running and self.next_change_time:
            remaining = (self.next_change_time - datetime.now()).total_seconds()
            if remaining > 0 and self.current_window_index < len(self.selected_windows):
                current_window = self.selected_windows[self.current_window_index][1]
                self.update_status_labels(current_window, remaining)
            else:
                self.update_status_labels("-", "-")
        
        self.root.after(1000, self.update_status)

def main():
    try:
        root = tk.Tk()
        app = FullscreenWindowManager(root)
        
        # تنظیم رفتار هنگام بسته شدن پنجره
        def on_closing():
            app.stop_rotation()
            root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("خطای غیرمنتظره", f"برنامه با خطا مواجه شد:\n{str(e)}")

if __name__ == "__main__":
    # بررسی نصب بودن pywin32
    try:
        import win32api
    except ImportError:
        messagebox.showerror("خطا", "لطفاً ابتدا کتابخانه pywin32 را نصب کنید:\n\npip install pywin32")
        sys.exit()
    
    main()