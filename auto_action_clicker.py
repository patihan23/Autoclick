#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auto Action Clicker
A Python application for automating mouse clicks and keyboard presses on target windows.
Supports Windows only.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
import pyautogui
import win32gui
import win32con
import win32api
import sys


class WindowManager:
    """Handle window-related operations"""
    
    @staticmethod
    def get_open_window_titles():
        """Get all currently open and visible window titles"""
        titles = []
        
        def enum_windows_callback(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title.strip():  # Ensure title is not empty
                    titles.append(title)
            return True
        
        try:
            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
            print(f"Error getting window titles: {e}")
        
        return sorted(list(set(titles)))  # Unique and sorted
    
    @staticmethod
    def find_target_window(target_window_name):
        """Find target windows in the system (doesn't need to be focused)"""
        found_windows = []
        
        def enum_windows_callback(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if window_title and target_window_name.lower() in window_title.lower():
                    found_windows.append((hwnd, window_title))
            return True
        
        try:
            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
            print(f"Error finding window: {e}")
        
        return found_windows
    
    @staticmethod
    def is_target_window_exists(target_window_name):
        """Check if target window exists"""
        found_windows = WindowManager.find_target_window(target_window_name)
        return len(found_windows) > 0


class KeyboardHandler:
    """Handle keyboard input operations"""
    
    # Virtual Key Code mapping
    KEY_MAP = {
        'space': 0x20, 'enter': 0x0D, 'tab': 0x09, 'esc': 0x1B,
        'backspace': 0x08, 'delete': 0x2E, 'insert': 0x2D,
        'home': 0x24, 'end': 0x23, 'pageup': 0x21, 'pagedown': 0x22,
        'up': 0x26, 'down': 0x28, 'left': 0x25, 'right': 0x27,
        'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73, 'f5': 0x74,
        'f6': 0x75, 'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79,
        'f11': 0x7A, 'f12': 0x7B, 'ctrl': 0x11, 'alt': 0x12,
        'shift': 0x10, 'win': 0x5B
    }
    
    # PyAutoGUI key mapping
    PYAUTOGUI_KEY_MAP = {
        'space': 'space', 'enter': 'enter', 'tab': 'tab', 'esc': 'escape',
        'backspace': 'backspace', 'delete': 'delete', 'insert': 'insert',
        'home': 'home', 'end': 'end', 'pageup': 'pageup', 'pagedown': 'pagedown',
        'up': 'up', 'down': 'down', 'left': 'left', 'right': 'right',
        'f1': 'f1', 'f2': 'f2', 'f3': 'f3', 'f4': 'f4', 'f5': 'f5',
        'f6': 'f6', 'f7': 'f7', 'f8': 'f8', 'f9': 'f9', 'f10': 'f10',
        'f11': 'f11', 'f12': 'f12', 'ctrl': 'ctrl', 'alt': 'alt',
        'shift': 'shift', 'win': 'winleft'
    }
    
    @staticmethod
    def get_virtual_key_code(key):
        """Get virtual key code for a given key"""
        if len(key) == 1:
            if key.isalpha():
                return ord(key.upper())
            elif key.isdigit():
                return ord(key)
            else:
                return None
        else:
            return KeyboardHandler.KEY_MAP.get(key.lower())
    
    @staticmethod
    def send_key_to_window(target_window_name, key):
        """Send keyboard key to target window with multiple fallback methods"""
        found_windows = WindowManager.find_target_window(target_window_name)
        if not found_windows:
            return False, "ไม่พบหน้าต่างเป้าหมาย"
        
        hwnd = found_windows[0][0]
        vk_code = KeyboardHandler.get_virtual_key_code(key)
        
        if vk_code is None:
            return False, f"ไม่รองรับปุ่ม '{key}'"
        
        # Method 1: Try PostMessage
        try:
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
            time.sleep(0.01)
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
            return True, "ส่งปุ่มด้วย PostMessage สำเร็จ"
        except Exception as e:
            if "Access is denied" not in str(e):
                return False, f"PostMessage ผิดพลาด: {e}"
        
        # Method 2: Try SendMessage
        try:
            win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, vk_code, 0)
            time.sleep(0.01)
            win32api.SendMessage(hwnd, win32con.WM_KEYUP, vk_code, 0)
            return True, "ส่งปุ่มด้วย SendMessage สำเร็จ"
        except Exception as e:
            if "Access is denied" not in str(e):
                return False, f"SendMessage ผิดพลาด: {e}"
        
        # Method 3: Focus window and use pyautogui
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.1)
            
            if len(key) == 1:
                pyautogui.press(key)
            else:
                pyautogui_key = KeyboardHandler.PYAUTOGUI_KEY_MAP.get(key.lower(), key)
                pyautogui.press(pyautogui_key)
            
            return True, "ส่งปุ่มด้วย pyautogui สำเร็จ (หน้าต่างถูกโฟกัส)"
        except Exception as e:
            return False, f"pyautogui ผิดพลาด: {e}"


class MouseHandler:
    """Handle mouse operations"""
    
    @staticmethod
    def perform_mouse_action(x, y, button):
        """Perform mouse action at specified coordinates"""
        try:
            if button == "left":
                pyautogui.click(x, y, button='left')
            elif button == "right":
                pyautogui.click(x, y, button='right')
            elif button == "middle":
                pyautogui.click(x, y, button='middle')
            elif button == "double":
                pyautogui.doubleClick(x, y)
            
            return True, f"คลิก{button}ที่ ({x}, {y})"
        except Exception as e:
            return False, f"ข้อผิดพลาดในการคลิก: {e}"


class AutoActionClicker:
    """Main application class"""
    
    def __init__(self, root):
        self.root = root
        self._setup_window()
        self._initialize_variables()
        self._configure_pyautogui()
        self._setup_ui()
        self._start_updates()
        self._setup_event_handlers()
    
    def _setup_window(self):
        """Setup main window properties"""
        self.root.title("Auto Action Clicker v2.0")
        self.root.geometry("550x580")
        self.root.resizable(False, False)
    
    def _initialize_variables(self):
        """Initialize all class variables"""
        self.is_clicking = False
        self.click_thread = None
        self.current_mouse_pos = (0, 0)
        self.all_window_titles = []
    
    def _configure_pyautogui(self):
        """Configure pyautogui settings"""
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.1
    
    def _setup_event_handlers(self):
        """Setup event handlers"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def _start_updates(self):
        """Start periodic updates"""
        self.update_mouse_position()
    def _setup_ui(self):
        """Setup the user interface"""
        self._create_main_frame()
        self._create_window_selection_section()
        self._create_action_type_section()
        self._create_mouse_settings_section()
        self._create_keyboard_settings_section()
        self._create_delay_section()
        self._create_status_section()
        self._create_control_buttons()
        self._configure_grid_weights()
        self._initialize_ui_state()
    
    def _create_main_frame(self):
        """Create main container frame"""
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
    
    def _create_window_selection_section(self):
        """Create target window selection section"""
        ttk.Label(self.main_frame, text="ชื่อหน้าต่างเป้าหมาย:").grid(
            row=0, column=0, sticky="w", pady=(0, 5))
        
        self.window_name_var = tk.StringVar()
        window_selection_frame = ttk.Frame(self.main_frame)
        window_selection_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 15))
        
        self.window_name_combobox = ttk.Combobox(
            window_selection_frame, textvariable=self.window_name_var)
        self.window_name_combobox.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.refresh_button = ttk.Button(
            window_selection_frame, text="รีเฟรช", command=self.refresh_window_list)
        self.refresh_button.pack(side="left")
        
        # Bind events
        self.window_name_combobox.bind('<KeyRelease>', self.on_window_filter)
        self.window_name_combobox.bind('<<ComboboxSelected>>', self.on_window_select)
    
    def _create_action_type_section(self):
        """Create action type selection section"""
        ttk.Label(self.main_frame, text="ประเภทการกระทำ:").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(0, 5))
        
        self.action_type_var = tk.StringVar(value="mouse")
        action_frame = ttk.Frame(self.main_frame)
        action_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 15))
        
        ttk.Radiobutton(action_frame, text="คลิกเมาส์", variable=self.action_type_var,
                       value="mouse", command=self.on_action_type_change).pack(side="left", padx=(0, 20))
        ttk.Radiobutton(action_frame, text="กดปุ่มคีย์บอร์ด", variable=self.action_type_var,
                       value="keyboard", command=self.on_action_type_change).pack(side="left")
    
    def _create_mouse_settings_section(self):
        """Create mouse settings section"""
        self.mouse_frame = ttk.LabelFrame(self.main_frame, text="การตั้งค่าเมาส์", padding="10")
        self.mouse_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        ttk.Label(self.mouse_frame, text="ปุ่มเมาส์:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        self.mouse_button_var = tk.StringVar(value="left")
        mouse_button_frame = ttk.Frame(self.mouse_frame)
        mouse_button_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        buttons = [("ซ้าย", "left"), ("ขวา", "right"), ("กลาง", "middle"), ("ดับเบิลคลิก", "double")]
        for i, (text, value) in enumerate(buttons):
            ttk.Radiobutton(mouse_button_frame, text=text, variable=self.mouse_button_var,
                           value=value).pack(side="left", padx=(0, 10))
        
        ttk.Label(self.mouse_frame, text="ตำแหน่งเมาส์ปัจจุบัน:").grid(row=2, column=0, sticky="w", pady=(0, 5))
        self.mouse_pos_label = ttk.Label(self.mouse_frame, text="X: 0, Y: 0", foreground="blue")
        self.mouse_pos_label.grid(row=3, column=0, sticky="w")
    
    def _create_keyboard_settings_section(self):
        """Create keyboard settings section"""
        self.keyboard_frame = ttk.LabelFrame(self.main_frame, text="การตั้งค่าคีย์บอร์ด", padding="10")
        self.keyboard_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        ttk.Label(self.keyboard_frame, text="ปุ่มที่ต้องการกด:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        self.key_var = tk.StringVar()
        self.key_combobox = ttk.Combobox(self.keyboard_frame, textvariable=self.key_var,
                                        width=30, state="readonly")
        self.key_combobox.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        # Common keys list
        common_keys = [
            "space", "enter", "tab", "esc", "backspace", "delete",
            "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m",
            "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z",
            "1", "2", "3", "4", "5", "6", "7", "8", "9", "0",
            "f1", "f2", "f3", "f4", "f5", "f6", "f7", "f8", "f9", "f10", "f11", "f12",
            "ctrl", "alt", "shift", "win", "up", "down", "left", "right",
            "home", "end", "pageup", "pagedown", "insert"
        ]
        self.key_combobox['values'] = common_keys
        self.key_combobox.set("space")
        
        ttk.Label(self.keyboard_frame, text="หรือกรอกปุ่มเอง:").grid(row=2, column=0, sticky="w", pady=(5, 2))
        self.custom_key_var = tk.StringVar()
        self.custom_key_entry = ttk.Entry(self.keyboard_frame, textvariable=self.custom_key_var, width=20)
        self.custom_key_entry.grid(row=3, column=0, sticky="w")
    
    def _create_delay_section(self):
        """Create delay settings section"""
        ttk.Label(self.main_frame, text="Delay ระหว่างการกระทำ (วินาที):").grid(
            row=5, column=0, sticky="w", pady=(0, 5))
        self.delay_var = tk.StringVar(value="1.0")
        self.delay_entry = ttk.Entry(self.main_frame, textvariable=self.delay_var, width=20)
        self.delay_entry.grid(row=6, column=0, sticky="w", pady=(0, 15))
    
    def _create_status_section(self):
        """Create status display section"""
        ttk.Label(self.main_frame, text="หน้าต่างที่กำลังโฟกัส:").grid(
            row=7, column=0, sticky="w", pady=(0, 5))
        self.focused_window_label = ttk.Label(self.main_frame, text="ไม่ทราบ", foreground="green")
        self.focused_window_label.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        
        ttk.Label(self.main_frame, text="สถานะหน้าต่างเป้าหมาย:").grid(
            row=9, column=0, sticky="w", pady=(0, 5))
        self.target_window_status_label = ttk.Label(self.main_frame, text="ยังไม่ได้ตั้งค่า", foreground="gray")
        self.target_window_status_label.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(0, 15))
    
    def _create_control_buttons(self):
        """Create control buttons"""
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=11, column=0, columnspan=2, pady=20)
        
        self.start_button = ttk.Button(button_frame, text="เริ่ม Auto Action", command=self.start_action)
        self.start_button.pack(side="left", padx=(0, 10))
        
        self.stop_button = ttk.Button(button_frame, text="หยุด Auto Action", 
                                     command=self.stop_action, state="disabled")
        self.stop_button.pack(side="left")
        
        self.status_label = ttk.Label(self.main_frame, text="สถานะ: พร้อมใช้งาน", foreground="black")
        self.status_label.grid(row=12, column=0, columnspan=2, pady=(10, 0))
    
    def _configure_grid_weights(self):
        """Configure grid weights for responsive layout"""
        self.main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def _initialize_ui_state(self):
        """Initialize UI state"""
        self.on_action_type_change()
        self.refresh_window_list()    
    # Window management methods
    def refresh_window_list(self):
        """Refresh the window list in the combobox"""
        self.status_label.config(text="สถานะ: กำลังรีเฟรชรายการหน้าต่าง...", foreground="blue")
        self.root.update_idletasks()
        
        self.all_window_titles = WindowManager.get_open_window_titles()
        
        if not self.all_window_titles:
            self.window_name_combobox['values'] = []
        else:
            self.window_name_combobox['values'] = self.all_window_titles
        
        self.on_window_filter()
        self.status_label.config(text="สถานะ: พร้อมใช้งาน", foreground="black")
    
    def on_window_filter(self, event=None):
        """Filter window list in combobox while typing"""
        current_text = self.window_name_var.get()
        
        if not self.all_window_titles:
            self.window_name_combobox['values'] = []
            return
        
        if not current_text:
            self.window_name_combobox['values'] = self.all_window_titles
        else:
            filtered_titles = [title for title in self.all_window_titles 
                             if current_text.lower() in title.lower()]
            self.window_name_combobox['values'] = filtered_titles
    
    def on_window_select(self, event=None):
        """Handle window selection from combobox"""
        pass  # The textvariable handles the update automatically
    
    def on_action_type_change(self):
        """Handle action type change"""
        if self.action_type_var.get() == "mouse":
            self.mouse_frame.grid()
            self.keyboard_frame.grid_remove()
        else:
            self.mouse_frame.grid_remove()
            self.keyboard_frame.grid()
    
    # Update methods
    def update_mouse_position(self):
        """Update mouse position and window focus every 100ms"""
        try:
            # Update mouse position
            x, y = pyautogui.position()
            self.current_mouse_pos = (x, y)
            self.mouse_pos_label.config(text=f"X: {x}, Y: {y}")
            
            # Update focused window
            try:
                hwnd = win32gui.GetForegroundWindow()
                window_title = win32gui.GetWindowText(hwnd)
                if window_title:
                    self.focused_window_label.config(text=window_title)
                else:
                    self.focused_window_label.config(text="ไม่ทราบชื่อหน้าต่าง")
            except:
                self.focused_window_label.config(text="ไม่สามารถตรวจสอบได้")
            
            # Update target window status
            target_window = self.window_name_var.get().strip()
            if target_window:
                found_windows = WindowManager.find_target_window(target_window)
                if found_windows:
                    window_count = len(found_windows)
                    if window_count == 1:
                        self.target_window_status_label.config(
                            text=f"พบหน้าต่าง: {found_windows[0][1]}", foreground="green")
                    else:
                        self.target_window_status_label.config(
                            text=f"พบ {window_count} หน้าต่าง", foreground="green")
                else:
                    self.target_window_status_label.config(
                        text="ไม่พบหน้าต่างเป้าหมาย", foreground="red")
            else:
                self.target_window_status_label.config(
                    text="ยังไม่ได้ตั้งค่า", foreground="gray")
        except:
            pass
        
        self.root.after(100, self.update_mouse_position)
    
    # Action methods
    def perform_mouse_action(self):
        """Perform mouse action"""
        x, y = self.current_mouse_pos
        button = self.mouse_button_var.get()
        success, message = MouseHandler.perform_mouse_action(x, y, button)
        return message
    
    def perform_keyboard_action(self, target_window):
        """Perform keyboard action"""
        key = self.custom_key_var.get().strip() or self.key_var.get()
        
        if not key:
            return "ไม่ได้เลือกปุ่ม"
        
        success, message = KeyboardHandler.send_key_to_window(target_window, key)
        if success:
            return f"กดปุ่ม '{key}' - {message}"
        else:
            return f"ล้มเหลว: {message}"
    
    # Validation methods
    def validate_inputs(self):
        """Validate user inputs"""
        target_window = self.window_name_var.get().strip()
        if not target_window:
            messagebox.showerror("ข้อผิดพลาด", "กรุณากรอกชื่อหน้าต่างเป้าหมาย")
            return False
        
        try:
            delay = float(self.delay_var.get())
            if delay <= 0:
                messagebox.showerror("ข้อผิดพลาด", "Delay ต้องมากกว่า 0")
                return False
        except ValueError:
            messagebox.showerror("ข้อผิดพลาด", "กรุณากรอก Delay เป็นตัวเลข")
            return False
        
        if self.action_type_var.get() == "keyboard":
            key = self.custom_key_var.get().strip() or self.key_var.get()
            if not key:
                messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกปุ่มที่ต้องการกด")
                return False
        
        return True
    
    # Main action control methods
    def action_worker(self):
        """Worker function for action thread"""
        target_window = self.window_name_var.get().strip()
        delay = float(self.delay_var.get())
        action_type = self.action_type_var.get()
        
        while self.is_clicking:
            try:
                if WindowManager.is_target_window_exists(target_window):
                    if action_type == "mouse":
                        result = self.perform_mouse_action()
                    else:  # keyboard
                        result = self.perform_keyboard_action(target_window)
                    
                    self.root.after(0, lambda r=result: self.status_label.config(
                        text=f"สถานะ: {r}", foreground="green"))
                else:
                    self.root.after(0, lambda: self.status_label.config(
                        text="สถานะ: ไม่พบหน้าต่างเป้าหมาย", foreground="orange"))
                
                time.sleep(delay)
            except Exception as e:
                print(f"ข้อผิดพลาดในการทำงาน: {e}")
                time.sleep(0.1)
        
        self.root.after(0, lambda: self.status_label.config(
            text="สถานะ: หยุดการทำงาน", foreground="red"))
    
    def start_action(self):
        """Start auto action"""
        if not self.validate_inputs():
            return
        
        if self.is_clicking:
            return
        
        target_window = self.window_name_var.get().strip()
        if not WindowManager.is_target_window_exists(target_window):
            messagebox.showwarning("คำเตือน",
                f"ไม่พบหน้าต่าง '{target_window}' ในระบบ\nกรุณาตรวจสอบชื่อหน้าต่างให้ถูกต้อง")
            return
        
        if self.action_type_var.get() == "keyboard":
            result = messagebox.askyesno("คำเตือน",
                "การส่งปุ่มคีย์บอร์ดอาจไม่ทำงานกับเกมบางเกมที่มีการป้องกัน\n"
                "หากไม่ทำงาน ลองทำดังนี้:\n"
                "1. รันโปรแกรมนี้ในฐานะ Administrator\n"
                "2. ปิดโปรแกรม Anti-cheat ชั่วคราว\n"
                "3. ใช้โหมดคลิกเมาส์แทน\n\n"
                "ต้องการดำเนินการต่อหรือไม่?")
            if not result:
                return
        
        self.is_clicking = True
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        
        self.click_thread = threading.Thread(target=self.action_worker, daemon=True)
        self.click_thread.start()
        
        action_text = "คลิกเมาส์" if self.action_type_var.get() == "mouse" else "กดปุ่มคีย์บอร์ด"
        self.status_label.config(text=f"สถานะ: เริ่ม{action_text}", foreground="green")
    
    def stop_action(self):
        """Stop auto action"""
        self.is_clicking = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        if self.click_thread and self.click_thread.is_alive():
            self.click_thread.join(timeout=1.0)
        
        self.status_label.config(text="สถานะ: พร้อมใช้งาน", foreground="black")
    
    def on_closing(self):
        """Handle window closing"""
        if self.is_clicking:
            self.stop_action()
        
        if self.click_thread and self.click_thread.is_alive():
            self.click_thread.join(timeout=2.0)
        
        self.root.destroy()


def check_system_requirements():
    """Check if running on Windows"""
    if sys.platform != "win32":
        print("โปรแกรมนี้รองรับเฉพาะ Windows เท่านั้น")
        return False
    return True


def main():
    """Main application entry point"""
    if not check_system_requirements():
        return
    
    try:
        root = tk.Tk()
        app = AutoActionClicker(root)
        root.mainloop()
    except ImportError as e:
        print(f"ไม่สามารถนำเข้าไลบรารีที่จำเป็น: {e}")
        print("กรุณาติดตั้ง: pip install pyautogui pywin32")
    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")


if __name__ == "__main__":
    main()
