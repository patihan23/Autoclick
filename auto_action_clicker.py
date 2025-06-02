#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Auto Action Clicker v3.0 - Premium Enhanced Edition (Performance Optimized)
A Python application for automating mouse clicks and keyboard presses on target windows.
Supports Windows only.

Created by Patihan
Performance Optimizations Applied:
- Reduced pyautogui pause from 0.1s to 0.05s
- Adaptive mouse position update frequency
- Cached window operations
- Optimized UI updates
- Efficient threading patterns
- Memory usage optimizations
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import datetime, timedelta
import webbrowser
import os
import json
import time
import threading
import logging
from functools import lru_cache

# Attempt to import critical dependencies and show error if missing
try:
    import pyautogui
    # Performance optimization: Reduce default pause
    pyautogui.PAUSE = 0.05  # Reduced from default 0.1s for better performance
    pyautogui.FAILSAFE = True  # Keep failsafe enabled for safety
except ImportError:
    print("CRITICAL ERROR: PyAutoGUI is not installed. Please install it via: pip install pyautogui")
    exit(1)

try:
    import keyboard
except ImportError:
    print("CRITICAL ERROR: keyboard module is not installed. Please install it via: pip install keyboard")
    exit(1)

try:
    import win32gui
    import win32api
    import win32con
except ImportError:
    print("CRITICAL ERROR: pywin32 is not installed. Please install it via: pip install pywin32")
    exit(1)

# Optional theme support
try:
    from ttkthemes import ThemedTk, ThemedStyle
    THEMES_AVAILABLE = True
except ImportError:
    # Provide fallback classes to avoid unbound variable warnings
    ThemedTk = None
    ThemedStyle = None
    THEMES_AVAILABLE = False

# Setup logging
logging.basicConfig(
    filename='autoclick.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


class PerformanceMonitor:
    """Monitor and optimize application performance"""
    
    def __init__(self):
        self.last_update_time = time.time()
        self.update_count = 0
        self.adaptive_interval = 100  # Start with 100ms
        
    def should_update(self):
        """Adaptive update frequency based on activity"""
        current_time = time.time()
        if current_time - self.last_update_time >= self.adaptive_interval / 1000:
            self.last_update_time = current_time
            self.update_count += 1
            return True
        return False
    
    def adjust_frequency(self, is_active):
        """Adjust update frequency based on activity"""
        if is_active:
            self.adaptive_interval = max(50, self.adaptive_interval - 10)  # Increase frequency when active
        else:
            self.adaptive_interval = min(200, self.adaptive_interval + 5)  # Decrease when idle


class WindowManager:
    """Handle window-related operations with caching for performance"""
    
    def __init__(self):
        self._window_cache = {}
        self._cache_time = 0
        self._cache_timeout = 2.0  # Cache for 2 seconds
    
    @lru_cache(maxsize=128)
    def get_open_window_titles(self):
        """Get all currently open and visible window titles (cached)"""
        current_time = time.time()
        if current_time - self._cache_time < self._cache_timeout and self._window_cache:
            return self._window_cache.get('titles', [])
        
        titles = []
        
        def enum_windows_callback(hwnd, lParam):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title.strip():
                    titles.append(title)
            return True
        
        try:
            win32gui.EnumWindows(enum_windows_callback, None)
            unique_titles = sorted(list(set(titles)))
            self._window_cache['titles'] = unique_titles
            self._cache_time = current_time
            return unique_titles
        except Exception as e:
            logging.error(f"Error getting window titles: {e}")
            return []
    
    def find_target_window(self, target_window_name):
        """Find target windows in the system"""
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
            logging.error(f"Error finding window: {e}")
        
        return found_windows
    
    def clear_cache(self):
        """Clear the window cache"""
        self._window_cache.clear()
        self.get_open_window_titles.cache_clear()


class KeyboardHandler:
    """Handle keyboard input efficiently"""
    
    def __init__(self):
        self.registered_hotkeys = set()
    
    def register_hotkey(self, hotkey, callback):
        """Register a hotkey with error handling"""
        try:
            if hotkey in self.registered_hotkeys:
                keyboard.remove_hotkey(hotkey)
            keyboard.add_hotkey(hotkey, callback)
            self.registered_hotkeys.add(hotkey)
            return True
        except Exception as e:
            logging.error(f"Failed to register hotkey {hotkey}: {e}")
            return False
    
    def unregister_all(self):
        """Unregister all hotkeys"""
        for hotkey in self.registered_hotkeys.copy():
            try:
                keyboard.remove_hotkey(hotkey)
                self.registered_hotkeys.remove(hotkey)
            except Exception as e:
                logging.error(f"Failed to unregister hotkey {hotkey}: {e}")


class MouseHandler:
    """Handle mouse operations with optimizations"""
    
    def __init__(self):
        self.last_position = None
        self.position_cache_time = 0
        self.cache_duration = 0.05  # Cache position for 50ms
    
    def get_mouse_position(self, force_update=False):
        """Get mouse position with caching"""
        current_time = time.time()
        if not force_update and self.last_position and (current_time - self.position_cache_time) < self.cache_duration:
            return self.last_position
        
        try:
            position = pyautogui.position()
            self.last_position = position
            self.position_cache_time = current_time
            return position
        except Exception as e:
            logging.error(f"Error getting mouse position: {e}")
            return (0, 0)
    
    def click_at_position(self, x, y, button='left', clicks=1):
        """Optimized click operation"""
        try:
            pyautogui.click(x, y, clicks=clicks, button=button)
            return True
        except Exception as e:
            logging.error(f"Error clicking at ({x}, {y}): {e}")
            return False


class AutoActionClicker:
    """Main application class with performance optimizations"""
    
    def __init__(self):
        # Initialize components
        self.performance_monitor = PerformanceMonitor()
        self.window_manager = WindowManager()
        self.keyboard_handler = KeyboardHandler()
        self.mouse_handler = MouseHandler()
        
        # Application state
        self.is_clicking = False
        self.is_recording_macro = False
        self.recorded_actions = []
        self.click_count = 0
        self.start_time = None
        
        # Configuration
        self.config_file = "autoclick_config.json"
        self.default_config = {
            'click_interval': 1.0,
            'action_type': 'mouse',
            'mouse_button': 'left',
            'click_type': 'single',
            'x_coordinate': 100,
            'y_coordinate': 100,
            'keyboard_key': 'space',
            'target_window': '',
            'hotkey_start_stop': 'f6',
            'emergency_stop_hotkey': 'f12',
            'auto_resize_window': True,
            'theme': 'arc'
        }
        
        # Load configuration
        self.config = self.load_config()
          # Setup GUI
        self._setup_window()
        self._create_widgets()
        self._setup_event_handlers()
        self._setup_hotkeys()
        
        # Start performance monitoring
        self._start_performance_monitoring()
    
    def _setup_window(self):
        """Setup main window properties"""
        if THEMES_AVAILABLE and ThemedTk is not None:
            self.root = ThemedTk(theme=self.config.get('theme', 'arc'))
        else:
            self.root = tk.Tk()
        
        self.root.title("Auto Action Clicker v3.0 - Performance Optimized")
        self.root.geometry("600x700")
        self.root.resizable(True, True)
        
        # Icon setup (if available)
        try:
            if os.path.exists("icon.ico"):
                self.root.iconbitmap("icon.ico")
        except Exception:
            pass
    
    def _create_widgets(self):
        """Create all GUI widgets"""
        # Main notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Create tabs
        self._create_main_tab()
        self._create_settings_tab()
        self._create_macro_tab()
        self._create_about_tab()
    
    def _create_main_tab(self):
        """Create main control tab"""
        main_frame = ttk.Frame(self.notebook)
        self.notebook.add(main_frame, text="Main Controls")
        
        # Target window selection
        window_frame = ttk.LabelFrame(main_frame, text="Target Window", padding=10)
        window_frame.pack(fill="x", padx=5, pady=5)
        
        self.target_window_var = tk.StringVar(value=self.config.get('target_window', ''))
        self.window_combo = ttk.Combobox(window_frame, textvariable=self.target_window_var, 
                                        state="readonly", width=50)
        self.window_combo.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ttk.Button(window_frame, text="Refresh", 
                  command=self.refresh_windows).pack(side="right")
        
        # Action type selection
        action_frame = ttk.LabelFrame(main_frame, text="Action Type", padding=10)
        action_frame.pack(fill="x", padx=5, pady=5)
        
        self.action_type_var = tk.StringVar(value=self.config.get('action_type', 'mouse'))
        ttk.Radiobutton(action_frame, text="Mouse Click", variable=self.action_type_var,
                       value="mouse", command=self.on_action_type_change).pack(side="left")
        ttk.Radiobutton(action_frame, text="Keyboard Press", variable=self.action_type_var,
                       value="keyboard", command=self.on_action_type_change).pack(side="left")
        
        # Settings container
        self.settings_container = ttk.Frame(main_frame)
        self.settings_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Create mouse/keyboard settings sections
        self._create_mouse_settings_section()
        self._create_keyboard_settings_section()
        
        # Interval setting
        interval_frame = ttk.LabelFrame(main_frame, text="Click Interval", padding=10)
        interval_frame.pack(fill="x", padx=5, pady=5)
        
        self.interval_var = tk.DoubleVar(value=self.config.get('click_interval', 1.0))
        interval_scale = ttk.Scale(interval_frame, from_=0.1, to=10.0, 
                                  variable=self.interval_var, orient="horizontal")
        interval_scale.pack(fill="x", padx=(0, 10), side="left", expand=True)
        
        self.interval_label = ttk.Label(interval_frame, text=f"{self.interval_var.get():.1f}s")
        self.interval_label.pack(side="right")
        interval_scale.configure(command=self.update_interval_label)
        
        # Control buttons
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill="x", padx=5, pady=10)
        
        self.start_button = ttk.Button(control_frame, text="Start (F6)", 
                                      command=self.toggle_clicking, style="Accent.TButton")
        self.start_button.pack(side="left", padx=5)
        
        self.emergency_button = ttk.Button(control_frame, text="Emergency Stop (F12)", 
                                          command=self.emergency_stop, style="TButton")
        self.emergency_button.pack(side="left", padx=5)
        
        # Status section
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding=10)
        status_frame.pack(fill="x", padx=5, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="Ready", font=("Arial", 10))
        self.status_label.pack()
        
        self.stats_label = ttk.Label(status_frame, text="Clicks: 0 | Time: 0s", 
                                    font=("Arial", 8))
        self.stats_label.pack()
        
        # Mouse position display
        self.position_label = ttk.Label(status_frame, text="Mouse: (0, 0)", 
                                       font=("Arial", 8))
        self.position_label.pack()
        
        # Show appropriate settings
        self.on_action_type_change()
    
    def _create_mouse_settings_section(self):
        """Create mouse settings section"""
        self.mouse_frame = ttk.LabelFrame(self.settings_container, text="Mouse Settings", padding=10)
        
        # Button selection
        button_frame = ttk.Frame(self.mouse_frame)
        button_frame.pack(fill="x", pady=5)
        
        ttk.Label(button_frame, text="Button:").pack(side="left", padx=(0, 5))
        self.mouse_button_var = tk.StringVar(value=self.config.get('mouse_button', 'left'))
        button_combo = ttk.Combobox(button_frame, textvariable=self.mouse_button_var,
                                   values=['left', 'right', 'middle'], state="readonly", width=10)
        button_combo.pack(side="left", padx=5)
        
        # Click type
        ttk.Label(button_frame, text="Type:").pack(side="left", padx=(20, 5))
        self.click_type_var = tk.StringVar(value=self.config.get('click_type', 'single'))
        type_combo = ttk.Combobox(button_frame, textvariable=self.click_type_var,
                                 values=['single', 'double'], state="readonly", width=10)
        type_combo.pack(side="left", padx=5)
        
        # Coordinates
        coord_frame = ttk.Frame(self.mouse_frame)
        coord_frame.pack(fill="x", pady=5)
        
        ttk.Label(coord_frame, text="X:").pack(side="left")
        self.x_var = tk.IntVar(value=self.config.get('x_coordinate', 100))
        x_spin = ttk.Spinbox(coord_frame, from_=0, to=9999, textvariable=self.x_var, width=8)
        x_spin.pack(side="left", padx=5)
        
        ttk.Label(coord_frame, text="Y:").pack(side="left", padx=(20, 0))
        self.y_var = tk.IntVar(value=self.config.get('y_coordinate', 100))
        y_spin = ttk.Spinbox(coord_frame, from_=0, to=9999, textvariable=self.y_var, width=8)
        y_spin.pack(side="left", padx=5)
        
        ttk.Button(coord_frame, text="Get Current Position", 
                  command=self.get_current_mouse_position).pack(side="left", padx=20)
    
    def _create_keyboard_settings_section(self):
        """Create keyboard settings section"""
        self.keyboard_frame = ttk.LabelFrame(self.settings_container, text="Keyboard Settings", padding=10)
        
        # Key selection
        key_frame = ttk.Frame(self.keyboard_frame)
        key_frame.pack(fill="x", pady=5)
        
        ttk.Label(key_frame, text="Key to press:").pack(side="left", padx=(0, 5))
        self.keyboard_key_var = tk.StringVar(value=self.config.get('keyboard_key', 'space'))
        key_entry = ttk.Entry(key_frame, textvariable=self.keyboard_key_var, width=20)
        key_entry.pack(side="left", padx=5)
        
        # Common keys
        common_frame = ttk.Frame(self.keyboard_frame)
        common_frame.pack(fill="x", pady=5)
        
        ttk.Label(common_frame, text="Common keys:").pack(side="left", padx=(0, 5))
        for key in ['space', 'enter', 'tab', 'esc', 'f1', 'f2', 'f3']:
            ttk.Button(common_frame, text=key, width=6,
                      command=lambda k=key: self.keyboard_key_var.set(k)).pack(side="left", padx=2)
    
    def _create_settings_tab(self):
        """Create settings tab"""
        settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(settings_frame, text="Settings")
        
        # Hotkeys section
        hotkey_frame = ttk.LabelFrame(settings_frame, text="Hotkeys", padding=10)
        hotkey_frame.pack(fill="x", padx=5, pady=5)
        
        # Start/Stop hotkey
        start_hotkey_frame = ttk.Frame(hotkey_frame)
        start_hotkey_frame.pack(fill="x", pady=2)
        
        ttk.Label(start_hotkey_frame, text="Start/Stop Hotkey:").pack(side="left")
        self.start_hotkey_var = tk.StringVar(value=self.config.get('hotkey_start_stop', 'f6'))
        start_hotkey_entry = ttk.Entry(start_hotkey_frame, textvariable=self.start_hotkey_var, width=15)
        start_hotkey_entry.pack(side="left", padx=10)
        
        # Emergency stop hotkey
        emergency_hotkey_frame = ttk.Frame(hotkey_frame)
        emergency_hotkey_frame.pack(fill="x", pady=2)
        
        ttk.Label(emergency_hotkey_frame, text="Emergency Stop Hotkey:").pack(side="left")
        self.emergency_hotkey_var = tk.StringVar(value=self.config.get('emergency_stop_hotkey', 'f12'))
        emergency_hotkey_entry = ttk.Entry(emergency_hotkey_frame, textvariable=self.emergency_hotkey_var, width=15)
        emergency_hotkey_entry.pack(side="left", padx=10)
        
        # Window options
        window_options_frame = ttk.LabelFrame(settings_frame, text="Window Options", padding=10)
        window_options_frame.pack(fill="x", padx=5, pady=5)
        
        self.auto_resize_var = tk.BooleanVar(value=self.config.get('auto_resize_window', True))
        ttk.Checkbutton(window_options_frame, text="Auto-resize target window",
                       variable=self.auto_resize_var).pack(anchor="w")
        
        # Theme selection (if available)
        if THEMES_AVAILABLE:
            theme_frame = ttk.LabelFrame(settings_frame, text="Theme", padding=10)
            theme_frame.pack(fill="x", padx=5, pady=5)
            
            self.theme_var = tk.StringVar(value=self.config.get('theme', 'arc'))
            theme_combo = ttk.Combobox(theme_frame, textvariable=self.theme_var,
                                      values=['arc', 'equilux', 'adapta', 'breeze'], 
                                      state="readonly")
            theme_combo.pack(side="left", padx=5)
            ttk.Button(theme_frame, text="Apply Theme", 
                      command=self.apply_theme).pack(side="left", padx=10)
        
        # Save/Load configuration
        config_frame = ttk.LabelFrame(settings_frame, text="Configuration", padding=10)
        config_frame.pack(fill="x", padx=5, pady=5)
        
        config_buttons_frame = ttk.Frame(config_frame)
        config_buttons_frame.pack(fill="x")
        
        ttk.Button(config_buttons_frame, text="Save Config", 
                  command=self.save_config).pack(side="left", padx=5)
        ttk.Button(config_buttons_frame, text="Load Config", 
                  command=self.load_config_file).pack(side="left", padx=5)
        ttk.Button(config_buttons_frame, text="Reset to Default", 
                  command=self.reset_config).pack(side="left", padx=5)
    
    def _create_macro_tab(self):
        """Create macro recording tab"""
        macro_frame = ttk.Frame(self.notebook)
        self.notebook.add(macro_frame, text="Macro Recorder")
        
        # Recording controls
        record_frame = ttk.LabelFrame(macro_frame, text="Recording Controls", padding=10)
        record_frame.pack(fill="x", padx=5, pady=5)
        
        record_buttons_frame = ttk.Frame(record_frame)
        record_buttons_frame.pack(fill="x")
        
        self.record_button = ttk.Button(record_buttons_frame, text="Start Recording", 
                                       command=self.toggle_macro_recording)
        self.record_button.pack(side="left", padx=5)
        
        ttk.Button(record_buttons_frame, text="Clear Macro", 
                  command=self.clear_macro).pack(side="left", padx=5)
        
        self.play_macro_button = ttk.Button(record_buttons_frame, text="Play Macro", 
                                           command=self.play_macro, state="disabled")
        self.play_macro_button.pack(side="left", padx=5)
        
        # Macro display
        macro_display_frame = ttk.LabelFrame(macro_frame, text="Recorded Actions", padding=10)
        macro_display_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.macro_text = tk.Text(macro_display_frame, height=15, state="disabled")
        macro_scrollbar = ttk.Scrollbar(macro_display_frame, orient="vertical", command=self.macro_text.yview)
        self.macro_text.configure(yscrollcommand=macro_scrollbar.set)
        
        self.macro_text.pack(side="left", fill="both", expand=True)
        macro_scrollbar.pack(side="right", fill="y")
        
        # Save/Load macros
        macro_file_frame = ttk.Frame(macro_frame)
        macro_file_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(macro_file_frame, text="Save Macro", 
                  command=self.save_macro).pack(side="left", padx=5)
        ttk.Button(macro_file_frame, text="Load Macro", 
                  command=self.load_macro).pack(side="left", padx=5)
    
    def _create_about_tab(self):
        """Create about tab"""
        about_frame = ttk.Frame(self.notebook)
        self.notebook.add(about_frame, text="About")
        
        # Title
        title_label = ttk.Label(about_frame, text="Auto Action Clicker v3.0", 
                               font=("Arial", 16, "bold"))
        title_label.pack(pady=20)
        
        subtitle_label = ttk.Label(about_frame, text="Performance Optimized Edition", 
                                  font=("Arial", 12))
        subtitle_label.pack(pady=5)
        
        # Description
        desc_text = """
A powerful automation tool for Windows that can:
• Automate mouse clicks and keyboard presses
• Target specific windows
• Record and playback macros
• Use customizable hotkeys
• Monitor performance

Performance Features:
• Adaptive update frequency
• Cached window operations
• Optimized threading
• Reduced input lag
• Memory usage optimization
        """
        
        desc_label = ttk.Label(about_frame, text=desc_text, justify="left")
        desc_label.pack(pady=20, padx=20)
        
        # Credits
        credits_label = ttk.Label(about_frame, text="Created by Patihan", 
                                 font=("Arial", 10))
        credits_label.pack(pady=10)
    
    def _setup_event_handlers(self):
        """Setup event handlers"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Bind window selection change
        self.window_combo.bind("<<ComboboxSelected>>", self.on_window_selection_change)
        
        # Refresh windows on startup
        self.root.after(500, self.refresh_windows)
    
    def _setup_hotkeys(self):
        """Setup hotkeys with error handling"""
        start_hotkey = self.config.get('hotkey_start_stop', 'f6')
        emergency_hotkey = self.config.get('emergency_stop_hotkey', 'f12')
        
        success1 = self.keyboard_handler.register_hotkey(start_hotkey, self.toggle_clicking)
        success2 = self.keyboard_handler.register_hotkey(emergency_hotkey, self.emergency_stop)
        
        if not success1 or not success2:
            messagebox.showwarning("Hotkey Warning", 
                                 "Some hotkeys could not be registered. They may be in use by another application.")
    
    def _start_performance_monitoring(self):
        """Start performance monitoring thread"""
        self._update_mouse_position()
        self._update_statistics()
    
    def _update_mouse_position(self):
        """Update mouse position display with adaptive frequency"""
        if self.performance_monitor.should_update():
            try:
                x, y = self.mouse_handler.get_mouse_position()
                self.position_label.config(text=f"Mouse: ({x}, {y})")
                
                # Adjust frequency based on clicking state
                self.performance_monitor.adjust_frequency(self.is_clicking)
            except Exception as e:
                logging.error(f"Error updating mouse position: {e}")
        
        # Schedule next update
        self.root.after(50, self._update_mouse_position)
    
    def _update_statistics(self):
        """Update statistics display"""
        if self.is_clicking and self.start_time:
            elapsed = time.time() - self.start_time
            self.stats_label.config(text=f"Clicks: {self.click_count} | Time: {elapsed:.0f}s")
        
        self.root.after(1000, self._update_statistics)
    
    def refresh_windows(self):
        """Refresh the list of available windows"""
        try:
            self.window_manager.clear_cache()
            windows = self.window_manager.get_open_window_titles()
            self.window_combo['values'] = windows
            
            current_value = self.target_window_var.get()
            if current_value not in windows and windows:
                self.target_window_var.set(windows[0])
                
        except Exception as e:
            logging.error(f"Error refreshing windows: {e}")
            messagebox.showerror("Error", f"Failed to refresh windows: {e}")
    
    def on_action_type_change(self):
        """Handle action type change"""
        action_type = self.action_type_var.get()
        
        # Hide all frames first
        self.mouse_frame.pack_forget()
        self.keyboard_frame.pack_forget()
        
        # Show appropriate frame
        if action_type == "mouse":
            self.mouse_frame.pack(fill="x", pady=5)
        else:
            self.keyboard_frame.pack(fill="x", pady=5)
    
    def on_window_selection_change(self, event=None):
        """Handle window selection change"""
        if self.auto_resize_var.get():
            self.auto_resize_target_window()
    
    def update_interval_label(self, value):
        """Update interval label"""
        self.interval_label.config(text=f"{float(value):.1f}s")
    
    def get_current_mouse_position(self):
        """Get current mouse position and set coordinates"""
        x, y = self.mouse_handler.get_mouse_position(force_update=True)
        self.x_var.set(x)
        self.y_var.set(y)
        messagebox.showinfo("Position Set", f"Coordinates set to ({x}, {y})")
    
    def auto_resize_target_window(self):
        """Auto-resize target window for better visibility"""
        target_window = self.target_window_var.get()
        if not target_window:
            return
        
        try:
            windows = self.window_manager.find_target_window(target_window)
            if windows:
                hwnd = windows[0][0]
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            logging.error(f"Error auto-resizing window: {e}")
    
    def toggle_clicking(self):
        """Toggle clicking state"""
        if self.is_clicking:
            self.stop_clicking()
        else:
            self.start_clicking()
    
    def start_clicking(self):
        """Start clicking"""
        if self.is_clicking:
            return
        
        self.is_clicking = True
        self.click_count = 0
        self.start_time = time.time()
        
        self.start_button.config(text="Stop (F6)")
        self.update_status("Running", "green")
        
        # Start worker thread
        self.worker_thread = threading.Thread(target=self._action_worker, daemon=True)
        self.worker_thread.start()
    
    def stop_clicking(self):
        """Stop clicking"""
        self.is_clicking = False
        
        self.start_button.config(text="Start (F6)")
        self.update_status("Stopped", "red")
    
    def emergency_stop(self):
        """Emergency stop all actions"""
        self.is_clicking = False
        self.is_recording_macro = False
        
        self.start_button.config(text="Start (F6)")
        self.record_button.config(text="Start Recording")
        self.update_status("Emergency Stop!", "red")
        
        messagebox.showwarning("Emergency Stop", "All actions have been stopped!")
    
    def _action_worker(self):
        """Worker thread for performing actions"""
        while self.is_clicking:
            try:
                if self.action_type_var.get() == "mouse":
                    self.perform_mouse_action()
                else:
                    self.perform_keyboard_action()
                
                self.click_count += 1
                
                # Sleep for the specified interval
                time.sleep(self.interval_var.get())
                
            except Exception as e:
                logging.error(f"Error in action worker: {e}")
                self.root.after(0, lambda: self.update_status(f"Error: {e}", "red"))
                break
    
    def perform_mouse_action(self):
        """Perform mouse click action"""
        target_window = self.target_window_var.get()
        
        # Focus target window if specified
        if target_window:
            windows = self.window_manager.find_target_window(target_window)
            if windows:
                try:
                    hwnd = windows[0][0]
                    win32gui.SetForegroundWindow(hwnd)
                except Exception as e:
                    logging.error(f"Error focusing window: {e}")
        
        # Perform click
        x = self.x_var.get()
        y = self.y_var.get()
        button = self.mouse_button_var.get()
        clicks = 2 if self.click_type_var.get() == "double" else 1
        
        success = self.mouse_handler.click_at_position(x, y, button, clicks)
        
        if not success:
            self.root.after(0, lambda: self.update_status("Click failed", "red"))
    
    def perform_keyboard_action(self):
        """Perform keyboard press action"""
        target_window = self.target_window_var.get()
        
        # Focus target window if specified
        if target_window:
            windows = self.window_manager.find_target_window(target_window)
            if windows:
                try:
                    hwnd = windows[0][0]
                    win32gui.SetForegroundWindow(hwnd)
                except Exception as e:
                    logging.error(f"Error focusing window: {e}")
        
        # Perform key press
        key = self.keyboard_key_var.get()
        try:
            pyautogui.press(key)
        except Exception as e:
            logging.error(f"Error pressing key {key}: {e}")
            self.root.after(0, lambda: self.update_status(f"Key press failed: {e}", "red"))
    
    def toggle_macro_recording(self):
        """Toggle macro recording"""
        if self.is_recording_macro:
            self.stop_macro_recording()
        else:
            self.start_macro_recording()
    
    def start_macro_recording(self):
        """Start macro recording"""
        self.is_recording_macro = True
        self.recorded_actions = []
        
        self.record_button.config(text="Stop Recording")
        self.update_status("Recording macro...", "blue")
        
        # Start recording thread
        self.record_thread = threading.Thread(target=self._macro_recorder, daemon=True)
        self.record_thread.start()
    
    def stop_macro_recording(self):
        """Stop macro recording"""
        self.is_recording_macro = False
        
        self.record_button.config(text="Start Recording")
        self.update_status("Recording stopped", "orange")
        
        if self.recorded_actions:
            self.play_macro_button.config(state="normal")
            self.display_recorded_actions()
    
    def _macro_recorder(self):
        """Record macro actions"""
        last_time = time.time()
        
        while self.is_recording_macro:
            try:
                # Record mouse position changes
                current_pos = self.mouse_handler.get_mouse_position(force_update=True)
                current_time = time.time()
                
                # Record significant position changes
                if len(self.recorded_actions) == 0 or \
                   abs(current_pos[0] - self.recorded_actions[-1].get('x', 0)) > 5 or \
                   abs(current_pos[1] - self.recorded_actions[-1].get('y', 0)) > 5:
                    
                    action = {
                        'type': 'move',
                        'x': current_pos[0],
                        'y': current_pos[1],
                        'delay': current_time - last_time
                    }
                    self.recorded_actions.append(action)
                    last_time = current_time
                
                time.sleep(0.1)
                
            except Exception as e:
                logging.error(f"Error recording macro: {e}")
                break
    
    def display_recorded_actions(self):
        """Display recorded actions in the text widget"""
        self.macro_text.config(state="normal")
        self.macro_text.delete(1.0, tk.END)
        
        for i, action in enumerate(self.recorded_actions):
            if action['type'] == 'move':
                text = f"{i+1}. Move to ({action['x']}, {action['y']}) - Delay: {action['delay']:.2f}s\n"
            elif action['type'] == 'click':
                text = f"{i+1}. Click at ({action['x']}, {action['y']}) - Button: {action['button']}\n"
            else:
                text = f"{i+1}. {action['type']}: {action}\n"
            
            self.macro_text.insert(tk.END, text)
        
        self.macro_text.config(state="disabled")
    
    def clear_macro(self):
        """Clear recorded macro"""
        self.recorded_actions = []
        self.play_macro_button.config(state="disabled")
        
        self.macro_text.config(state="normal")
        self.macro_text.delete(1.0, tk.END)
        self.macro_text.config(state="disabled")
        
        self.update_status("Macro cleared", "orange")
    
    def play_macro(self):
        """Play recorded macro"""
        if not self.recorded_actions:
            messagebox.showwarning("No Macro", "No macro has been recorded yet.")
            return
        
        self.update_status("Playing macro...", "blue")
        
        # Start playback thread
        playback_thread = threading.Thread(target=self._play_macro_worker, daemon=True)
        playback_thread.start()
    
    def _play_macro_worker(self):
        """Worker thread for macro playback"""
        try:
            for action in self.recorded_actions:
                if action['type'] == 'move':
                    pyautogui.moveTo(action['x'], action['y'])
                    time.sleep(action.get('delay', 0.1))
                elif action['type'] == 'click':
                    pyautogui.click(action['x'], action['y'], button=action.get('button', 'left'))
                
            self.root.after(0, lambda: self.update_status("Macro playback complete", "green"))
            
        except Exception as e:
            logging.error(f"Error playing macro: {e}")
            self.root.after(0, lambda: self.update_status(f"Macro error: {e}", "red"))
    
    def save_macro(self):
        """Save macro to file"""
        if not self.recorded_actions:
            messagebox.showwarning("No Macro", "No macro to save.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    json.dump(self.recorded_actions, f, indent=2)
                messagebox.showinfo("Success", f"Macro saved to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save macro: {e}")
    
    def load_macro(self):
        """Load macro from file"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    self.recorded_actions = json.load(f)
                self.display_recorded_actions()
                self.play_macro_button.config(state="normal")
                messagebox.showinfo("Success", f"Macro loaded from {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load macro: {e}")
    
    def apply_theme(self):
        """Apply selected theme"""
        if THEMES_AVAILABLE and ThemedStyle is not None:
            try:
                theme = self.theme_var.get()
                style = ThemedStyle(self.root)
                style.set_theme(theme)
                self.config['theme'] = theme
                messagebox.showinfo("Theme Applied", f"Theme '{theme}' has been applied.")
            except Exception as e:
                messagebox.showerror("Theme Error", f"Failed to apply theme: {e}")
        else:
            messagebox.showwarning("Theme Unavailable", "Theme support is not available. Please install ttkthemes: pip install ttkthemes")
    
    def save_config(self):
        """Save current configuration"""
        try:
            self.config.update({
                'click_interval': self.interval_var.get(),
                'action_type': self.action_type_var.get(),
                'mouse_button': self.mouse_button_var.get(),
                'click_type': self.click_type_var.get(),
                'x_coordinate': self.x_var.get(),
                'y_coordinate': self.y_var.get(),
                'keyboard_key': self.keyboard_key_var.get(),
                'target_window': self.target_window_var.get(),
                'hotkey_start_stop': self.start_hotkey_var.get(),
                'emergency_stop_hotkey': self.emergency_hotkey_var.get(),
                'auto_resize_window': self.auto_resize_var.get(),
            })
            
            if hasattr(self, 'theme_var'):
                self.config['theme'] = self.theme_var.get()
            
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=2)
            
            messagebox.showinfo("Success", "Configuration saved successfully!")
            
        except Exception as e:
            logging.error(f"Error saving configuration: {e}")
            messagebox.showerror("Error", f"Failed to save configuration: {e}")
    
    def load_config(self):
        """Load configuration from file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    loaded_config = json.load(f)
                
                # Merge with defaults
                for key, value in self.default_config.items():
                    if key not in loaded_config:
                        loaded_config[key] = value
                
                return loaded_config
            else:
                return self.default_config.copy()
                
        except Exception as e:
            logging.error(f"Error loading configuration: {e}")
            return self.default_config.copy()
    
    def load_config_file(self):
        """Load configuration from a selected file"""
        filename = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    loaded_config = json.load(f)
                
                # Update variables
                self.interval_var.set(loaded_config.get('click_interval', 1.0))
                self.action_type_var.set(loaded_config.get('action_type', 'mouse'))
                self.mouse_button_var.set(loaded_config.get('mouse_button', 'left'))
                self.click_type_var.set(loaded_config.get('click_type', 'single'))
                self.x_var.set(loaded_config.get('x_coordinate', 100))
                self.y_var.set(loaded_config.get('y_coordinate', 100))
                self.keyboard_key_var.set(loaded_config.get('keyboard_key', 'space'))
                self.target_window_var.set(loaded_config.get('target_window', ''))
                self.start_hotkey_var.set(loaded_config.get('hotkey_start_stop', 'f6'))
                self.emergency_hotkey_var.set(loaded_config.get('emergency_stop_hotkey', 'f12'))
                self.auto_resize_var.set(loaded_config.get('auto_resize_window', True))
                
                if hasattr(self, 'theme_var'):
                    self.theme_var.set(loaded_config.get('theme', 'arc'))
                
                self.config = loaded_config
                self.on_action_type_change()
                
                messagebox.showinfo("Success", f"Configuration loaded from {filename}")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load configuration: {e}")
    
    def reset_config(self):
        """Reset configuration to defaults"""
        if messagebox.askyesno("Reset Configuration", 
                              "Are you sure you want to reset all settings to default?"):
            self.config = self.default_config.copy()
            
            # Update all variables
            self.interval_var.set(self.config['click_interval'])
            self.action_type_var.set(self.config['action_type'])
            self.mouse_button_var.set(self.config['mouse_button'])
            self.click_type_var.set(self.config['click_type'])
            self.x_var.set(self.config['x_coordinate'])
            self.y_var.set(self.config['y_coordinate'])
            self.keyboard_key_var.set(self.config['keyboard_key'])
            self.target_window_var.set(self.config['target_window'])
            self.start_hotkey_var.set(self.config['hotkey_start_stop'])
            self.emergency_hotkey_var.set(self.config['emergency_stop_hotkey'])
            self.auto_resize_var.set(self.config['auto_resize_window'])
            
            if hasattr(self, 'theme_var'):
                self.theme_var.set(self.config['theme'])
            
            self.on_action_type_change()
            messagebox.showinfo("Reset Complete", "Configuration has been reset to defaults.")
    
    def update_status(self, text, color=None):
        """Update status label with color"""
        actual_color = color if color else "black"
        try:
            self.status_label.config(text=text, foreground=actual_color)
        except Exception as e:
            logging.error(f"Error updating status: {e}")
    
    def on_closing(self):
        """Handle application closing"""
        try:
            # Stop all actions
            self.is_clicking = False
            self.is_recording_macro = False
            
            # Unregister hotkeys
            self.keyboard_handler.unregister_all()
            
            # Save configuration
            self.save_config()
            
            # Close application
            self.root.destroy()
            
        except Exception as e:
            logging.error(f"Error during closing: {e}")
            self.root.destroy()
    
    def run(self):
        """Start the application"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.on_closing()


def main():
    """Main entry point"""
    try:
        # Check if running on Windows
        if os.name != 'nt':
            print("This application is designed for Windows only.")
            return
        
        # Create and run application
        app = AutoActionClicker()
        app.run()
        
    except Exception as e:
        logging.error(f"Fatal error: {e}", exc_info=True)
        try:
            messagebox.showerror("Fatal Error", f"Application encountered a fatal error: {e}")
        except:
            print(f"Fatal error: {e}")


if __name__ == "__main__":
    main()
