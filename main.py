import json
import os
import re
import tkinter as tk
import datetime
import webbrowser

import psutil
from tkinter import ttk  # Import themed Tkinter for better widgets
import win32gui
import win32process
import win32api
import win32con
from PIL import Image, ImageTk


def format_duration(seconds):
    days, seconds = divmod(int(seconds), 86400)  # 86400 seconds in a day
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)
    return f"{days}d {hours}h {minutes}m {seconds}s"


def get_active_window():
    window = win32gui.GetForegroundWindow()
    _, process_id = win32process.GetWindowThreadProcessId(window)
    try:
        process = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, process_id)
        executable_path = win32process.GetModuleFileNameEx(process, 0)
        win32api.CloseHandle(process)
        process_name = psutil.Process(process_id).name()  # Get the process name
        window_title = win32gui.GetWindowText(window)
        return ("Desktop", None, "Desktop") if not window_title else (window_title, executable_path, process_name)
    except:
        return ("Desktop", None, "Desktop")


# Function to get app icon
def get_app_icon(path, size=(16, 16)):
    if path is None:
        return ImageTk.PhotoImage(Image.new('RGBA', size, (255, 0, 0, 0)))  # Red square as a fallback
    try:
        large, _ = win32gui.ExtractIconEx(path, 0)
        if large:
            hicon = large[0]
            win32gui.DestroyIcon(large[0])
            icon = Image.frombytes('RGBA', size, win32gui.GetIconInfo(hicon)[3])
            win32gui.DestroyIcon(hicon)
            return ImageTk.PhotoImage(icon)
    except Exception as e:
        print(f"Error getting icon for {path}: {e}")
    return ImageTk.PhotoImage(Image.new('RGBA', size, (255, 0, 0, 0)))  # Red square as a fallback


class AppTracker(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("App Usage Tracker - Alpha")
        self.geometry("800x400")

        self.tree = ttk.Treeview(self, columns=("Title", "Duration"), show="headings")
        self.tree.heading("Title", text="Title", command=lambda: self.sort_column("Title"))
        self.tree.heading("Duration", text="Duration", command=lambda: self.sort_column("Duration"))
        self.tree.column("Title", width=400)  # Adjust the column width as needed
        self.tree.column("Duration", width=200)  # Adjust the column width as needed
        self.tree.pack(fill=tk.BOTH, expand=True)

        report_button = tk.Button(self, text="Open Report", command=self.open_report)
        report_button.pack()

        self.app_durations = {}
        self.current_app = None
        self.current_app_start_time = datetime.datetime.now()

        self.current_week = datetime.datetime.now().isocalendar()[1]

        self.load_data()
        self.save_data_periodically()
        self.track_apps()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def generate_html_report(self):
        html_content = """
        <html><head>
        <link rel="stylesheet" type="text/css" href="styles.css">
        <title>App Usage Report</title>
        </head><body>"""
        html_content += "<h1>App Usage Report</h1>"
        html_content += "<p>Report generated on: " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "</p>"
        html_content += "<table border='1'><tr><th>Application</th><th>Duration</th></tr>"

        for app, data in self.app_durations.items():
            duration = format_duration(data["duration"])
            html_content += f"<tr><td>{app}</td><td>{duration}</td></tr>"

        html_content += "</table></body></html>"
        return html_content

    def save_html_report(self):
        html_content = self.generate_html_report()
        with open("app_usage_report.html", "w") as file:
            file.write(html_content)

    def open_report(self):
        self.save_html_report()
        webbrowser.open("app_usage_report.html")

    def load_data(self):
        filename = f"app_durations_week_{self.current_week}.json"
        if os.path.exists(filename):
            with open(filename, 'r') as file:
                loaded_data = json.load(file)
                for app, data in loaded_data.items():
                    last_timestamp = datetime.datetime.fromisoformat(data["last_timestamp"])
                    self.app_durations[app] = {"duration": data["duration"], "last_timestamp": last_timestamp}
                self.update_ui()

    def update_duration(self):
        if self.current_app is not None:
            now = datetime.datetime.now()
            elapsed = (now - self.current_app_start_time).total_seconds()
            if self.current_app not in self.app_durations:
                self.app_durations[self.current_app] = {"duration": 0, "last_timestamp": now}
            self.app_durations[self.current_app]["duration"] += elapsed
            self.app_durations[self.current_app]["last_timestamp"] = now

    def save_data(self):
        self.update_duration()
        filename = f"app_durations_week_{self.current_week}.json"
        with open(filename, 'w') as file:
            data_to_save = {app: {"duration": data["duration"], "last_timestamp": data["last_timestamp"].isoformat()}
                            for app, data in self.app_durations.items()}
            json.dump(data_to_save, file, indent=4)

    def save_data_periodically(self):
        self.save_data()
        self.after(60000, self.save_data_periodically)

    def on_close(self):
        self.save_data()
        self.destroy()

    def update_ui(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        data_to_display = [(app, format_duration(data['duration'])) for app, data in self.app_durations.items()]
        data_to_display.sort(key=lambda x: x[0])  # Sort by title initially
        for app, duration in data_to_display:
            self.tree.insert("", "end", values=(app, duration))

    def contains_file_path(self, title):
        path_pattern = r'([a-zA-Z]:\\[^\/:*?"<>|\r\n]+|\\[^\/:*?"<>|\r\n]+)'
        return re.search(path_pattern, title) is not None

    def track_apps(self):
        window_title, _, _ = get_active_window()
        if self.contains_file_path(window_title):
            return

        now = datetime.datetime.now()
        if self.current_app is None:
            self.current_app = window_title
            self.current_app_start_time = now
        elif window_title != self.current_app:
            self.update_duration()
            self.current_app = window_title
            self.current_app_start_time = now
        else:
            elapsed = (now - self.current_app_start_time).total_seconds()
            if self.current_app not in self.app_durations:
                self.app_durations[self.current_app] = {"duration": 0, "last_timestamp": now}
            self.app_durations[self.current_app]["duration"] += elapsed
            self.app_durations[self.current_app]["last_timestamp"] = now
            self.current_app_start_time = now

        self.update_ui()
        self.after(1000, self.track_apps)

    def sort_column(self, col):
        data_to_display = [(app, format_duration(data['duration'])) for app, data in self.app_durations.items()]

        if col == "Title":
            data_to_display.sort(key=lambda x: x[0])
        elif col == "Duration":
            data_to_display.sort(key=lambda x: self.parse_duration(x[1]))

        for item in self.tree.get_children():
            self.tree.delete(item)

        for app, duration in data_to_display:
            self.tree.insert("", "end", values=(app, duration))

    @staticmethod
    def parse_duration(duration_str):
        parts = duration_str.split()
        total_seconds = 0
        for i in range(0, len(parts), 2):
            value = int(parts[i][:-1])
            unit = parts[i + 1]
            if unit == "d":
                total_seconds += value * 86400
            elif unit == "h":
                total_seconds += value * 3600
            elif unit == "m":
                total_seconds += value * 60
            elif unit == "s":
                total_seconds += value
        return total_seconds


if __name__ == "__main__":
    app = AppTracker()
    app.mainloop()
