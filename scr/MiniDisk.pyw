import psutil
import pystray
from PIL import Image, ImageDraw
import threading
import time
import os
import json
import sys
import tkinter as tk
from tkinter import ttk, messagebox, colorchooser
import multiprocessing

ICON_PATH = os.path.join("icons", "disk.ico")
CONFIG_PATH = "config.json"

EXIT_EVENT = threading.Event()
TRAY_ICONS = []

DEFAULT_CONFIG = {
    "update_interval": 10,
    "autostart": False,
    "colors": {
        "normal": [0, 200, 0, 200],
        "warning": [230, 180, 40, 200],
        "critical": [200, 50, 50, 200]
    }
}

# ---------- CONFIG ----------

def load_config():
    if not os.path.exists(CONFIG_PATH):
        return DEFAULT_CONFIG.copy()
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            return {**DEFAULT_CONFIG, **data}
    except Exception:
        return DEFAULT_CONFIG.copy()


def save_config(config):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)


# ---------- AUTOSTART ----------

def get_startup_path():
    return os.path.join(
        os.getenv("APPDATA"),
        "Microsoft",
        "Windows",
        "Start Menu",
        "Programs",
        "Startup"
    )


def set_autostart(enable):
    startup = get_startup_path()
    shortcut = os.path.join(startup, "MiniDisk.lnk")

    if enable and not os.path.exists(shortcut):
        from win32com.client import Dispatch
        shell = Dispatch("WScript.Shell")
        sc = shell.CreateShortcut(shortcut)
        sc.TargetPath = sys.executable
        sc.Arguments = os.path.abspath(__file__)
        sc.WorkingDirectory = os.getcwd()
        sc.save()

    if not enable and os.path.exists(shortcut):
        os.remove(shortcut)


# ---------- DISKS ----------

def get_fixed_disks():
    return [p.mountpoint for p in psutil.disk_partitions(all=False) if p.fstype]


def get_disk_info(disk):
    u = psutil.disk_usage(disk)
    return u.percent, u.used, u.free, u.total


def load_icon(percent, colors):
    base = Image.open(ICON_PATH).convert("RGBA")
    draw = ImageDraw.Draw(base)

    w, h = base.size
    bar_height = int(h * percent / 100)

    if percent >= 90:
        color = tuple(colors["critical"])
    elif percent >= 70:
        color = tuple(colors["warning"])
    else:
        color = tuple(colors["normal"])

    draw.rectangle((0, h - bar_height, w, h), fill=color)
    return base


# ---------- SETTINGS WINDOW ----------

def settings_process():
    config = load_config()
    colors = config["colors"]

    root = tk.Tk()
    root.title("MiniDisk Settings")
    root.iconbitmap(ICON_PATH)
    root.resizable(False, False)

    ttk.Label(root, text="Update interval (seconds):")\
        .grid(row=0, column=0, padx=10, pady=6, sticky="w")

    interval_var = tk.IntVar(value=config["update_interval"])
    ttk.Entry(root, textvariable=interval_var, width=10)\
        .grid(row=0, column=1, padx=10)

    autostart_var = tk.BooleanVar(value=config["autostart"])
    ttk.Checkbutton(
        root,
        text="Run at Windows startup",
        variable=autostart_var
    ).grid(row=1, column=0, columnspan=2, padx=10, pady=6, sticky="w")

    def choose_color(key):
        rgb, _ = colorchooser.askcolor()
        if rgb:
            colors[key][:3] = list(map(int, rgb))

    ttk.Button(root, text="Normal color",
               command=lambda: choose_color("normal"))\
        .grid(row=2, column=0, padx=10, pady=4, sticky="w")

    ttk.Button(root, text="Warning color",
               command=lambda: choose_color("warning"))\
        .grid(row=3, column=0, padx=10, pady=4, sticky="w")

    ttk.Button(root, text="Critical color",
               command=lambda: choose_color("critical"))\
        .grid(row=4, column=0, padx=10, pady=4, sticky="w")

    def save():
        try:
            interval = int(interval_var.get())
            if interval < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Update interval must be a positive number.")
            return

        new_config = {
            "update_interval": interval,
            "autostart": autostart_var.get(),
            "colors": colors
        }

        save_config(new_config)
        set_autostart(new_config["autostart"])

        messagebox.showinfo(
            "MiniDisk",
            "Settings saved.\nRestart MiniDisk to apply changes."
        )
        root.destroy()

    ttk.Button(root, text="Save", command=save)\
        .grid(row=5, column=0, padx=10, pady=10)
    ttk.Button(root, text="Cancel", command=root.destroy)\
        .grid(row=5, column=1, padx=10, pady=10)

    root.mainloop()


def open_settings(icon=None, item=None):
    multiprocessing.Process(target=settings_process).start()


# ---------- EXIT ALL ----------

def exit_all(icon=None, item=None):
    EXIT_EVENT.set()
    for ic in TRAY_ICONS:
        try:
            ic.stop()
        except Exception:
            pass
    os._exit(0)


# ---------- TRAY ----------

def tray_worker(disk, config):
    percent, used, free, total = get_disk_info(disk)

    icon = pystray.Icon(
        name=f"MiniDisk_{disk}",
        icon=load_icon(percent, config["colors"]),
        title="MiniDisk",
        menu=pystray.Menu(
            pystray.MenuItem("Settings", open_settings),
            pystray.MenuItem("Exit", exit_all)
        )
    )

    TRAY_ICONS.append(icon)

    def update():
        while not EXIT_EVENT.is_set():
            percent, used, free, total = get_disk_info(disk)
            icon.icon = load_icon(percent, config["colors"])
            icon.title = (
                f"MiniDisk — {disk}\n"
                f"Used: {used // (1024**3)} GB\n"
                f"Free: {free // (1024**3)} GB\n"
                f"Total: {total // (1024**3)} GB\n"
                f"Usage: {percent}%"
            )
            time.sleep(config["update_interval"])

    threading.Thread(target=update, daemon=True).start()
    icon.run()


# ---------- MAIN ----------

def main():
    config = load_config()
    set_autostart(config["autostart"])

    for disk in get_fixed_disks():
        threading.Thread(
            target=tray_worker,
            args=(disk, config),
            daemon=True
        ).start()

    while not EXIT_EVENT.is_set():
        time.sleep(1)


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
