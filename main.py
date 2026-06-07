import os
import sys
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

import gspread
from google.oauth2.service_account import Credentials

import pystray
from PIL import Image, ImageDraw

# =========================================================
# APP DIR
# =========================================================

if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(APP_DIR, "config.json")

# =========================================================
# CONFIG
# =========================================================

config = {
    "sheet_url": "",
    "credentials_path": "",
    "folders": []
}

observers = []
running = False
tray_icon = None

# =========================================================
# SAVE CONFIG
# =========================================================

def save_config():

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:

        json.dump(
            config,
            f,
            indent=4,
            ensure_ascii=False
        )

# =========================================================
# LOAD CONFIG
# =========================================================

def load_config():

    global config

    if os.path.exists(CONFIG_FILE):

        with open(CONFIG_FILE, "r", encoding="utf-8") as f:

            config = json.load(f)

        migrated_folders = []

        for folder in config.get("folders", []):

            if isinstance(folder, str):

                migrated_folders.append({
                    "path": folder,
                    "worksheet": "Sheet1",
                    "column": "A",
                    "mode": "fullpath"
                })

            elif isinstance(folder, dict):

                migrated_folders.append({
                    "path": folder.get("path", ""),
                    "worksheet": folder.get("worksheet", "Sheet1"),
                    "column": folder.get("column", "A"),
                    "mode": folder.get("mode", "fullpath")
                })

        config["folders"] = migrated_folders

        save_config()

# =========================================================
# LOGGER
# =========================================================

def log(text):

    log_box.insert(tk.END, text + "\n")

    log_box.see(tk.END)

# =========================================================
# FORMAT PATH
# =========================================================

def format_path(path, mode):

    if mode == "filename":
        return os.path.basename(path)

    return path

# =========================================================
# GOOGLE SHEETS
# =========================================================

def connect_sheet(sheet_url, worksheet_name):

    credentials_path = config.get("credentials_path", "")

    if not credentials_path:

        raise Exception("Не выбран credentials.json")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets"
    ]

    creds = Credentials.from_service_account_file(
        credentials_path,
        scopes=scopes
    )

    client = gspread.authorize(creds)

    spreadsheet = client.open_by_url(sheet_url)

    try:

        sheet = spreadsheet.worksheet(worksheet_name)

    except:

        sheet = spreadsheet.add_worksheet(
            title=worksheet_name,
            rows=1000,
            cols=10
        )

    return sheet

# =========================================================
# APPEND PATH
# =========================================================

def append_path(sheet, value, column):

    column_number = ord(column.upper()) - 64

    values = sheet.col_values(column_number)

    next_row = len(values) + 1

    cell = f"{column}{next_row}"

    sheet.update(cell, [[value]])

# =========================================================
# DELETE ROW
# =========================================================

def delete_row_by_value(sheet, value, column):

    try:

        column_number = ord(column.upper()) - 64

        values = sheet.col_values(column_number)

        for i, cell_value in enumerate(values, start=1):

            if cell_value == value:

                sheet.delete_rows(i)

                log(f"[DEL] {value}")

                return

    except Exception as e:

        log(f"[ERROR] {e}")

# =========================================================
# CHUNKS
# =========================================================

def chunks(lst, n):

    for i in range(0, len(lst), n):

        yield lst[i:i + n]

# =========================================================
# INITIAL SYNC
# =========================================================

def initial_sync(folder_config, sheet, known_paths):

    folder = folder_config["path"]

    column = folder_config["column"]

    mode = folder_config.get("mode", "fullpath")

    log(f"=== SYNC: {folder} ===")

    rows_to_add = []

    for root_dir, dirs, files in os.walk(folder):

        for file in files:

            full_path = os.path.join(root_dir, file)

            formatted = format_path(
                full_path,
                mode
            )

            if formatted not in known_paths:

                rows_to_add.append([formatted])

                known_paths.add(formatted)

                log(f"[SYNC +] {formatted}")

    if rows_to_add:

        try:

            column_number = ord(column.upper()) - 64

            existing_values = sheet.col_values(column_number)

            start_row = len(existing_values) + 1

            for chunk in chunks(rows_to_add, 500):

                end_row = start_row + len(chunk) - 1

                range_name = f"{column}{start_row}:{column}{end_row}"

                sheet.update(range_name, chunk)

                start_row = end_row + 1

            log(f"✔ Добавлено: {len(rows_to_add)}")

        except Exception as e:

            log(f"[ERROR] {e}")

    else:

        log("✔ Новых файлов нет")

# =========================================================
# WATCHDOG
# =========================================================

class FileHandler(FileSystemEventHandler):

    def __init__(self, sheet, known_paths, folder_config):

        self.sheet = sheet
        self.known_paths = known_paths
        self.folder_config = folder_config

    def on_created(self, event):

        if event.is_directory:
            return

        path = event.src_path

        column = self.folder_config["column"]

        mode = self.folder_config.get(
            "mode",
            "fullpath"
        )

        formatted = format_path(
            path,
            mode
        )

        if formatted not in self.known_paths:

            self.known_paths.add(formatted)

            append_path(
                self.sheet,
                formatted,
                column
            )

            log(f"[+] {formatted}")

    def on_deleted(self, event):

        if event.is_directory:
            return

        path = event.src_path

        column = self.folder_config["column"]

        mode = self.folder_config.get(
            "mode",
            "fullpath"
        )

        formatted = format_path(
            path,
            mode
        )

        if formatted in self.known_paths:

            self.known_paths.remove(formatted)

            delete_row_by_value(
                self.sheet,
                formatted,
                column
            )

            log(f"[-] {formatted}")

# =========================================================
# START WATCH
# =========================================================

def start_watch():

    global observers
    global running

    if not config["sheet_url"]:

        messagebox.showerror(
            "Ошибка",
            "Укажи Google Sheet URL"
        )

        return

    if not config["credentials_path"]:

        messagebox.showerror(
            "Ошибка",
            "Выбери credentials.json"
        )

        return

    if not config["folders"]:

        messagebox.showerror(
            "Ошибка",
            "Добавь папки"
        )

        return

    observers = []

    for folder_config in config["folders"]:

        try:

            folder = folder_config["path"]

            worksheet = folder_config["worksheet"]

            column = folder_config["column"]

            sheet = connect_sheet(
                config["sheet_url"],
                worksheet
            )

            column_number = ord(column.upper()) - 64

            existing_values = sheet.col_values(
                column_number
            )

            known_paths = set(existing_values)

            initial_sync(
                folder_config,
                sheet,
                known_paths
            )

            handler = FileHandler(
                sheet,
                known_paths,
                folder_config
            )

            observer = Observer()

            observer.schedule(
                handler,
                folder,
                recursive=True
            )

            observer.start()

            observers.append(observer)

            log(f"👁 {folder}")

        except Exception as e:

            log(f"[ERROR] {folder}: {e}")

    running = True

    start_btn.config(text="Стоп")

    log("=== WATCHER STARTED ===")

# =========================================================
# STOP WATCH
# =========================================================

def stop_watch():

    global running

    for obs in observers:

        obs.stop()

    for obs in observers:

        obs.join()

    running = False

    start_btn.config(text="Старт")

    log("=== WATCHER STOPPED ===")

# =========================================================
# TOGGLE WATCH
# =========================================================

def toggle_watch():

    if running:
        stop_watch()
    else:
        start_watch()

# =========================================================
# REFRESH TREE
# =========================================================

def refresh_tree():

    for item in tree.get_children():

        tree.delete(item)

    for folder_data in config.get("folders", []):

        if not isinstance(folder_data, dict):
            continue

        tree.insert(
            "",
            tk.END,
            values=(
                folder_data.get("path", ""),
                folder_data.get("worksheet", "Sheet1"),
                folder_data.get("column", "A"),
                folder_data.get("mode", "fullpath")
            )
        )

# =========================================================
# ADD FOLDER
# =========================================================

def add_folder():

    folder = filedialog.askdirectory()

    if not folder:
        return

    popup = tk.Toplevel(root)

    popup.title("Добавить папку")

    popup.geometry("300x300")

    tk.Label(
        popup,
        text="Worksheet"
    ).pack(pady=5)

    worksheet_entry = tk.Entry(popup)

    worksheet_entry.insert(0, "Sheet1")

    worksheet_entry.pack()

    tk.Label(
        popup,
        text="Column"
    ).pack(pady=5)

    column_entry = tk.Entry(popup)

    column_entry.insert(0, "A")

    column_entry.pack()

    tk.Label(
        popup,
        text="Mode"
    ).pack(pady=5)

    mode_var = tk.StringVar(
        value="fullpath"
    )

    mode_menu = ttk.Combobox(
        popup,
        textvariable=mode_var,
        values=[
            "fullpath",
            "filename"
        ],
        state="readonly"
    )

    mode_menu.pack()

    def save_folder():

        folder_data = {
            "path": folder,
            "worksheet": worksheet_entry.get(),
            "column": column_entry.get().upper(),
            "mode": mode_var.get()
        }

        config["folders"].append(folder_data)

        save_config()

        refresh_tree()

        popup.destroy()

    tk.Button(
        popup,
        text="Сохранить",
        command=save_folder
    ).pack(pady=20)

# =========================================================
# REMOVE FOLDER
# =========================================================

def remove_folder():

    selected = tree.selection()

    if not selected:
        return

    item = selected[0]

    values = tree.item(item, "values")

    folder_path = values[0]

    config["folders"] = [

        f for f in config["folders"]

        if f["path"] != folder_path
    ]

    save_config()

    refresh_tree()

# =========================================================
# EDIT FOLDER
# =========================================================

def edit_folder():

    selected = tree.selection()

    if not selected:
        return

    item = selected[0]

    values = tree.item(item, "values")

    folder_path = values[0]

    folder_data = next(
        (
            f for f in config["folders"]

            if f["path"] == folder_path
        ),
        None
    )

    if not folder_data:
        return

    popup = tk.Toplevel(root)

    popup.title("Редактирование")

    popup.geometry("300x300")

    tk.Label(
        popup,
        text="Worksheet"
    ).pack(pady=5)

    worksheet_entry = tk.Entry(popup)

    worksheet_entry.insert(
        0,
        folder_data.get(
            "worksheet",
            "Sheet1"
        )
    )

    worksheet_entry.pack()

    tk.Label(
        popup,
        text="Column"
    ).pack(pady=5)

    column_entry = tk.Entry(popup)

    column_entry.insert(
        0,
        folder_data.get(
            "column",
            "A"
        )
    )

    column_entry.pack()

    tk.Label(
        popup,
        text="Mode"
    ).pack(pady=5)

    mode_var = tk.StringVar(
        value=folder_data.get(
            "mode",
            "fullpath"
        )
    )

    mode_menu = ttk.Combobox(
        popup,
        textvariable=mode_var,
        values=[
            "fullpath",
            "filename"
        ],
        state="readonly"
    )

    mode_menu.pack()

    def save_edit():

        folder_data["worksheet"] = worksheet_entry.get()

        folder_data["column"] = column_entry.get().upper()

        folder_data["mode"] = mode_var.get()

        save_config()

        refresh_tree()

        popup.destroy()

    tk.Button(
        popup,
        text="Сохранить",
        command=save_edit
    ).pack(pady=20)

# =========================================================
# SAVE SHEET URL
# =========================================================

def save_sheet_url():

    config["sheet_url"] = sheet_entry.get()

    save_config()

    log("✔ URL сохранён")

# =========================================================
# SELECT CREDENTIALS
# =========================================================

def select_credentials():

    file_path = filedialog.askopenfilename(
        title="Выбери credentials.json",
        filetypes=[
            ("JSON", "*.json")
        ]
    )

    if not file_path:
        return

    config["credentials_path"] = file_path

    credentials_entry.delete(0, tk.END)

    credentials_entry.insert(0, file_path)

    save_config()

    log("✔ credentials.json выбран")

# =========================================================
# TRAY
# =========================================================

def create_image():

    image = Image.new(
        'RGB',
        (64, 64),
        color='black'
    )

    draw = ImageDraw.Draw(image)

    draw.rectangle(
        (16, 16, 48, 48),
        fill='white'
    )

    return image

def tray_show(icon, item):

    root.after(
        0,
        root.deiconify
    )

def tray_exit(icon, item):

    stop_watch()

    icon.stop()

    root.destroy()

def run_tray():

    global tray_icon

    menu = pystray.Menu(

        pystray.MenuItem(
            "Показать",
            tray_show
        ),

        pystray.MenuItem(
            "Выход",
            tray_exit
        )
    )

    tray_icon = pystray.Icon(
        "Watcher",
        create_image(),
        "File Watcher PRO",
        menu
    )

    tray_icon.run()

def minimize_to_tray():

    root.withdraw()

    threading.Thread(
        target=run_tray,
        daemon=True
    ).start()

# =========================================================
# UI
# =========================================================

load_config()

root = tk.Tk()

root.title("File Watcher PRO")

root.geometry("1000x800")

# =========================================================
# SHEET URL
# =========================================================

tk.Label(
    root,
    text="Google Sheet URL"
).pack(pady=5)

sheet_entry = tk.Entry(
    root,
    width=120
)

sheet_entry.insert(
    0,
    config.get(
        "sheet_url",
        ""
    )
)

sheet_entry.pack()

tk.Button(
    root,
    text="Сохранить URL",
    command=save_sheet_url
).pack(pady=5)

# =========================================================
# CREDENTIALS
# =========================================================

tk.Label(
    root,
    text="credentials.json"
).pack(pady=5)

credentials_entry = tk.Entry(
    root,
    width=120
)

credentials_entry.insert(
    0,
    config.get(
        "credentials_path",
        ""
    )
)

credentials_entry.pack()

tk.Button(
    root,
    text="Выбрать credentials.json",
    command=select_credentials
).pack(pady=5)

# =========================================================
# TREEVIEW
# =========================================================

tree = ttk.Treeview(
    root,
    columns=(
        "Folder",
        "Worksheet",
        "Column",
        "Mode"
    ),
    show="headings",
    height=12
)

tree.heading(
    "Folder",
    text="Folder"
)

tree.heading(
    "Worksheet",
    text="Worksheet"
)

tree.heading(
    "Column",
    text="Column"
)

tree.heading(
    "Mode",
    text="Mode"
)

tree.column(
    "Folder",
    width=500
)

tree.column(
    "Worksheet",
    width=180
)

tree.column(
    "Column",
    width=80
)

tree.column(
    "Mode",
    width=120
)

tree.pack(
    pady=10
)

refresh_tree()

# =========================================================
# BUTTONS
# =========================================================

btn_frame = tk.Frame(root)

btn_frame.pack(pady=10)

tk.Button(
    btn_frame,
    text="Добавить",
    command=add_folder
).pack(side=tk.LEFT, padx=5)

tk.Button(
    btn_frame,
    text="Редактировать",
    command=edit_folder
).pack(side=tk.LEFT, padx=5)

tk.Button(
    btn_frame,
    text="Удалить",
    command=remove_folder
).pack(side=tk.LEFT, padx=5)

# =========================================================
# WATCH BUTTON
# =========================================================

start_btn = tk.Button(
    root,
    text="Старт",
    width=20,
    height=2,
    command=toggle_watch
)

start_btn.pack(pady=10)

# =========================================================
# TRAY BUTTON
# =========================================================

tk.Button(
    root,
    text="Свернуть в трей",
    command=minimize_to_tray
).pack()

# =========================================================
# INSTRUCTION
# =========================================================

instruction = tk.Text(
    root,
    height=14
)

instruction.insert(
    tk.END,
    """ИНСТРУКЦИЯ:

1. Создай Google таблицу

2. В Google Cloud:
- включи Google Sheets API
- включи Google Drive API
- создай Service Account
- скачай credentials.json

3. В Google таблице:
Поделиться -> добавь email сервисного аккаунта как Редактор

4. В программе:
- выбери credentials.json
- вставь ссылку на Google таблицу
- добавь папки

5. Настройки папки:
- worksheet
- column
- mode

Modes:
- fullpath
- filename

6. Нажми Старт

Все настройки сохраняются автоматически.
"""
)

instruction.pack(
    fill="x",
    padx=10,
    pady=10
)

# =========================================================
# LOG
# =========================================================

log_box = tk.Text(
    root,
    height=15
)

log_box.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=10
)

# =========================================================
# CLOSE
# =========================================================

def on_close():

    minimize_to_tray()

root.protocol(
    "WM_DELETE_WINDOW",
    on_close
)

# =========================================================
# START
# =========================================================

root.mainloop()