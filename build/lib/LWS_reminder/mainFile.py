
import os
import sys
import traceback
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk


FILE_NAME = "plan.xlsx"  # change if your file has different name or path
DATE_COL_CANDIDATES = ["Date", "date"]
TASK_COL_CANDIDATES = ["Task", "task", "Description", "description"]
STATUS_COL = "Status"     # column we'll write ("Done" / "Pending")
DATE_DISPLAY_FORMAT = "%d-%m-%Y"  # Excel date shown like "12-08-2025"
LOGO_FILE = "Logo.png"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(SCRIPT_DIR, FILE_NAME)


def err(msg):
    print(msg, file=sys.stderr)


def load_dataframe(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"{path} not found.")
    # read with openpyxl engine (need openpyxl installed)
    df = pd.read_excel(path, engine="openpyxl")
    return df


def find_column(df, candidates):
    # Return first matching column name in df.columns or None
    cols = list(df.columns)
    for c in candidates:
        for col in cols:
            if col.lower() == c.lower():
                return col
    return None


def ensure_status_column(df):
    if STATUS_COL not in df.columns:
        df[STATUS_COL] = "Pending"
    # normalize values
    df[STATUS_COL] = df[STATUS_COL].fillna("Pending").apply(lambda v: "Done" if str(v).strip().lower() in ["done", "true", "1", "yes"] else "Pending")
    return df


def parse_dates(df, date_col):
    # Try to parse the date column into a date object; keep original column untouched
    parsed = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce").dt.date
    return parsed


def safe_save(df, path):
    """
    Try to save to path. If file is locked by Excel (PermissionError / OSError),
    save to a timestamped fallback file and notify user.
    """
    try:
        df.to_excel(path, index=False, engine="openpyxl")
        return path
    except Exception as e:
        err(f"Failed to write to {path}: {e}")
        # fallback
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = os.path.splitext(path)[0] + f"_saved_{ts}.xlsx"
        try:
            df.to_excel(fallback, index=False, engine="openpyxl")
            return fallback
        except Exception as e2:
            err(f"Failed to write fallback file {fallback}: {e2}")
            raise e2


def build_and_show_ui(df_all, tasks_today, date_col, task_col):
    # Root window
    root = tk.Tk()
    root.title("Today's Tasks")
    root.configure(bg="#f4f6f8")
    root.attributes("-topmost", True)
    root.geometry("520x520")
    root.resizable(True, True)

    # --- Header Frame ---
    header_frame = tk.Frame(root, bg="white", bd=0, highlightthickness=0)
    header_frame.pack(pady=15, padx=15, fill="x")
    header_frame.configure(highlightbackground="#ddd", highlightcolor="#ddd")

    # Rounded effect simulation using padding
    header_frame.grid_columnconfigure(1, weight=1)

    if os.path.exists(LOGO_FILE):
        img = Image.open(LOGO_FILE)
        img = img.resize((60, 60), Image.LANCZOS)
        logo_img = ImageTk.PhotoImage(img)
        logo_label = tk.Label(header_frame, image=logo_img, bg="white")
        logo_label.image = logo_img
        logo_label.grid(row=0, column=0, padx=(5, 15), pady=10, sticky="w")

    title_label = tk.Label(
        header_frame,
        text="LEARN WITH PSUDO",
        font=("Segoe UI", 22, "bold"),
        fg="#2c3e50",
        bg="white"
    )
    title_label.grid(row=0, column=1, sticky="w")

    # --- Subheader ---
    header = tk.Label(
        root,
        text=f"Tasks for {datetime.now().strftime(DATE_DISPLAY_FORMAT)}",
        font=("Segoe UI", 14, "bold"),
        fg="#34495e",
        bg="#f4f6f8"
    )
    header.pack(padx=15, pady=(5, 10), anchor="w")

    # --- Scrollable Frame for Tasks ---
    container = tk.Frame(root, bg="#f4f6f8")
    container.pack(fill="both", expand=True, padx=15, pady=10)

    canvas = tk.Canvas(container, bg="white", highlightthickness=0)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scroll_frame = tk.Frame(canvas, bg="white")

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # dictionary to hold tk.BooleanVar for each row index
    var_map = {}

    def toggle_and_save(index, var):
        # Update df_all and save immediately
        try:
            df_all.at[index, STATUS_COL] = "Done" if var.get() else "Pending"
            saved_path = safe_save(df_all, FILE_PATH)
            # If saved to fallback, notify user
            if os.path.abspath(saved_path) != os.path.abspath(FILE_PATH):
                messagebox.showwarning("Save warning", f"Could not overwrite {FILE_PATH}. Saved to {saved_path} instead.\nClose Excel if it's open to allow saving to the original file.")
        except Exception as e:
            messagebox.showerror("Save error", f"Error saving status: {e}")

    # Create a row per task
    for idx, row in tasks_today.iterrows():
        task_text = str(row.get(task_col, "(no task)"))
        time_text = str(row.get("Time", ""))
        status_val = row.get(STATUS_COL, "Pending")
        checked = True if str(status_val).strip().lower() in ["done", "true", "1", "yes"] else False

        var = tk.BooleanVar(value=checked)
        cb_text = f"{time_text + ' - ' if time_text and time_text != 'nan' else ''}{task_text}"
        cb = ttk.Checkbutton(scroll_frame, text=cb_text, variable=var,
                             command=lambda i=idx, v=var: toggle_and_save(i, v))
        cb.pack(anchor="w", pady=4, padx=4)
        var_map[idx] = var

    # Buttons
    btn_frame = ttk.Frame(root)
    btn_frame.pack(fill="x", padx=10, pady=10)

    def mark_all_done():
        for i, v in var_map.items():
            if not v.get():
                v.set(True)
                df_all.at[i, STATUS_COL] = "Done"
        try:
            saved = safe_save(df_all, FILE_PATH)
            if os.path.abspath(saved) != os.path.abspath(FILE_PATH):
                messagebox.showwarning("Save warning", f"Saved to {saved} (couldn't overwrite original file).")
        except Exception as e:
            messagebox.showerror("Save error", f"Error saving: {e}")

    def mark_all_pending():
        for i, v in var_map.items():
            if v.get():
                v.set(False)
                df_all.at[i, STATUS_COL] = "Pending"
        try:
            saved = safe_save(df_all, FILE_PATH)
            if os.path.abspath(saved) != os.path.abspath(FILE_PATH):
                messagebox.showwarning("Save warning", f"Saved to {saved} (couldn't overwrite original file).")
        except Exception as e:
            messagebox.showerror("Save error", f"Error saving: {e}")

    ttk.Button(btn_frame, text="Mark all Done", command=mark_all_done).pack(side="left", padx=(0, 6))
    ttk.Button(btn_frame, text="Mark all Pending", command=mark_all_pending).pack(side="left", padx=(0, 6))
    ttk.Button(btn_frame, text="Close", command=root.destroy).pack(side="right")

    # make sure window stays on top at start
    root.lift()
    root.after(1000, lambda: root.attributes("-topmost", False))  # allow user to move it after first focus
    root.mainloop()


def main():
    try:
        # load
        df = load_dataframe(FILE_PATH)

        # find columns
        date_col = find_column(df, DATE_COL_CANDIDATES)
        task_col = find_column(df, TASK_COL_CANDIDATES)

        if date_col is None or task_col is None:
            messagebox.showerror("Column error", f"Required columns not found. Need a Date column and a Task column.\nFound columns: {list(df.columns)}")
            return

        # ensure status column
        df = ensure_status_column(df)

        # parse date into date objects
        parsed_dates = parse_dates(df, date_col)
        df["_parsed_date"] = parsed_dates  # helper column

        today = datetime.now().date()
        tasks_today = df[df["_parsed_date"] == today]

        if tasks_today.empty:
            # nothing to show; exit quietly
            return

        # show UI and allow updating statuses
        build_and_show_ui(df, tasks_today, date_col, task_col)

    except Exception as e:
        # show a user-friendly error dialog and also print traceback
        traceback.print_exc()
        try:
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
        except Exception:
            err(f"Fatal error: {e}")


if __name__ == "__main__":
    # quick check for required libs
    try:
        import openpyxl  # noqa: F401
    except Exception:
        message = "Missing dependency: openpyxl. Install it with:\n\npip install openpyxl"
        print(message, file=sys.stderr)
        try:
            messagebox.showerror("Missing dependency", message)
        except Exception:
            pass
        sys.exit(1)

    main()
