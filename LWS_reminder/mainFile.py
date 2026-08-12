import os
import sys
import traceback
from datetime import datetime
import pandas as pd
import customtkinter as ctk
from PIL import Image

FILE_NAME = "plan.xlsx"
DATE_COL_CANDIDATES = ["Date", "date"]
TASK_COL_CANDIDATES = ["Task", "task", "Description", "description"]
STATUS_COL = "Status"
DATE_DISPLAY_FORMAT = "%d-%m-%Y"
LOGO_FILE = "Logo.png"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(SCRIPT_DIR, FILE_NAME)


def err(msg):
    print(msg, file=sys.stderr)


def load_dataframe(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"{path} not found.")
    return pd.read_excel(path, engine="openpyxl")


def find_column(df, candidates):
    cols = list(df.columns)
    for c in candidates:
        for col in cols:
            if col.lower() == c.lower():
                return col
    return None


def ensure_status_column(df):
    if STATUS_COL not in df.columns:
        df[STATUS_COL] = "Pending"
    df[STATUS_COL] = df[STATUS_COL].fillna("Pending").apply(
        lambda v: "Done" if str(v).strip().lower() in ["done", "true", "1", "yes"] else "Pending"
    )
    return df


def parse_dates(df, date_col):
    return pd.to_datetime(df[date_col], dayfirst=True, errors="coerce").dt.date


def safe_save(df, path):
    try:
        df.to_excel(path, index=False, engine="openpyxl")
        return path
    except Exception as e:
        err(f"Failed to write to {path}: {e}")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = os.path.splitext(path)[0] + f"_saved_{ts}.xlsx"
        try:
            df.to_excel(fallback, index=False, engine="openpyxl")
            return fallback
        except Exception as e2:
            err(f"Failed to write fallback file {fallback}: {e2}")
            raise e2


def show_alert(parent, title, message, alert_type="info"):
    """Displays a modern native overlay dialog directly over the UI."""
    # Fixed: Replaced 'rgba(...)' with a solid hex tint color supported by Tkinter
    overlay = ctk.CTkFrame(parent, fg_color="#CBD5E1", corner_radius=0)
    overlay.place(relx=0, rely=0, relwidth=1, relheight=1)

    border_color = "#EF4444" if alert_type == "error" else "#F59E0B" if alert_type == "warning" else "#2563EB"

    card = ctk.CTkFrame(
        overlay,
        fg_color="#FFFFFF",
        corner_radius=16,
        border_width=2,
        border_color=border_color,
        width=380
    )
    card.place(relx=0.5, rely=0.5, anchor="center")

    ctk.CTkLabel(
        card,
        text=title,
        font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
        text_color="#1E293B"
    ).pack(padx=20, pady=(16, 6), anchor="w")

    ctk.CTkLabel(
        card,
        text=message,
        font=ctk.CTkFont(family="Segoe UI", size=12),
        text_color="#64748B",
        wraplength=320,
        justify="left"
    ).pack(padx=20, pady=(0, 16), anchor="w")

    ctk.CTkButton(
        card,
        text="OK",
        fg_color="#1E293B",
        hover_color="#334155",
        text_color="#FFFFFF",
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        corner_radius=8,
        height=32,
        width=80,
        command=overlay.destroy
    ).pack(padx=20, pady=(0, 16), anchor="e")


def build_and_show_ui(df_all, tasks_today, date_col, task_col):
    ctk.set_appearance_mode("Light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Today's Tasks")
    root.geometry("540x600")
    root.resizable(True, True)
    root.configure(fg_color="#F4F6F9")
    root.attributes("-topmost", True)

    # --- Header Banner ---
    header_frame = ctk.CTkFrame(root, corner_radius=16, fg_color="#FFFFFF", border_width=1, border_color="#E2E8F0")
    header_frame.pack(pady=(16, 8), padx=16, fill="x")

    header_content = ctk.CTkFrame(header_frame, fg_color="transparent")
    header_content.pack(fill="x", padx=16, pady=12)

    col_idx = 0
    if os.path.exists(LOGO_FILE):
        try:
            logo_img = ctk.CTkImage(
                light_image=Image.open(LOGO_FILE),
                dark_image=Image.open(LOGO_FILE),
                size=(48, 48)
            )
            logo_label = ctk.CTkLabel(header_content, image=logo_img, text="")
            logo_label.grid(row=0, column=0, rowspan=2, padx=(0, 12), sticky="w")
            col_idx = 1
        except Exception as e:
            err(f"Could not load logo: {e}")

    title_label = ctk.CTkLabel(
        header_content,
        text="LEARN WITH PSUDO",
        font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
        text_color="#1E293B"
    )
    title_label.grid(row=0, column=col_idx, sticky="w")

    date_str = datetime.now().strftime(DATE_DISPLAY_FORMAT)
    subheader = ctk.CTkLabel(
        header_content,
        text=f"Tasks for {date_str}",
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="normal"),
        text_color="#64748B"
    )
    subheader.grid(row=1, column=col_idx, sticky="w")

    # --- Task List Section ---
    task_container = ctk.CTkScrollableFrame(
        root,
        corner_radius=16,
        fg_color="#FFFFFF",
        border_width=1,
        border_color="#E2E8F0",
        scrollbar_button_color="#CBD5E1",
        scrollbar_button_hover_color="#94A3B8"
    )
    task_container.pack(fill="both", expand=True, padx=16, pady=8)

    var_map = {}
    cb_map = {}
    label_map = {}

    def update_item_visuals(idx, is_done):
        if idx in label_map:
            lbl = label_map[idx]
            lbl.configure(text_color="#94A3B8" if is_done else "#1E293B")

    def toggle_and_save(index, var):
        is_done = var.get()
        update_item_visuals(index, is_done)
        try:
            df_all.at[index, STATUS_COL] = "Done" if is_done else "Pending"
            saved_path = safe_save(df_all, FILE_PATH)
            if os.path.abspath(saved_path) != os.path.abspath(FILE_PATH):
                show_alert(
                    root,
                    "Save Warning",
                    f"Could not overwrite {FILE_PATH}. Saved to {saved_path} instead.",
                    alert_type="warning"
                )
        except Exception as e:
            show_alert(root, "Save Error", f"Error saving status: {e}", alert_type="error")

    # Render Task Cards
    for idx, row in tasks_today.iterrows():
        task_text = str(row.get(task_col, "(no task)"))
        time_text = str(row.get("Time", ""))
        status_val = row.get(STATUS_COL, "Pending")
        checked = str(status_val).strip().lower() in ["done", "true", "1", "yes"]

        var = ctk.BooleanVar(value=checked)

        card = ctk.CTkFrame(
            task_container,
            corner_radius=10,
            fg_color="#F8FAFC",
            border_width=1,
            border_color="#F1F5F9"
        )
        card.pack(fill="x", pady=4, padx=4)

        cb = ctk.CTkCheckBox(
            card,
            text="",
            variable=var,
            width=24,
            height=24,
            checkbox_width=20,
            checkbox_height=20,
            corner_radius=6,
            border_width=2,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            border_color="#94A3B8",
            command=lambda i=idx, v=var: toggle_and_save(i, v)
        )
        cb.pack(side="left", padx=(12, 6), pady=10)

        if time_text and time_text.lower() != "nan":
            time_badge = ctk.CTkFrame(card, corner_radius=6, fg_color="#DBEAFE")
            time_badge.pack(side="left", padx=(0, 8), pady=10)

            time_lbl = ctk.CTkLabel(
                time_badge,
                text=time_text,
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                text_color="#1E40AF"
            )
            time_lbl.pack(padx=6, pady=2)

        task_lbl = ctk.CTkLabel(
            card,
            text=task_text,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            text_color="#94A3B8" if checked else "#1E293B",
            anchor="w",
            justify="left",
            wraplength=340
        )
        task_lbl.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=10)

        var_map[idx] = var
        cb_map[idx] = cb
        label_map[idx] = task_lbl

    # --- Footer Action Bar ---
    footer_frame = ctk.CTkFrame(root, fg_color="transparent")
    footer_frame.pack(fill="x", padx=16, pady=(8, 16))

    def mark_all_done():
        for i, v in var_map.items():
            if not v.get():
                v.set(True)
                update_item_visuals(i, True)
                df_all.at[i, STATUS_COL] = "Done"
        try:
            safe_save(df_all, FILE_PATH)
        except Exception as e:
            show_alert(root, "Save Error", f"Error saving: {e}", alert_type="error")

    def mark_all_pending():
        for i, v in var_map.items():
            if v.get():
                v.set(False)
                update_item_visuals(i, False)
                df_all.at[i, STATUS_COL] = "Pending"
        try:
            safe_save(df_all, FILE_PATH)
        except Exception as e:
            show_alert(root, "Save Error", f"Error saving: {e}", alert_type="error")

    btn_done = ctk.CTkButton(
        footer_frame,
        text="Mark All Done",
        fg_color="#10B981",
        hover_color="#059669",
        text_color="#FFFFFF",
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        corner_radius=8,
        height=36,
        command=mark_all_done
    )
    btn_done.pack(side="left", padx=(0, 8))

    btn_pending = ctk.CTkButton(
        footer_frame,
        text="Mark All Pending",
        fg_color="#F3F4F6",
        hover_color="#E5E7EB",
        text_color="#374151",
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        corner_radius=8,
        height=36,
        command=mark_all_pending
    )
    btn_pending.pack(side="left")

    btn_close = ctk.CTkButton(
        footer_frame,
        text="Close",
        fg_color="#E2E8F0",
        hover_color="#CBD5E1",
        text_color="#334155",
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        corner_radius=8,
        height=36,
        width=80,
        command=root.destroy
    )
    btn_close.pack(side="right")

    root.lift()
    root.after(1000, lambda: root.attributes("-topmost", False))
    root.mainloop()


def main():
    try:
        df = load_dataframe(FILE_PATH)

        date_col = find_column(df, DATE_COL_CANDIDATES)
        task_col = find_column(df, TASK_COL_CANDIDATES)

        if date_col is None or task_col is None:
            err(f"Required columns not found. Need Date and Task columns. Found: {list(df.columns)}")
            return

        df = ensure_status_column(df)
        df["_parsed_date"] = parse_dates(df, date_col)

        today = datetime.now().date()
        tasks_today = df[df["_parsed_date"] == today]

        if tasks_today.empty:
            return

        build_and_show_ui(df, tasks_today, date_col, task_col)

    except Exception as e:
        traceback.print_exc()
        err(f"Fatal error: {e}")


if __name__ == "__main__":
    try:
        import openpyxl  # noqa: F401
    except Exception:
        print("Missing dependency: openpyxl. Install it with: pip install openpyxl", file=sys.stderr)
        sys.exit(1)

    main()