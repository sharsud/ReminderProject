"""
===============================================================================
PROJECT: Learn with Psudo - Daily Task Reminder
MODULE:  mainFile.py
AUTHOR:  Sudhanshu Sharma
DESCRIPTION:
    Automated daily task reminder utility designed for Windows environment startup.
    - Reads, parses, and normalizes task entries from Excel ('plan.xlsx').
    - Includes a UI toggle to fetch Jira tasks directly via REST API (/rest/api/3/search/jql)
      and merge them into Sheet2.
    - Displays today's pending tasks using a modern CustomTkinter interface.
    - Features inline task creation with an interactive date picker widget.
    - Implements safe file-writing logic with timestamped fallbacks if the source
      Excel file is locked.
    - Designed to run via Windows Task Scheduler ('pythonw.exe') at user logon.

DEPENDENCIES:
    - customtkinter (UI framework)
    - tkcalendar (Calendar popup picker)
    - pandas (Excel data processing)
    - openpyxl (Excel engine)
    - requests (Direct Jira REST API fetching)
    - Pillow / PIL (Logo image handling)

links/URLs:
    -https://id.atlassian.com/manage-profile/security/api-tokens
===============================================================================
"""

import os
import sys
import traceback
from datetime import datetime
import pandas as pd
import customtkinter as ctk
from tkcalendar import DateEntry
from PIL import Image
import requests
from requests.auth import HTTPBasicAuth
from requests.exceptions import (
    ConnectionError,
    Timeout,
    RequestException
)
import warnings

# --- CONSTANTS & CONFIGURATION ---
FILE_NAME = "plan.xlsx"
SHEET_NAME_JIRA = "Sheet2"
DATE_COL_CANDIDATES = ["Created", "Due Date", "Due date", "Date", "date"]
TASK_COL_CANDIDATES = ["Summary", "Task", "task", "Description", "description"]
KEY_COL_CANDIDATES = ["Key", "Issue Key", "key"]
STATUS_COL = "Status"
TIME_COL = "Time"
DATE_DISPLAY_FORMAT = "%d-%m-%Y"
LOGO_FILE = "Logo.png"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_PATH = os.path.join(SCRIPT_DIR, FILE_NAME)

# --- JIRA API CONFIGURATION ---
try:
    # 2. Try to import the external file
    import config

    # Safely extract variables if they exist in the file
    DOMAIN = getattr(config, "DOMAIN", "your-domain.atlassian.net")
    EMAIL = getattr(config, "EMAIL", "email")
    API_TOKEN = getattr(config, "API_TOKEN", "api_token from jira")
    print("Variables loaded successfully from external script.")

except ImportError:
    # 3. Show warning if the file does not exist
    warnings.warn("'config.py' not found! Using local variables.")

JIRA_URL = f"https://{DOMAIN}/rest/api/3/search/jql"

HEADERS = {
    "Accept": "application/json"
}

PARAMS = {
    "jql": (
        "assignee = currentUser() "
        "AND status != 'Done' "
        "ORDER BY created DESC"
    ),
    "maxResults": 100,
    "fields": "summary,status,created,duedate"
}


def err(msg):
    print(msg, file=sys.stderr)


# ============================================================
# Get Jira Data
# ============================================================

def get_jira_issues():
    try:
        print("Connecting to Jira...")
        print(f"URL: {JIRA_URL}")

        response = requests.get(
            JIRA_URL,
            params=PARAMS,
            auth=HTTPBasicAuth(EMAIL, API_TOKEN),
            headers=HEADERS,
            timeout=30
        )

        print(f"Jira Status: {response.status_code}")

        if response.status_code == 401:
            raise PermissionError("Jira authentication failed. Check your email and API token.")

        if response.status_code == 403:
            raise PermissionError("Jira access forbidden. Your account may not have permission to access the issues.")

        if response.status_code == 404:
            raise RuntimeError("Jira endpoint was not found. Check the Jira domain and API endpoint.")

        if response.status_code == 410:
            raise RuntimeError("The Jira API endpoint has been removed. Use the current /rest/api/3/search/jql endpoint.")

        response.raise_for_status()

        try:
            data = response.json()
        except ValueError:
            raise ValueError("Jira returned an invalid JSON response.")

        if not isinstance(data, dict):
            raise ValueError("Unexpected Jira response format.")

        issues = data.get("issues")

        if issues is None:
            raise ValueError("Jira response does not contain an 'issues' field.")

        if not isinstance(issues, list):
            raise ValueError("Jira 'issues' field is not a list.")

        if not issues:
            print("No Jira issues found.")
            return pd.DataFrame(columns=["Key", "Summary", "Status", "Created", "Due Date"])

        rows = []
        for issue in issues:
            try:
                fields = issue.get("fields", {})
                status = fields.get("status") or {}

                rows.append({
                    "Key": issue.get("key"),
                    "Summary": fields.get("summary"),
                    "Status": status.get("name"),
                    "Created": fields.get("created"),
                    "Due Date": fields.get("duedate")
                })
            except Exception as issue_error:
                print(f"Warning: Could not process Jira issue: {issue.get('key', 'Unknown')}")
                print(f"Reason: {issue_error}")

        return pd.DataFrame(rows)

    except Timeout:
        raise TimeoutError("Jira request timed out after 30 seconds.")
    except ConnectionError:
        raise ConnectionError("Could not connect to Jira. Check your internet connection and Jira URL.")
    except RequestException as request_error:
        raise RuntimeError(f"Jira request failed: {request_error}")


# ============================================================
# Format Data
# ============================================================

def format_data(df):
    if df.empty:
        return df

    try:
        df["Created"] = pd.to_datetime(df["Created"], errors="coerce").dt.strftime("%Y-%m-%d %H:%M")
        df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce").dt.strftime("%Y-%m-%d")
        return df
    except Exception as error:
        raise RuntimeError(f"Failed to format Jira data: {error}")


# ============================================================
# Write Data to Excel
# ============================================================

def write_to_excel(df):
    if not os.path.exists(FILE_PATH):
        raise FileNotFoundError(f"Excel file '{FILE_PATH}' was not found.")

    try:
        print(f"Writing data to {FILE_PATH}...")
        print(f"Sheet: {SHEET_NAME_JIRA}")

        with pd.ExcelWriter(
            FILE_PATH,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:
            df.to_excel(writer, sheet_name=SHEET_NAME_JIRA, index=False)

    except PermissionError:
        raise PermissionError(
            f"Could not write to '{FILE_PATH}'. "
            "Make sure the Excel file is not open in Excel."
        )
    except ImportError:
        raise ImportError(
            "openpyxl is required to write Excel files. "
            "Install it using: pip install openpyxl"
        )
    except Exception as error:
        raise RuntimeError(f"Failed to write data to Excel: {error}")


def fetch_and_sync_jira():
    """Wrapper function to perform fetch, format, and write operations."""
    try:
        df = get_jira_issues()
        df = format_data(df)
        write_to_excel(df)
        return True, f"Successfully updated {len(df)} issue(s) from Jira."
    except Exception as e:
        return False, str(e)


def load_dataframe(path, sheet_name=0):
    if not os.path.exists(path):
        raise FileNotFoundError(f"{path} not found.")
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def find_column(df, candidates):
    cols = list(df.columns)
    for c in candidates:
        for col in cols:
            if str(col).lower() == c.lower():
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
    if date_col is None or date_col not in df.columns:
        return pd.Series([datetime.now().date()] * len(df))
    return pd.to_datetime(df[date_col], dayfirst=True, errors="coerce").dt.date


def safe_save(df, path, sheet_name=0):
    try:
        save_df = df.drop(columns=["_parsed_date"], errors="ignore")
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            save_df.to_excel(writer, sheet_name=str(sheet_name), index=False)
        return path
    except Exception as e:
        err(f"Failed to write to {path}: {e}")
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = os.path.splitext(path)[0] + f"_saved_{ts}.xlsx"
        try:
            save_df = df.drop(columns=["_parsed_date"], errors="ignore")
            save_df.to_excel(fallback, index=False, engine="openpyxl")
            return fallback
        except Exception as e2:
            err(f"Failed to write fallback file {fallback}: {e2}")
            raise e2


def show_alert(parent, title, message, alert_type="info"):
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


def build_and_show_ui(df_main, main_date_col, main_task_col):
    ctk.set_appearance_mode("Light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Today's Tasks")
    root.geometry("640x700")
    root.resizable(True, True)
    root.configure(fg_color="#F4F6F9")
    root.attributes("-topmost", True)

    df_jira = pd.DataFrame()

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

    today_date = datetime.now().date()
    date_str = datetime.now().strftime(DATE_DISPLAY_FORMAT)

    subheader = ctk.CTkLabel(
        header_content,
        text=f"Tasks for {date_str}",
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="normal"),
        text_color="#64748B"
    )
    subheader.grid(row=1, column=col_idx, sticky="w")

    # --- Jira Toggle Switch ---
    jira_var = ctk.BooleanVar(value=False)

    def on_jira_toggle():
        nonlocal df_jira
        if jira_var.get():
            success, msg = fetch_and_sync_jira()
            if not success:
                show_alert(root, "Jira Sync Error", msg, alert_type="error")
                jira_var.set(False)
                df_jira = pd.DataFrame()
                render_all_today_tasks()
                return

            try:
                df_jira = load_dataframe(FILE_PATH, sheet_name=SHEET_NAME_JIRA)
                df_jira = ensure_status_column(df_jira)
                j_date_col = find_column(df_jira, DATE_COL_CANDIDATES)
                df_jira["_parsed_date"] = parse_dates(df_jira, j_date_col)
            except Exception as e:
                show_alert(root, "Jira Load Error", f"Could not load {SHEET_NAME_JIRA}: {e}", alert_type="error")
                jira_var.set(False)
                df_jira = pd.DataFrame()
                render_all_today_tasks()
                return
        else:
            df_jira = pd.DataFrame()

        render_all_today_tasks()

    jira_cb = ctk.CTkCheckBox(
        header_content,
        text="Include Jira Tasks",
        variable=jira_var,
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        text_color="#2563EB",
        fg_color="#2563EB",
        command=on_jira_toggle
    )
    jira_cb.grid(row=0, column=col_idx + 1, rowspan=2, padx=(20, 0), sticky="e")
    header_content.columnconfigure(col_idx + 1, weight=1)

    # --- Add Task Input Bar ---
    add_frame = ctk.CTkFrame(root, corner_radius=12, fg_color="#FFFFFF", border_width=1, border_color="#E2E8F0")
    add_frame.pack(fill="x", padx=16, pady=(4, 8))

    task_entry = ctk.CTkEntry(
        add_frame,
        placeholder_text="Enter task description...",
        font=ctk.CTkFont(family="Segoe UI", size=12),
        fg_color="#F8FAFC",
        border_color="#CBD5E1",
        text_color="#1E293B",
        height=36
    )
    task_entry.pack(side="left", fill="x", expand=True, padx=(12, 6), pady=10)

    time_entry = ctk.CTkEntry(
        add_frame,
        placeholder_text="Time (e.g. 10:00 AM)",
        font=ctk.CTkFont(family="Segoe UI", size=11),
        fg_color="#F8FAFC",
        border_color="#CBD5E1",
        text_color="#1E293B",
        width=120,
        height=36
    )
    time_entry.pack(side="left", padx=(0, 6), pady=10)

    date_picker = DateEntry(
        add_frame,
        width=10,
        background="#2563EB",
        foreground="white",
        borderwidth=0,
        date_pattern="dd-mm-yyyy",
        font=("Segoe UI", 9)
    )
    date_picker.pack(side="left", padx=(0, 8), pady=10, ipady=4)

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
    task_container.pack(fill="both", expand=True, padx=16, pady=4)

    var_map = {}
    label_map = {}

    def update_item_visuals(key, is_done):
        if key in label_map:
            lbl = label_map[key]
            lbl.configure(text_color="#94A3B8" if is_done else "#1E293B")

    def toggle_and_save(index, var, is_jira=False):
        is_done = var.get()
        unique_key = f"jira_{index}" if is_jira else f"main_{index}"
        update_item_visuals(unique_key, is_done)
        target_df = df_jira if is_jira else df_main
        sheet = SHEET_NAME_JIRA if is_jira else "Sheet1"

        try:
            target_df.at[index, STATUS_COL] = "Done" if is_done else "Pending"
            saved_path = safe_save(target_df, FILE_PATH, sheet_name=sheet)
            if os.path.abspath(saved_path) != os.path.abspath(FILE_PATH):
                show_alert(
                    root,
                    "Save Warning",
                    f"Could not overwrite {FILE_PATH}. Saved to {saved_path} instead.",
                    alert_type="warning"
                )
        except Exception as e:
            show_alert(root, "Save Error", f"Error saving status: {e}", alert_type="error")

    def render_single_task_card(idx, row, is_jira=False):
        status_val = row.get(STATUS_COL, "Pending")
        checked = str(status_val).strip().lower() in ["done", "true", "1", "yes"]

        if is_jira:
            key_col = find_column(df_jira, KEY_COL_CANDIDATES)
            sum_col = find_column(df_jira, TASK_COL_CANDIDATES)
            key_str = str(row.get(key_col, "")) if key_col else ""
            sum_str = str(row.get(sum_col, "")) if sum_col else ""
            task_text = f"[{key_str}] {sum_str}" if key_str else sum_str
            time_text = ""
        else:
            task_text = str(row.get(main_task_col, "(no task)"))
            time_text = str(row.get(TIME_COL, ""))

        var = ctk.BooleanVar(value=checked)
        unique_key = f"jira_{idx}" if is_jira else f"main_{idx}"

        card = ctk.CTkFrame(
            task_container,
            corner_radius=10,
            fg_color="#F0F9FF" if is_jira else "#F8FAFC",
            border_width=1,
            border_color="#BAE6FD" if is_jira else "#F1F5F9"
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
            fg_color="#0284C7" if is_jira else "#2563EB",
            hover_color="#0369A1" if is_jira else "#1D4ED8",
            border_color="#94A3B8",
            command=lambda i=idx, v=var, j=is_jira: toggle_and_save(i, v, is_jira=j)
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
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold" if is_jira else "normal"),
            text_color="#94A3B8" if checked else "#0369A1" if is_jira else "#1E293B",
            anchor="w",
            justify="left",
            wraplength=360
        )
        task_lbl.pack(side="left", fill="x", expand=True, padx=(0, 12), pady=10)

        var_map[unique_key] = (var, idx, is_jira)
        label_map[unique_key] = task_lbl

    def render_all_today_tasks():
        for widget in task_container.winfo_children():
            widget.destroy()
        var_map.clear()
        label_map.clear()

        # 1. Local Excel Tasks
        main_tasks = df_main[df_main["_parsed_date"] == today_date]
        for idx, row in main_tasks.iterrows():
            render_single_task_card(idx, row, is_jira=False)

        # 2. Jira Tasks
        if jira_var.get() and not df_jira.empty:
            for idx, row in df_jira.iterrows():
                render_single_task_card(idx, row, is_jira=True)

    def add_new_task():
        nonlocal df_main

        new_task_str = task_entry.get().strip()
        new_time_str = time_entry.get().strip()
        selected_date_str = date_picker.get_date().strftime("%d-%m-%Y")
        selected_date_obj = date_picker.get_date()

        if not new_task_str:
            show_alert(root, "Input Required", "Please enter a task description before adding.", alert_type="warning")
            return

        new_row = {
            main_date_col: selected_date_str,
            main_task_col: new_task_str,
            STATUS_COL: "Pending",
            "_parsed_date": selected_date_obj
        }

        if TIME_COL in df_main.columns or new_time_str:
            new_row[TIME_COL] = new_time_str if new_time_str else ""

        new_idx = len(df_main)
        df_main.loc[new_idx] = new_row

        try:
            safe_save(df_main, FILE_PATH, sheet_name="Sheet1")
            task_entry.delete(0, "end")
            time_entry.delete(0, "end")
            render_all_today_tasks()
        except Exception as e:
            show_alert(root, "Save Error", f"Could not add task: {e}", alert_type="error")

    btn_add = ctk.CTkButton(
        add_frame,
        text="+ Add Task",
        fg_color="#2563EB",
        hover_color="#1D4ED8",
        text_color="#FFFFFF",
        font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        corner_radius=8,
        height=36,
        width=90,
        command=add_new_task
    )
    btn_add.pack(side="right", padx=(0, 12), pady=10)

    render_all_today_tasks()

    # --- Footer Action Bar ---
    footer_frame = ctk.CTkFrame(root, fg_color="transparent")
    footer_frame.pack(fill="x", padx=16, pady=(8, 16))

    def mark_all_done():
        for key, (v, idx, is_j) in var_map.items():
            if not v.get():
                v.set(True)
                update_item_visuals(key, True)
                target_df = df_jira if is_j else df_main
                target_df.at[idx, STATUS_COL] = "Done"
        try:
            safe_save(df_main, FILE_PATH, sheet_name="Sheet1")
            if jira_var.get() and not df_jira.empty:
                safe_save(df_jira, FILE_PATH, sheet_name=SHEET_NAME_JIRA)
        except Exception as e:
            show_alert(root, "Save Error", f"Error saving: {e}", alert_type="error")

    def mark_all_pending():
        for key, (v, idx, is_j) in var_map.items():
            if v.get():
                v.set(False)
                update_item_visuals(key, False)
                target_df = df_jira if is_j else df_main
                target_df.at[idx, STATUS_COL] = "Pending"
        try:
            safe_save(df_main, FILE_PATH, sheet_name="Sheet1")
            if jira_var.get() and not df_jira.empty:
                safe_save(df_jira, FILE_PATH, sheet_name=SHEET_NAME_JIRA)
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
        df_main = load_dataframe(FILE_PATH, sheet_name=0)

        main_date_col = find_column(df_main, DATE_COL_CANDIDATES)
        main_task_col = find_column(df_main, TASK_COL_CANDIDATES)

        if main_date_col is None or main_task_col is None:
            err(f"Required columns not found on main sheet. Found: {list(df_main.columns)}")
            return

        df_main = ensure_status_column(df_main)
        df_main["_parsed_date"] = parse_dates(df_main, main_date_col)

        build_and_show_ui(df_main, main_date_col, main_task_col)

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