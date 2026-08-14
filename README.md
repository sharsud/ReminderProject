# 🚀 Jira & Excel Sync Task Manager

A modern, automated Windows desktop task manager built with **Python**, **CustomTkinter**, and **Pandas**.

This utility bridges the gap between **Atlassian Jira Cloud** workflows and local Excel task planning (`plan.xlsx`), helping you manage daily tasks seamlessly without context switching.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)
![CustomTkinter](https://img.shields.io/badge/GUI-CustomTkinter-blueviolet.svg)
![Jira API](https://img.shields.io/badge/API-Atlassian%20Jira%20v3-0052CC.svg)

---

## ✨ Features

* **Direct Jira Cloud Sync:** Fetches active, assigned Jira issues via Atlassian REST API v3 using JQL queries:
  `assignee = currentUser() AND status != 'Done'`
* **Unified Checklist View:** Consolidates local Excel tasks and Jira items into a single, clean desktop interface.
* **Modern Glassmorphic UI:** Built using CustomTkinter with responsive cards, visual completion status, and dynamic styling.
* **Inline Task Creation:** Add new local tasks directly from the UI with an interactive date picker using `tkcalendar`.
* **Safe Excel Writer:** Implements robust multi-sheet handling with `openpyxl`, including automatic timestamped fallback saves if `plan.xlsx` is locked or open in Excel.
* **Windows Startup Ready:** Designed to run silently through Windows Task Scheduler using `pythonw.exe`.

---

## 📁 Repository Structure

```text
.
├── mainFile.py       # Core application logic, GUI, and Jira API integration
├── config.py         # Configuration file for Jira API credentials & domain
├── plan.xlsx         # Excel database storing local (Sheet1) & Jira (Sheet2) tasks
├── Logo.png          # App header icon / logo
└── README.md         # Project documentation
```

---

## ⚙️ Configuration Setup

For security, Jira connection details and API credentials are kept separately inside `config.py`.

Create a file named `config.py` in the root directory with the following variables:

```python
# config.py

# Atlassian Jira Configuration
DOMAIN = "your-domain.atlassian.net"
EMAIL = "your-email@example.com"
API_TOKEN = "your_atlassian_api_token"

# JQL Query Configuration
JIRA_URL = f"https://{DOMAIN}/rest/api/3/search/jql"
```

> **⚠️ Security Note:** Do not commit `config.py` to public version control. Add `config.py` to your `.gitignore` to protect your Atlassian API token.

Example `.gitignore` entry:

```gitignore
config.py
```

---

## 📦 Prerequisites & Installation

### 1. Clone the Repository

```bash
git clone https://github.com/your-username/jira-excel-task-reminder.git
cd jira-excel-task-reminder
```

### 2. Install Required Dependencies

Ensure you have **Python 3.8 or higher** installed.

Install the required Python packages:

```bash
pip install pandas openpyxl customtkinter tkcalendar requests Pillow
```

---

## 🚀 Usage

### Running the Application

Launch the desktop application using:

```bash
python mainFile.py
```

---

## 🪟 Setting Up Windows Startup

To have the task manager launch automatically when you log in to Windows:

1. Open **Windows Task Scheduler**.
2. Create a new task triggered **At log on**.
3. Set the action to **Start a program**.
4. Set the program to `pythonw.exe`.
5. Pass `mainFile.py` as the argument.
6. Configure the **Start in** directory to the project folder.

Using `pythonw.exe` allows the application to run without opening a console window.

---

## 🛠️ Built With

* **CustomTkinter** — Modern desktop GUI framework
* **Pandas** — Data processing and Excel data handling
* **OpenPyXL** — Excel file I/O and multi-sheet management
* **Atlassian Jira REST API v3** — Jira Cloud task integration
* **tkcalendar** — Interactive GUI date picker
* **Requests** — HTTP communication with Jira
* **Pillow** — Image and logo handling

---

## 📄 License

Distributed under the **MIT License**. See `LICENSE` for more information.

---

## ❤️ Author

Developed with ❤️ by **Learn with Psudo**
