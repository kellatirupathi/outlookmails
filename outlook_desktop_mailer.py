from __future__ import annotations

import csv
import json
import os
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def get_resource_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


RESOURCE_DIR = get_resource_dir()
USER_DATA_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / "OutlookDesktopMailer"
SCRIPT_PATH = RESOURCE_DIR / "outlook_mailer.ps1"
TEMPLATE_SEED_PATH = RESOURCE_DIR / "templates.json"
TEMPLATES_PATH = USER_DATA_DIR / "templates.json"

DEFAULT_TEMPLATES = {
    "Welcome": {
        "subject": "Welcome {name}",
        "body_mode": "Plain Text",
        "body": (
            "Hi {name},\n\n"
            "Welcome to our service.\n"
            "Your registered email is {email}.\n\n"
            "Regards,\n"
            "Team"
        ),
    },
    "Follow Up": {
        "subject": "Follow up for {name}",
        "body_mode": "Plain Text",
        "body": (
            "Hi {name},\n\n"
            "This is a follow-up message for {company}.\n"
            "Please reply to this email if you need help.\n\n"
            "Regards,\n"
            "Team"
        ),
    },
}

EXAMPLE_RECIPIENTS = """name,email,company
John Doe,john@example.com,Acme
Jane Smith,jane@example.com,Globex
"""


class MailerError(Exception):
    pass


class StrictFormatDict(dict):
    def __missing__(self, key: str) -> str:
        raise KeyError(key)


class OutlookDesktopMailerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Outlook Desktop Mailer")
        self.root.geometry("1240x820")
        self.root.minsize(1080, 720)

        self.templates = self._load_templates()
        self.attachments: list[str] = []

        self.account_var = tk.StringVar()
        self.template_choice_var = tk.StringVar()
        self.template_name_var = tk.StringVar()
        self.subject_var = tk.StringVar()
        self.body_mode_var = tk.StringVar(value="Plain Text")
        self.preview_target_var = tk.StringVar(value="1")

        self._build_ui()
        self._populate_template_choices()
        if self.templates:
            first_name = next(iter(self.templates))
            self.template_choice_var.set(first_name)
            self._load_template_into_form(first_name)

        self.recipients_text.insert("1.0", EXAMPLE_RECIPIENTS)

        self.refresh_accounts()

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        top = ttk.Frame(self.root, padding=12)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Outlook Account").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.account_combo = ttk.Combobox(top, textvariable=self.account_var, state="readonly")
        self.account_combo.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(top, text="Refresh Accounts", command=self.refresh_accounts).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(top, text="Add Attachments", command=self.add_attachments).grid(row=0, column=3, padx=(0, 8))
        ttk.Button(top, text="Remove Selected Attachment", command=self.remove_attachment).grid(row=0, column=4)

        main = ttk.Panedwindow(self.root, orient=tk.HORIZONTAL)
        main.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))

        left = ttk.Frame(main, padding=12)
        right = ttk.Frame(main, padding=12)
        main.add(left, weight=3)
        main.add(right, weight=2)

        left.columnconfigure(0, weight=1)
        left.rowconfigure(4, weight=1)
        left.rowconfigure(5, weight=1)

        template_frame = ttk.LabelFrame(left, text="Template", padding=12)
        template_frame.grid(row=0, column=0, sticky="ew")
        template_frame.columnconfigure(1, weight=1)

        ttk.Label(template_frame, text="Saved Template").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.template_combo = ttk.Combobox(template_frame, textvariable=self.template_choice_var, state="readonly")
        self.template_combo.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        ttk.Button(template_frame, text="Load", command=self.load_selected_template).grid(row=0, column=2)

        ttk.Label(template_frame, text="Template Name").grid(row=1, column=0, sticky="w", pady=(10, 0), padx=(0, 8))
        ttk.Entry(template_frame, textvariable=self.template_name_var).grid(row=1, column=1, sticky="ew", pady=(10, 0), padx=(0, 8))
        ttk.Button(template_frame, text="Save Template", command=self.save_template).grid(row=1, column=2, pady=(10, 0))

        ttk.Label(left, text="Subject Template").grid(row=1, column=0, sticky="w", pady=(12, 4))
        ttk.Entry(left, textvariable=self.subject_var).grid(row=2, column=0, sticky="ew")

        body_header = ttk.Frame(left)
        body_header.grid(row=3, column=0, sticky="ew", pady=(12, 4))
        body_header.columnconfigure(0, weight=1)
        ttk.Label(body_header, text="Body Template").grid(row=0, column=0, sticky="w")
        ttk.Label(body_header, text="Mode").grid(row=0, column=1, padx=(12, 8))
        ttk.Combobox(
            body_header,
            textvariable=self.body_mode_var,
            state="readonly",
            values=("Plain Text", "HTML"),
            width=12,
        ).grid(row=0, column=2, sticky="e")

        self.body_text = tk.Text(left, wrap="word", undo=True, font=("Segoe UI", 10))
        self.body_text.grid(row=4, column=0, sticky="nsew")

        recipients_frame = ttk.LabelFrame(left, text="Recipients CSV", padding=12)
        recipients_frame.grid(row=5, column=0, sticky="nsew", pady=(12, 0))
        recipients_frame.columnconfigure(0, weight=1)
        recipients_frame.rowconfigure(1, weight=1)

        recipients_toolbar = ttk.Frame(recipients_frame)
        recipients_toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        recipients_toolbar.columnconfigure(4, weight=1)
        ttk.Button(recipients_toolbar, text="Load CSV File", command=self.import_csv).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(recipients_toolbar, text="Insert Example", command=self.insert_example_recipients).grid(row=0, column=1, padx=(0, 8))
        ttk.Label(recipients_toolbar, text="Preview Row").grid(row=0, column=2, padx=(0, 8))
        ttk.Entry(recipients_toolbar, textvariable=self.preview_target_var, width=6).grid(row=0, column=3, padx=(0, 8))
        ttk.Label(recipients_toolbar, text="Use CSV headers like name,email,company").grid(row=0, column=4, sticky="w")

        self.recipients_text = tk.Text(recipients_frame, wrap="none", undo=True, font=("Consolas", 10))
        self.recipients_text.grid(row=1, column=0, sticky="nsew")

        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)
        right.rowconfigure(4, weight=1)

        attachment_frame = ttk.LabelFrame(right, text="Attachments", padding=12)
        attachment_frame.grid(row=0, column=0, sticky="nsew")
        attachment_frame.columnconfigure(0, weight=1)
        attachment_frame.rowconfigure(0, weight=1)
        self.attachment_list = tk.Listbox(attachment_frame, height=6)
        self.attachment_list.grid(row=0, column=0, sticky="nsew")

        action_frame = ttk.LabelFrame(right, text="Actions", padding=12)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        action_frame.columnconfigure((0, 1, 2), weight=1)
        ttk.Button(action_frame, text="Preview Email", command=self.preview_email).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(action_frame, text="Create Drafts", command=lambda: self.deliver_messages(draft_only=True)).grid(
            row=0, column=1, sticky="ew", padx=(0, 8)
        )
        ttk.Button(action_frame, text="Send Emails", command=lambda: self.deliver_messages(draft_only=False)).grid(
            row=0, column=2, sticky="ew"
        )

        preview_frame = ttk.LabelFrame(right, text="Preview", padding=12)
        preview_frame.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.preview_text = tk.Text(preview_frame, wrap="word", state="disabled", font=("Consolas", 10))
        self.preview_text.grid(row=0, column=0, sticky="nsew")

        log_frame = ttk.LabelFrame(right, text="Status", padding=12)
        log_frame.grid(row=4, column=0, sticky="nsew", pady=(12, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled", height=10, font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")

    def _load_templates(self) -> dict[str, dict[str, str]]:
        USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
        if not TEMPLATES_PATH.exists():
            self._reset_templates_file()

        try:
            data = json.loads(TEMPLATES_PATH.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("templates.json must contain an object")
            return data
        except Exception:
            self._reset_templates_file()
            return DEFAULT_TEMPLATES.copy()

    def _save_templates_file(self) -> None:
        USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
        TEMPLATES_PATH.write_text(json.dumps(self.templates, indent=2), encoding="utf-8")

    def _reset_templates_file(self) -> None:
        USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
        if TEMPLATE_SEED_PATH.exists():
            TEMPLATES_PATH.write_text(TEMPLATE_SEED_PATH.read_text(encoding="utf-8"), encoding="utf-8")
        else:
            TEMPLATES_PATH.write_text(json.dumps(DEFAULT_TEMPLATES, indent=2), encoding="utf-8")

    def _populate_template_choices(self) -> None:
        names = sorted(self.templates)
        self.template_combo["values"] = names

    def load_selected_template(self) -> None:
        name = self.template_choice_var.get().strip()
        if not name:
            return
        self._load_template_into_form(name)

    def _load_template_into_form(self, name: str) -> None:
        template = self.templates.get(name)
        if not template:
            return
        self.template_name_var.set(name)
        self.subject_var.set(template.get("subject", ""))
        self.body_mode_var.set(template.get("body_mode", "Plain Text"))
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert("1.0", template.get("body", ""))

    def save_template(self) -> None:
        name = self.template_name_var.get().strip()
        if not name:
            messagebox.showerror("Template Name Required", "Enter a template name before saving.")
            return

        self.templates[name] = {
            "subject": self.subject_var.get().strip(),
            "body_mode": self.body_mode_var.get(),
            "body": self.body_text.get("1.0", tk.END).rstrip(),
        }
        self._save_templates_file()
        self._populate_template_choices()
        self.template_choice_var.set(name)
        self._log(f"Saved template '{name}'.")

    def insert_example_recipients(self) -> None:
        self.recipients_text.delete("1.0", tk.END)
        self.recipients_text.insert("1.0", EXAMPLE_RECIPIENTS)

    def import_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Select recipient CSV",
            filetypes=[("CSV Files", "*.csv"), ("Text Files", "*.txt"), ("All Files", "*.*")],
        )
        if not path:
            return
        content = Path(path).read_text(encoding="utf-8-sig")
        self.recipients_text.delete("1.0", tk.END)
        self.recipients_text.insert("1.0", content)
        self._log(f"Loaded recipients from {path}")

    def add_attachments(self) -> None:
        paths = filedialog.askopenfilenames(title="Select attachment files")
        if not paths:
            return
        for path in paths:
            resolved = str(Path(path).resolve())
            if resolved not in self.attachments:
                self.attachments.append(resolved)
                self.attachment_list.insert(tk.END, resolved)
        self._log(f"Added {len(paths)} attachment(s).")

    def remove_attachment(self) -> None:
        selected = list(self.attachment_list.curselection())
        if not selected:
            return
        for index in reversed(selected):
            path = self.attachment_list.get(index)
            self.attachment_list.delete(index)
            if path in self.attachments:
                self.attachments.remove(path)
        self._log("Removed selected attachment(s).")

    def refresh_accounts(self) -> None:
        result = self._run_script("list-accounts")
        accounts = result.get("accounts", [])
        self.account_combo["values"] = accounts

        if accounts:
            if self.account_var.get() not in accounts:
                self.account_var.set(accounts[0])
            self._log(f"Loaded {len(accounts)} Outlook account(s).")
        else:
            self.account_var.set("")
            error = result.get("error", "No Outlook accounts available.")
            self._log(f"Could not load Outlook accounts: {error}")

    def preview_email(self) -> None:
        try:
            messages = self._build_messages()
            preview_row = max(int(self.preview_target_var.get()) - 1, 0)
            if preview_row >= len(messages):
                raise MailerError(f"Preview row {preview_row + 1} is out of range. Loaded rows: {len(messages)}")
            message = messages[preview_row]
            attachment_text = "\n".join(message["attachments"]) if message["attachments"] else "(none)"
            preview = (
                f"To: {message['to']}\n"
                f"Subject: {message['subject']}\n"
                f"Mode: {message['body_mode']}\n"
                f"Attachments:\n{attachment_text}\n\n"
                f"Body:\n{message['body']}"
            )
            self._set_text(self.preview_text, preview)
            self._log(f"Preview generated for row {preview_row + 1}.")
        except Exception as exc:
            messagebox.showerror("Preview Failed", str(exc))

    def deliver_messages(self, draft_only: bool) -> None:
        action_name = "create drafts" if draft_only else "send emails"
        try:
            messages = self._build_messages()
        except Exception as exc:
            messagebox.showerror("Validation Failed", str(exc))
            return

        if not self.account_var.get().strip():
            messagebox.showerror("Outlook Account Required", "Select an Outlook account before continuing.")
            return

        confirm = messagebox.askyesno(
            "Confirm Delivery",
            f"Ready to {action_name} for {len(messages)} recipient(s) using {self.account_var.get()}?",
        )
        if not confirm:
            return

        payload = {
            "sender_account": self.account_var.get().strip(),
            "draft_only": draft_only,
            "messages": messages,
        }

        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as handle:
            payload_path = Path(handle.name)
            json.dump(payload, handle, indent=2)

        try:
            result = self._run_script("deliver", payload_path)
        finally:
            payload_path.unlink(missing_ok=True)

        if not result.get("success"):
            error = result.get("error", "Unknown Outlook delivery failure.")
            messagebox.showerror("Outlook Delivery Failed", error)
            self._log(f"Delivery failed: {error}")
            return

        processed = int(result.get("processed", 0))
        self._log(f"Completed {processed} item(s). Mode: {'drafts' if draft_only else 'send'}.")
        if draft_only:
            messagebox.showinfo("Drafts Created", f"Created {processed} Outlook draft(s).")
        else:
            messagebox.showinfo("Emails Sent", f"Sent {processed} email(s) through Outlook.")

    def _build_messages(self) -> list[dict[str, Any]]:
        recipients = self._parse_recipients()
        subject_template = self.subject_var.get().strip()
        body_template = self.body_text.get("1.0", tk.END).rstrip()
        body_mode = self.body_mode_var.get()

        if not subject_template:
            raise MailerError("Subject template is required.")
        if not body_template:
            raise MailerError("Body template is required.")
        if body_mode not in {"Plain Text", "HTML"}:
            raise MailerError("Body mode must be Plain Text or HTML.")

        messages: list[dict[str, Any]] = []

        # Render the subject and body from recipient CSV columns such as {name}, {email}, {company}.
        for row_number, recipient in enumerate(recipients, start=2):
            email = recipient.get("email", "").strip()
            if not email:
                raise MailerError(f"Recipient CSV row {row_number} is missing the 'email' column value.")
            values = StrictFormatDict(recipient)
            try:
                subject = subject_template.format_map(values)
                body = body_template.format_map(values)
            except KeyError as exc:
                raise MailerError(
                    f"Template placeholder '{exc.args[0]}' does not exist in the recipient CSV headers."
                ) from exc

            messages.append(
                {
                    "to": email,
                    "subject": subject,
                    "body": body,
                    "body_mode": "HTML" if body_mode == "HTML" else "PlainText",
                    "attachments": list(self.attachments),
                }
            )

        return messages

    def _parse_recipients(self) -> list[dict[str, str]]:
        raw_text = self.recipients_text.get("1.0", tk.END).strip()
        if not raw_text:
            raise MailerError("Recipient CSV is empty.")

        reader = csv.DictReader(raw_text.splitlines())
        if not reader.fieldnames:
            raise MailerError("Recipient CSV must include a header row.")

        headers = [header.strip() for header in reader.fieldnames]
        if "email" not in headers:
            raise MailerError("Recipient CSV must include an 'email' header.")

        rows: list[dict[str, str]] = []
        for index, row in enumerate(reader, start=2):
            if not row:
                continue
            normalized: dict[str, str] = {}
            for key, value in row.items():
                if key is None:
                    continue
                normalized[key.strip()] = (value or "").strip()

            if not any(normalized.values()):
                continue
            rows.append(normalized)

        if not rows:
            raise MailerError("Recipient CSV has no data rows.")
        return rows

    def _run_script(self, action: str, payload_path: Path | None = None) -> dict[str, Any]:
        command = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-WindowStyle",
            "Hidden",
            "-File",
            str(SCRIPT_PATH),
            "-Action",
            action,
        ]
        if payload_path is not None:
            command.extend(["-PayloadPath", str(payload_path)])

        completed = subprocess.run(
            command,
            cwd=RESOURCE_DIR,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        stdout = completed.stdout.strip()
        stderr = completed.stderr.strip()

        if not stdout:
            error_message = stderr or "PowerShell script returned no output."
            return {"success": False, "error": error_message}

        try:
            result = json.loads(stdout)
        except json.JSONDecodeError:
            return {"success": False, "error": stdout or stderr}

        if completed.returncode != 0 and "error" not in result:
            result["error"] = stderr or f"PowerShell exited with code {completed.returncode}"
            result["success"] = False
        return result

    def _set_text(self, widget: tk.Text, content: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", content)
        widget.configure(state="disabled")

    def _log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")


def main() -> None:
    if not SCRIPT_PATH.exists():
        messagebox.showerror("Missing Script", f"Could not find PowerShell script: {SCRIPT_PATH}")
        return

    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    OutlookDesktopMailerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
