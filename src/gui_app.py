"""
Tkinter GUI wrapper around the Excel -> Google Sheets automation script.
"""
from __future__ import annotations

import logging
import threading
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Dict

from dotenv import dotenv_values

from excel_to_gsheet import Settings, configure_logging, run_transfer


@dataclass
class EnvConfig:
    excel_file_path: str = ""
    service_account_file: str = ""
    spreadsheet_id: str = ""
    worksheet_name: str = ""
    columns: str = ""


class TextHandler(logging.Handler):
    """Send logging output to a Tkinter Text widget."""

    def __init__(self, text_widget: tk.Text) -> None:
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        self.text_widget.after(0, self._append, msg)

    def _append(self, msg: str) -> None:
        self.text_widget.configure(state="normal")
        self.text_widget.insert(tk.END, msg + "\n")
        self.text_widget.configure(state="disabled")
        self.text_widget.see(tk.END)


class AutomationGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel -> Google Sheets Automation")
        self.root.geometry("720x640")

        self.project_root = Path(__file__).resolve().parents[1]
        self.env_path = self.project_root / ".env"

        self.entries: Dict[str, tk.Entry] = {}
        self.status_var = tk.StringVar(value="Idle")
        self._is_running = False

        self._build_layout()
        self._configure_logging()
        self._load_env_defaults()

    def _configure_logging(self) -> None:
        # Configure the root logger only once.
        if not logging.getLogger().handlers:
            configure_logging(verbose=False)
        self.text_handler = TextHandler(self.log_text)
        formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        self.text_handler.setFormatter(formatter)
        logging.getLogger().addHandler(self.text_handler)
        logging.getLogger().setLevel(logging.INFO)

    def _build_layout(self) -> None:
        container = ttk.Frame(self.root, padding=20)
        container.pack(fill=tk.BOTH, expand=True)

        def add_entry_row(label: str, key: str, browse: bool = False) -> None:
            row = ttk.Frame(container)
            row.pack(fill=tk.X, pady=5)
            ttk.Label(row, text=label, width=22, anchor=tk.W).pack(side=tk.LEFT)

            entry = ttk.Entry(row)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.entries[key] = entry

            if browse:
                ttk.Button(
                    row,
                    text="Browse…",
                    command=lambda k=key: self._browse_file(k),
                ).pack(side=tk.LEFT, padx=(6, 0))

        add_entry_row("Excel file path:", "excel", browse=True)
        add_entry_row("Service account file:", "service", browse=True)
        add_entry_row("Spreadsheet ID:", "spreadsheet")
        add_entry_row("Worksheet name:", "worksheet")
        add_entry_row("Columns (optional):", "columns")

        button_row = ttk.Frame(container)
        button_row.pack(fill=tk.X, pady=(15, 5))

        self.dry_run_button = ttk.Button(
            button_row, text="Dry Run", command=lambda: self._trigger_run(dry_run=True)
        )
        self.dry_run_button.pack(side=tk.LEFT)

        self.run_button = ttk.Button(
            button_row, text="Run Transfer", command=lambda: self._trigger_run(dry_run=False)
        )
        self.run_button.pack(side=tk.LEFT, padx=(10, 0))

        ttk.Button(button_row, text="Save Config", command=self._save_config).pack(
            side=tk.RIGHT
        )

        status_row = ttk.Frame(container)
        status_row.pack(fill=tk.X, pady=(5, 10))
        ttk.Label(status_row, text="Status:").pack(side=tk.LEFT)
        ttk.Label(status_row, textvariable=self.status_var).pack(side=tk.LEFT, padx=(6, 0))

        log_label = ttk.Label(container, text="Log")
        log_label.pack(anchor=tk.W)

        log_frame = ttk.Frame(container)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, state="disabled", wrap=tk.WORD, height=20)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _browse_file(self, key: str) -> None:
        initial_dir = self.project_root
        file_path = filedialog.askopenfilename(initialdir=initial_dir)
        if file_path:
            self.entries[key].delete(0, tk.END)
            self.entries[key].insert(0, file_path)

    def _load_env_defaults(self) -> None:
        if not self.env_path.exists():
            logging.info("No .env file found at %s", self.env_path)
            return

        values = dotenv_values(self.env_path)
        config = EnvConfig(
            excel_file_path=values.get("EXCEL_FILE_PATH", ""),
            service_account_file=values.get("GOOGLE_SERVICE_ACCOUNT_FILE", ""),
            spreadsheet_id=values.get("GOOGLE_SPREADSHEET_ID", ""),
            worksheet_name=values.get("GOOGLE_WORKSHEET_NAME", ""),
            columns=values.get("GOOGLE_COLUMNS", ""),
        )

        self.entries["excel"].insert(0, config.excel_file_path)
        self.entries["service"].insert(0, config.service_account_file)
        self.entries["spreadsheet"].insert(0, config.spreadsheet_id)
        self.entries["worksheet"].insert(0, config.worksheet_name)
        self.entries["columns"].insert(0, config.columns)
        logging.info("Loaded defaults from %s", self.env_path)

    def _save_config(self) -> None:
        env_lines = [
            f"EXCEL_FILE_PATH={self._normalise_env_path(self.entries['excel'].get())}",
            f"GOOGLE_SERVICE_ACCOUNT_FILE={self._normalise_env_path(self.entries['service'].get())}",
            f"GOOGLE_SPREADSHEET_ID={self.entries['spreadsheet'].get().strip()}",
            f"GOOGLE_WORKSHEET_NAME={self.entries['worksheet'].get().strip()}",
            f"GOOGLE_COLUMNS={self.entries['columns'].get().strip()}",
        ]
        try:
            self.env_path.write_text("\n".join(env_lines) + "\n", encoding="utf-8")
            logging.info("Configuration saved to %s", self.env_path)
            messagebox.showinfo("Saved", f"Configuration written to {self.env_path}")
        except OSError as exc:
            logging.error("Failed to save configuration: %s", exc)
            messagebox.showerror("Error", f"Unable to write .env file:\n{exc}")

    def _normalise_env_path(self, value: str) -> str:
        value = value.strip()
        if not value:
            return value
        path = Path(value)
        if not path.is_absolute():
            return value
        try:
            return str(path.relative_to(self.project_root))
        except ValueError:
            return str(path)

    def _trigger_run(self, *, dry_run: bool) -> None:
        if self._is_running:
            return
        self._set_running(True)

        def worker() -> None:
            try:
                settings = Settings.from_values(
                    excel_file_path=self.entries["excel"].get().strip(),
                    service_account_file=self.entries["service"].get().strip(),
                    spreadsheet_id=self.entries["spreadsheet"].get().strip(),
                    worksheet_name=self.entries["worksheet"].get().strip(),
                    columns=[col for col in self.entries["columns"].get().split(",") if col.strip()],
                    base_dir=self.project_root,
                )
            except Exception as exc:  # noqa: BLE001
                logging.error("Invalid configuration: %s", exc)
                self._set_status("Configuration error")
                messagebox.showerror("Configuration error", str(exc))
                self._set_running(False)
                return

            run_label = "Dry run" if dry_run else "Transfer"
            logging.info("%s started.", run_label)
            self._set_status(f"{run_label} in progress…")
            try:
                run_transfer(settings, dry_run=dry_run)
                logging.info("%s completed.", run_label)
                self._set_status(f"{run_label} completed")
            except Exception as exc:  # noqa: BLE001
                logging.error("Automation failed: %s", exc, exc_info=True)
                self._set_status("Failed")
                messagebox.showerror("Automation failed", str(exc))
            finally:
                self._set_running(False)

        threading.Thread(target=worker, daemon=True).start()

    def _set_running(self, value: bool) -> None:
        self._is_running = value
        state = tk.DISABLED if value else tk.NORMAL
        self.dry_run_button.configure(state=state)
        self.run_button.configure(state=state)

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)


def main() -> None:
    root = tk.Tk()
    gui = AutomationGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
