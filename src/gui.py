"""BridgeTrace GUI (customtkinter).

Pick the raw leave export, the working calendar and an output folder; the tool saves
the two-sheet report, dated from the export's file name.

Run with:  python src/gui.py
"""
from __future__ import annotations

import queue
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

import config
from main import export_date, report_filename, run_pipeline

RAW_FILETYPES = [("Excel workbook", "*.xlsx *.xls"), ("All files", "*.*")]
CALENDAR_FILETYPES = [("Excel workbook", "*.xlsx"), ("All files", "*.*")]

MUTED = "gray40"
GREEN = "#2FA36B"
RED = "#E5484D"
READY_TEXT = "Choose the three inputs above, then generate the report."


class StepCard(ctk.CTkFrame):
    """A numbered input: title, path entry with a Browse button, and a hint line."""

    def __init__(self, master, number: int, title: str, placeholder: str, hint: str, command) -> None:
        super().__init__(master, corner_radius=12)
        self.grid_columnconfigure(1, weight=1)
        self._default_hint = hint

        ctk.CTkLabel(
            self, text=str(number), width=28, height=28, corner_radius=8,
            fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"], text_color="white",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, padx=(16, 12), pady=(12, 0))
        ctk.CTkLabel(self, text=title, font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=1, columnspan=2, sticky="w", pady=(12, 0)
        )
        self.entry = ctk.CTkEntry(self, placeholder_text=placeholder, height=34)
        self.entry.grid(row=1, column=1, sticky="ew", pady=(4, 0))
        self.button = ctk.CTkButton(self, text="Browse…", width=100, height=34, command=command)
        self.button.grid(row=1, column=2, padx=(10, 16), pady=(4, 0))
        self.hint = ctk.CTkLabel(self, height=20, anchor="w", font=ctk.CTkFont(size=12))
        self.hint.grid(row=2, column=1, columnspan=2, sticky="w", pady=(4, 10))
        self.set_hint()

    def get(self) -> str:
        return self.entry.get().strip()

    def set(self, value: str) -> None:
        self.entry.delete(0, "end")
        if value:
            self.entry.insert(0, value)
            self.entry.xview_moveto(1)

    def set_hint(self, text: str | None = None, color=MUTED) -> None:
        self.hint.configure(text=text or self._default_hint, text_color=color)

    def set_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.entry.configure(state=state)
        self.button.configure(state=state)


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("BridgeTrace")
        self.geometry("760x740")
        self.minsize(640, 680)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        self._events: "queue.Queue[tuple[str, object]]" = queue.Queue()
        self._auto_folder: str | None = None  # output folder auto-filled from the raw file

        self._build_header().grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 10))
        self.raw = StepCard(
            self, 1, "Raw leave export", "Select the leave export (.xlsx)",
            "The report is dated from this file's name.", self._browse_raw,
        )
        self.calendar = StepCard(
            self, 2, "Working calendar", "Select the calendar workbook (.xlsx)",
            "Continuous leave is checked for this calendar's year only.", self._browse_calendar,
        )
        self.output = StepCard(
            self, 3, "Save to folder", "Select the folder for the report",
            "The report name is generated automatically.", self._browse_output,
        )
        for row, card in enumerate((self.raw, self.calendar, self.output), start=1):
            card.grid(row=row, column=0, sticky="ew", padx=24, pady=5)
        self.raw.entry.bind("<KeyRelease>", lambda _event: self._refresh_names())

        self._build_actions().grid(row=4, column=0, sticky="ew", padx=24, pady=(12, 4))
        self._build_log().grid(row=5, column=0, sticky="nsew", padx=24, pady=(4, 20))

        self.after(100, self._drain_events)

    # --- UI construction ---------------------------------------------------
    def _build_header(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text="BridgeTrace", font=ctk.CTkFont(size=24, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkLabel(
            frame, text="Turns a raw leave export into a two-sheet review workbook.", text_color=MUTED
        ).grid(row=1, column=0, sticky="w")
        return frame

    def _build_actions(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        self.generate_btn = ctk.CTkButton(
            frame, text="Generate report", height=42, font=ctk.CTkFont(size=15, weight="bold"),
            command=self._on_generate,
        )
        self.generate_btn.grid(row=0, column=0, sticky="ew")
        self.clear_btn = ctk.CTkButton(
            frame, text="Clear", width=110, height=42, fg_color="transparent", border_width=2,
            text_color="gray10", hover_color="gray85", command=self._on_clear,
        )
        self.clear_btn.grid(row=0, column=1, padx=(10, 0))
        self.progress = ctk.CTkProgressBar(frame, height=6)
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.progress.grid_remove()
        self.status = ctk.CTkLabel(frame, height=24, anchor="w")
        self.status.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self._set_status(READY_TEXT)
        return frame

    def _build_log(self) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(frame, text="Activity", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w"
        )
        self.log = ctk.CTkTextbox(frame, height=130, wrap="word", font=ctk.CTkFont(family="Consolas", size=12))
        self.log.grid(row=1, column=0, sticky="nsew", pady=(4, 0))
        self.log.tag_config("time", foreground="gray55")
        self.log.tag_config("error", foreground=RED)
        self.log.tag_config("success", foreground=GREEN)
        self.log.configure(state="disabled")
        return frame

    # --- Inputs ------------------------------------------------------------
    def _browse_raw(self) -> None:
        path = filedialog.askopenfilename(parent=self, title="Select the raw leave export", filetypes=RAW_FILETYPES)
        if not path:
            return
        path = str(Path(path))
        self.raw.set(path)
        if self.output.get() in ("", self._auto_folder):
            self._auto_folder = str(Path(path).parent)
            self.output.set(self._auto_folder)
        self._refresh_names()

    def _browse_calendar(self) -> None:
        path = filedialog.askopenfilename(parent=self, title="Select the calendar workbook", filetypes=CALENDAR_FILETYPES)
        if path:
            self.calendar.set(str(Path(path)))
            self.calendar.set_hint()

    def _browse_output(self) -> None:
        start = self.output.get() or (str(Path(self.raw.get()).parent) if self.raw.get() else None)
        folder = filedialog.askdirectory(parent=self, title="Choose where to save the report", initialdir=start, mustexist=True)
        if folder:
            self.output.set(str(Path(folder)))
            self._auto_folder = None
            self._refresh_names()

    def _refresh_names(self) -> None:
        """Show the export date and the resulting report name under the inputs."""
        raw = self.raw.get()
        if not raw:
            self.raw.set_hint()
            self.output.set_hint()
            return
        found = export_date(raw)
        self.raw.set_hint(
            f"Export date: {found:%b} {found.day}, {found.year}" if found
            else "No YYYYMMDD date in the file name, so today's date will be used."
        )
        self.output.set_hint(f"Saved as {report_filename(raw)}")

    # --- Run ---------------------------------------------------------------
    def _on_generate(self) -> None:
        self.calendar.set_hint()
        self._refresh_names()
        raw, cal, folder = self.raw.get(), self.calendar.get(), self.output.get()
        for card, ok, message in (
            (self.raw, bool(raw) and Path(raw).is_file(), "Select an existing raw export file."),
            (self.calendar, bool(cal) and Path(cal).is_file(), "Select an existing calendar workbook."),
            (self.output, bool(folder) and Path(folder).is_dir(), "Select an existing folder for the report."),
        ):
            if not ok:
                card.set_hint(message, RED)
                self._set_status(message, RED)
                return

        out = Path(folder) / report_filename(raw)
        if out.exists() and not messagebox.askyesno(
            "Replace report?", f"{out.name} already exists in this folder.\n\nReplace it?", parent=self
        ):
            return

        self._set_running(True)
        self._log(f"Generating {out.name}")
        threading.Thread(target=self._worker, args=(raw, cal, out), daemon=True).start()

    def _worker(self, raw: str, cal: str, out: Path) -> None:
        try:
            summary = run_pipeline(raw, cal, out, log=lambda msg: self._events.put(("log", msg)))
        except Exception as exc:  # surface any failure to the user
            self._events.put(("error", exc))
        else:
            self._events.put(("done", (summary, out)))

    def _drain_events(self) -> None:
        """Apply worker events on the UI thread (Tk widgets are not thread-safe)."""
        while True:
            try:
                kind, payload = self._events.get_nowait()
            except queue.Empty:
                break
            if kind == "log":
                self._log(payload)
            elif kind == "done":
                self._on_success(*payload)
            else:
                self._on_error(payload)
        self.after(100, self._drain_events)

    def _set_running(self, running: bool) -> None:
        for card in (self.raw, self.calendar, self.output):
            card.set_enabled(not running)
        state = "disabled" if running else "normal"
        self.generate_btn.configure(state=state, text="Generating report…" if running else "Generate report")
        self.clear_btn.configure(state=state)
        if running:
            self.progress.grid()
            self.progress.configure(mode="indeterminate")
            self.progress.start()
            self._set_status("Working… large exports can take a little while.")
        else:
            self.progress.stop()
            self.progress.configure(mode="determinate")

    def _on_success(self, summary, out: Path) -> None:
        self._set_running(False)
        self.progress.set(1)
        self._log(f"Done: {summary.target_rows:,} target-leave rows and {summary.sequences:,} sequence(s).", "success")
        self._set_status(f"Saved {out.name}", GREEN)
        messagebox.showinfo(
            "Report created",
            f"{config.SHEET_FILTERED}: {summary.target_rows:,} rows\n"
            f"{config.SHEET_CONTINUOUS}: {summary.sequences:,} sequence(s)\n\n"
            f"Saved to:\n{out}",
            parent=self,
        )

    def _on_error(self, exc: Exception) -> None:
        self._set_running(False)
        self.progress.grid_remove()
        message = str(exc)
        if isinstance(exc, PermissionError):
            message = f"Access denied:\n{exc.filename}\n\nIf the file is open in Excel, close it and try again."
        self._log(f"ERROR: {message}", "error")
        self._set_status("The report was not created. See Activity for details.", RED)
        messagebox.showerror("Report failed", message, parent=self)

    def _on_clear(self) -> None:
        for card in (self.raw, self.calendar, self.output):
            card.set("")
            card.set_hint()
        self._auto_folder = None
        self.progress.grid_remove()
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self._set_status(READY_TEXT)

    # --- Helpers -----------------------------------------------------------
    def _log(self, message: str, tag: str | None = None) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", datetime.now().strftime("%H:%M:%S  "), "time")
        self.log.insert("end", message + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_status(self, text: str, color=MUTED) -> None:
        self.status.configure(text=text, text_color=color)


def main() -> int:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    App().mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
