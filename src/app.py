"""
Ship Report Builder — cross-platform desktop GUI (Tkinter).

Inputs:
  1. Sample report      (template whose format + formulas are cloned)
  2. PO Number          (optional text)
  3. Build Code         (optional text)
  4. Raw-data folder    (directory of unique-ID-coded test reports)
  5. OK-TO-SHIP report  (Excel workbook)
Then generates a filled report and saves it.
"""

from __future__ import annotations

import queue
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from engine import ReportRequest, generate_report
from lookup import lookup_device, format_lookup

APP_TITLE = "Ship Report Builder"
PAD = 8


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.minsize(680, 460)
        # Tk 8.5 on macOS can paint a black/blank window (esp. dark mode). Use an
        # explicit light background and the aqua theme, and nudge a redraw once
        # the window is up so content is drawn.
        self.configure(bg="#ececec")
        try:
            ttk.Style().theme_use("aqua")
        except tk.TclError:
            pass
        self.after(50, self._force_redraw)

        self.sample_var = tk.StringVar()
        self.po_var = tk.StringVar()
        self.build_var = tk.StringVar()
        self.raw_dir_var = tk.StringVar()
        self.ok_ship_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.lookup_var = tk.StringVar()

        self._msg_q: queue.Queue[tuple[str, str]] = queue.Queue()
        self._build_ui()
        self.after(100, self._drain_queue)

    # ------------------------------------------------------------------ #
    # UI
    # ------------------------------------------------------------------ #
    def _force_redraw(self):
        """Nudge the window size by 1px to force Tk to repaint (macOS Tk 8.5)."""
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"{w + 1}x{h + 1}")
        self.after(30, lambda: self.geometry(f"{w}x{h}"))
        self.lift()
        self.focus_force()

    def _build_ui(self):
        root = ttk.Frame(self, padding=PAD)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(root, text="Ship Report Builder", font=("", 15, "bold")).grid(
            row=row, column=0, columnspan=3, sticky="w", pady=(0, PAD)
        )
        row += 1

        # 1. Sample report ------------------------------------------------
        self._file_row(
            root, row, "Sample report (template):", self.sample_var,
            lambda: self._pick_file(self.sample_var, "Select sample report",
                                    [("Excel", "*.xlsx *.xlsm *.xls")]),
        )
        row += 1

        # 2. PO Number ----------------------------------------------------
        ttk.Label(root, text="PO Number (optional):").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(root, textvariable=self.po_var).grid(row=row, column=1, columnspan=2, sticky="ew", pady=4)
        row += 1

        # 3. Build Code ---------------------------------------------------
        ttk.Label(root, text="Build Code (optional):").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(root, textvariable=self.build_var).grid(row=row, column=1, columnspan=2, sticky="ew", pady=4)
        row += 1

        # 4. Raw-data folder ---------------------------------------------
        self._file_row(
            root, row, "Raw-data folder:", self.raw_dir_var,
            lambda: self._pick_dir(self.raw_dir_var, "Select raw-data folder"),
        )
        row += 1

        # 5. OK-TO-SHIP report -------------------------------------------
        self._file_row(
            root, row, "OK-TO-SHIP report (optional):", self.ok_ship_var,
            lambda: self._pick_file(self.ok_ship_var, "Select OK-TO-SHIP report",
                                    [("Excel", "*.xlsx *.xlsm *.xls")]),
        )
        row += 1

        # Device lookup (uses the raw-data folder + optional PO / Build Code) --
        ttk.Separator(root, orient="horizontal").grid(
            row=row, column=0, columnspan=3, sticky="ew", pady=(PAD, 4))
        row += 1
        ttk.Label(root, text="Look up device (IMEI/DevEUI):").grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(root, textvariable=self.lookup_var).grid(row=row, column=1, sticky="ew", pady=4, padx=(0, PAD))
        ttk.Button(root, text="Look Up", command=self._on_lookup).grid(row=row, column=2, sticky="e", pady=4)
        row += 1

        # Output ----------------------------------------------------------
        self._file_row(
            root, row, "Save report as:", self.output_var,
            lambda: self._pick_save(self.output_var),
        )
        row += 1

        # Generate button -------------------------------------------------
        self.generate_btn = ttk.Button(root, text="Generate Report", command=self._on_generate)
        self.generate_btn.grid(row=row, column=0, columnspan=3, pady=(PAD, 4), sticky="ew")
        row += 1

        # Log -------------------------------------------------------------
        ttk.Label(root, text="Progress:").grid(row=row, column=0, sticky="w")
        row += 1
        self.log = tk.Text(root, height=8, wrap="word", state="disabled")
        self.log.grid(row=row, column=0, columnspan=3, sticky="nsew")
        root.rowconfigure(row, weight=1)

    def _file_row(self, root, row, label, var, cmd):
        ttk.Label(root, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(root, textvariable=var).grid(row=row, column=1, sticky="ew", pady=4, padx=(0, PAD))
        ttk.Button(root, text="Browse…", command=cmd).grid(row=row, column=2, sticky="e", pady=4)

    # ------------------------------------------------------------------ #
    # Pickers
    # ------------------------------------------------------------------ #
    def _pick_file(self, var, title, filetypes):
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if path:
            var.set(path)

    def _pick_dir(self, var, title):
        path = filedialog.askdirectory(title=title)
        if path:
            var.set(path)

    def _pick_save(self, var):
        path = filedialog.asksaveasfilename(
            title="Save report as", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            var.set(path)

    # ------------------------------------------------------------------ #
    # Generate (runs off the UI thread)
    # ------------------------------------------------------------------ #
    def _on_generate(self):
        if not self.sample_var.get():
            messagebox.showerror(APP_TITLE, "Please select a sample report (the template).")
            return
        if not self.raw_dir_var.get():
            messagebox.showerror(APP_TITLE, "Please select the raw-data folder.")
            return
        if not (self.po_var.get().strip() or self.build_var.get().strip()):
            messagebox.showerror(APP_TITLE, "Please enter a PO Number and/or Build Code.")
            return
        if not self.output_var.get():
            messagebox.showerror(APP_TITLE, "Please choose where to save the report.")
            return

        req = ReportRequest(
            sample_path=Path(self.sample_var.get()),
            raw_reports_dir=Path(self.raw_dir_var.get()),
            output_path=Path(self.output_var.get()),
            po_number=self.po_var.get().strip(),
            build_code=self.build_var.get().strip(),
            ok_to_ship_path=Path(self.ok_ship_var.get()) if self.ok_ship_var.get() else None,
        )

        self.generate_btn.config(state="disabled")
        self._clear_log()
        threading.Thread(target=self._worker, args=(req,), daemon=True).start()

    # ------------------------------------------------------------------ #
    # Device lookup
    # ------------------------------------------------------------------ #
    def _on_lookup(self):
        device_id = self.lookup_var.get().strip()
        if not device_id:
            messagebox.showerror(APP_TITLE, "Enter a device ID (IMEI/DevEUI) to look up.")
            return
        if not self.raw_dir_var.get():
            messagebox.showerror(APP_TITLE, "Select the raw-data folder first.")
            return
        raw_dir = self.raw_dir_var.get()
        po, bc = self.po_var.get().strip(), self.build_var.get().strip()

        def work():
            try:
                dl = lookup_device(raw_dir, device_id, po, bc)
                self._msg_q.put(("lookup", format_lookup(dl, device_id)))
            except Exception:
                self._msg_q.put(("lookup", traceback.format_exc()))

        threading.Thread(target=work, daemon=True).start()

    def _show_lookup_result(self, text: str):
        win = tk.Toplevel(self)
        win.title("Device Lookup")
        win.minsize(640, 320)
        txt = tk.Text(win, wrap="none", font=("Menlo", 11))
        yscroll = ttk.Scrollbar(win, orient="vertical", command=txt.yview)
        xscroll = ttk.Scrollbar(win, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        txt.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        win.rowconfigure(0, weight=1)
        win.columnconfigure(0, weight=1)
        txt.insert("1.0", text)
        txt.config(state="disabled")

    def _worker(self, req: ReportRequest):
        try:
            generate_report(req, progress=lambda m: self._msg_q.put(("log", m)))
            self._msg_q.put(("done", "Report generated successfully."))
        except Exception:
            self._msg_q.put(("error", traceback.format_exc()))

    # ------------------------------------------------------------------ #
    # Cross-thread UI updates
    # ------------------------------------------------------------------ #
    def _drain_queue(self):
        try:
            while True:
                kind, msg = self._msg_q.get_nowait()
                if kind == "log":
                    self._append_log(msg)
                elif kind == "done":
                    self._append_log(msg)
                    self.generate_btn.config(state="normal")
                    messagebox.showinfo(APP_TITLE, msg)
                elif kind == "error":
                    self._append_log(msg)
                    self.generate_btn.config(state="normal")
                    messagebox.showerror(APP_TITLE, "Failed. See progress log for details.")
                elif kind == "lookup":
                    self._show_lookup_result(msg)
        except queue.Empty:
            pass
        self.after(100, self._drain_queue)

    def _append_log(self, text: str):
        self.log.config(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
