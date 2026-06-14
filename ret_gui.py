"""Tkinter GUI for generating RETConfigWDTInternal from a CDD file.

Run:  uv run ret_gui.py
"""
import os
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import ret_core

# When frozen by PyInstaller, bundled data lives in sys._MEIPASS, while the
# folder the user actually sees (for default output) is next to the .exe.
if getattr(sys, "frozen", False):
    BUNDLE_DIR = sys._MEIPASS  # noqa: SLF001 - PyInstaller resource root
    APP_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))
    APP_DIR = BUNDLE_DIR

HERE = APP_DIR


def _default(path):
    """Default input path: prefer a file next to the app, else the bundled copy."""
    for base in (APP_DIR, BUNDLE_DIR):
        full = os.path.join(base, path)
        if os.path.exists(full):
            return full
    return ""


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RET Config Generator")
        self.geometry("960x640")
        self.minsize(820, 560)

        self.cdd_var = tk.StringVar(value=_default("CDD.xlsx"))
        self.sheet_var = tk.StringVar()
        self.mapping_var = tk.StringVar(value=_default("mapping.json"))
        self.template_var = tk.StringVar(value=_default("RETConfigWDTInternal.xlsx"))
        self.output_var = tk.StringVar(value=os.path.join(HERE, "RETConfigWDTInternal_new.xlsx"))
        self.status_var = tk.StringVar(value="Select a CDD file to begin.")

        # Text-conversion tab (RET_template.txt + RET_input.txt -> RET_output.txt).
        self.txt_input_var = tk.StringVar(value=_default("RET_input.txt"))
        self.txt_template_var = tk.StringVar(value=_default("RET_template.txt"))
        self.txt_output_var = tk.StringVar(value=os.path.join(HERE, "RET_output.txt"))
        self.txt_status_var = tk.StringVar(value="Uses the CDD, Sheet and Mapping from the first tab.")

        self._rows = []  # last previewed rows

        self._build()
        if self.cdd_var.get():
            self._load_sheets()

    # ---------- layout ----------
    def _build(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True)

        self.tab_xlsx = ttk.Frame(self.nb)
        self.tab_text = ttk.Frame(self.nb)
        self.nb.add(self.tab_xlsx, text="CDD → WDT (xlsx)")
        self.nb.add(self.tab_text, text="RET MML (txt)")

        self._build_xlsx_tab(self.tab_xlsx)
        self._build_text_tab(self.tab_text)

    def _build_xlsx_tab(self, parent):
        pad = {"padx": 8, "pady": 4}

        # --- 1. Inputs ---
        inp = ttk.LabelFrame(parent, text="1. Inputs")
        inp.pack(fill="x", **pad)
        inp.columnconfigure(1, weight=1)

        self._file_row(inp, 0, "CDD file:", self.cdd_var, self._browse_cdd)

        ttk.Label(inp, text="Sheet:").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self.sheet_combo = ttk.Combobox(inp, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=4)

        self._file_row(inp, 2, "Mapping (config):", self.mapping_var,
                       lambda: self._browse_into(self.mapping_var, "JSON", "*.json"))
        self._file_row(inp, 3, "Template (style):", self.template_var,
                       lambda: self._browse_into(self.template_var, "Excel", "*.xlsx"))

        # --- 2. Preview ---
        prev = ttk.LabelFrame(parent, text="2. Preview")
        prev.pack(fill="both", expand=True, **pad)

        bar = ttk.Frame(prev)
        bar.pack(fill="x")
        ttk.Button(bar, text="Validate / Preview", command=self.preview).pack(side="left", padx=4, pady=4)
        ttk.Label(bar, textvariable=self.status_var).pack(side="left", padx=10)

        cols = ret_core.HEADERS
        self.tree = ttk.Treeview(prev, columns=cols, show="headings", height=12)
        widths = [120, 200, 70, 70, 70, 100, 70, 200]
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="w")
        vsb = ttk.Scrollbar(prev, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        vsb.pack(side="left", fill="y", pady=4)

        # --- 3. Output ---
        outp = ttk.LabelFrame(parent, text="3. Output")
        outp.pack(fill="x", **pad)
        outp.columnconfigure(1, weight=1)
        self._file_row(outp, 0, "Save to:", self.output_var, self._browse_output, save=True)

        actions = ttk.Frame(outp)
        actions.grid(row=1, column=0, columnspan=3, sticky="e", padx=8, pady=6)
        ttk.Button(actions, text="Generate", command=self.generate).pack(side="right")

    def _file_row(self, parent, row, label, var, cmd, save=False):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", padx=8, pady=4)
        ttk.Button(parent, text="Browse…", command=cmd).grid(row=row, column=2, padx=8, pady=4)

    def _build_text_tab(self, parent):
        pad = {"padx": 8, "pady": 4}

        info = ttk.Label(
            parent,
            text=("Rewrites RET_template.txt into an output MML script:\n"
                  "  • DEVICENAME prefix → {SiteName_New}_{Ne ID} (from CDD)\n"
                  "  • SERIALNO → input serial matched by CTRLSRN + last 3 chars\n"
                  "  • MOD RETTILT TILT → RCU Tilt of the same device (from CDD)"),
            justify="left",
        )
        info.pack(fill="x", **pad)

        inp = ttk.LabelFrame(parent, text="1. Inputs")
        inp.pack(fill="x", **pad)
        inp.columnconfigure(1, weight=1)

        # CDD / Sheet / Mapping share the same vars as the first tab, so setting
        # them on either tab keeps both in sync.
        self._file_row(inp, 0, "CDD file:", self.cdd_var, self._browse_cdd)
        ttk.Label(inp, text="Sheet:").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        self.text_sheet_combo = ttk.Combobox(inp, textvariable=self.sheet_var, state="readonly")
        self.text_sheet_combo.grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        self._file_row(inp, 2, "Mapping (config):", self.mapping_var,
                       lambda: self._browse_into(self.mapping_var, "JSON", "*.json"))

        self._file_row(inp, 3, "RET input (.txt):", self.txt_input_var,
                       lambda: self._browse_into(self.txt_input_var, "Text", "*.txt"))
        self._file_row(inp, 4, "RET template (.txt):", self.txt_template_var,
                       lambda: self._browse_into(self.txt_template_var, "Text", "*.txt"))

        prev = ttk.LabelFrame(parent, text="2. Preview")
        prev.pack(fill="both", expand=True, **pad)
        bar = ttk.Frame(prev)
        bar.pack(fill="x")
        ttk.Button(bar, text="Preview", command=self.text_preview).pack(side="left", padx=4, pady=4)
        ttk.Button(bar, text="Copy", command=self.text_copy).pack(side="left", padx=4, pady=4)
        ttk.Label(bar, textvariable=self.txt_status_var).pack(side="left", padx=10)

        self.txt_preview = tk.Text(prev, height=14, wrap="none")
        ysb = ttk.Scrollbar(prev, orient="vertical", command=self.txt_preview.yview)
        xsb = ttk.Scrollbar(prev, orient="horizontal", command=self.txt_preview.xview)
        self.txt_preview.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.txt_preview.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)
        ysb.pack(side="left", fill="y", pady=4)
        xsb.pack(side="bottom", fill="x")

        outp = ttk.LabelFrame(parent, text="3. Output")
        outp.pack(fill="x", **pad)
        outp.columnconfigure(1, weight=1)
        self._file_row(outp, 0, "Save to:", self.txt_output_var, self._browse_text_output, save=True)
        actions = ttk.Frame(outp)
        actions.grid(row=1, column=0, columnspan=3, sticky="e", padx=8, pady=6)
        ttk.Button(actions, text="Generate", command=self.text_generate).pack(side="right")

    def _browse_text_output(self):
        path = filedialog.asksaveasfilename(
            title="Save output as", defaultextension=".txt",
            filetypes=[("Text", "*.txt")], initialfile="RET_output.txt",
        )
        if path:
            self.txt_output_var.set(path)

    def _gather_text(self):
        """Validate inputs and return (mapping, sheet, text, warnings)."""
        if not os.path.exists(self.txt_input_var.get()):
            raise ValueError("Please select a valid RET input (.txt) file.")
        if not os.path.exists(self.txt_template_var.get()):
            raise ValueError("Please select a valid RET template (.txt) file.")
        if not self.cdd_var.get() or not os.path.exists(self.cdd_var.get()):
            raise ValueError("Select a valid CDD file on the first tab.")
        if not self.sheet_var.get():
            raise ValueError("Select a sheet on the first tab.")
        if not os.path.exists(self.mapping_var.get()):
            raise ValueError("Mapping file not found (first tab).")
        mapping = ret_core.load_mapping(self.mapping_var.get())
        text, warnings = ret_core.build_text_output(
            self.txt_template_var.get(), self.txt_input_var.get(),
            self.cdd_var.get(), self.sheet_var.get(), mapping,
        )
        return mapping, text, warnings

    def text_preview(self):
        try:
            _, text, warnings = self._gather_text()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self.txt_preview.delete("1.0", "end")
        self.txt_preview.insert("1.0", text)
        self.txt_status_var.set(
            "Preview ready." + (f"  ·  {len(warnings)} warning(s)" if warnings else "")
        )
        if warnings:
            self._show_warnings(warnings)

    def text_copy(self):
        """Copy the whole MML preview to the clipboard."""
        text = self.txt_preview.get("1.0", "end-1c")
        if not text.strip():
            messagebox.showinfo("Nothing to copy", "Click Preview (or Generate) first.")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()  # keep the clipboard populated after the app yields
        self.txt_status_var.set("Copied preview to clipboard.")

    def text_generate(self):
        try:
            _, text, warnings = self._gather_text()
            with open(self.txt_output_var.get(), "w", encoding="utf-8") as f:
                f.write(text)
        except Exception as e:
            messagebox.showerror("Generation failed", f"{e}\n\n{traceback.format_exc()}")
            return
        self.txt_preview.delete("1.0", "end")
        self.txt_preview.insert("1.0", text)
        self.txt_status_var.set(f"Wrote → {os.path.basename(self.txt_output_var.get())}")
        if warnings:
            self._show_warnings(warnings)
        if messagebox.askyesno(
            "Done",
            f"Wrote:\n{self.txt_output_var.get()}\n\nOpen containing folder?",
        ):
            self._open_folder(os.path.dirname(self.txt_output_var.get()))

    @staticmethod
    def _show_warnings(warnings):
        preview = "\n".join(warnings[:20])
        more = "" if len(warnings) <= 20 else f"\n… +{len(warnings) - 20} more"
        messagebox.showwarning("Warnings", preview + more)

    # ---------- actions ----------
    def _browse_cdd(self):
        path = filedialog.askopenfilename(
            title="Select CDD file",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.cdd_var.set(path)
            self._load_sheets()

    def _browse_into(self, var, label, pattern):
        path = filedialog.askopenfilename(filetypes=[(label, pattern), ("All files", "*.*")])
        if path:
            var.set(path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save output as",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="RETConfigWDTInternal_new.xlsx",
        )
        if path:
            self.output_var.set(path)

    def _load_sheets(self):
        try:
            sheets = ret_core.list_sheets(self.cdd_var.get())
        except Exception as e:
            messagebox.showerror("Cannot read CDD", str(e))
            return
        self.sheet_combo["values"] = sheets
        # Keep the text tab's sheet dropdown in sync (shares self.sheet_var).
        if getattr(self, "text_sheet_combo", None) is not None:
            self.text_sheet_combo["values"] = sheets
        # Prefer the sheet named in mapping.json, else first.
        preferred = None
        try:
            m = ret_core.load_mapping(self.mapping_var.get())
            preferred = m.get("source", {}).get("sheet")
        except Exception:
            pass
        self.sheet_var.set(preferred if preferred in sheets else (sheets[0] if sheets else ""))
        self.status_var.set(f"Loaded {len(sheets)} sheet(s). Click Validate / Preview.")

    def _gather(self):
        if not self.cdd_var.get() or not os.path.exists(self.cdd_var.get()):
            raise ValueError("Please select a valid CDD file.")
        if not self.sheet_var.get():
            raise ValueError("Please select a sheet.")
        if not os.path.exists(self.mapping_var.get()):
            raise ValueError("Mapping file not found.")
        mapping = ret_core.load_mapping(self.mapping_var.get())
        rows, skipped, sectors = ret_core.build_rows(
            self.cdd_var.get(), self.sheet_var.get(), mapping
        )
        return mapping, rows, skipped, sectors

    def preview(self):
        try:
            _, rows, skipped, sectors = self._gather()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._rows = rows
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", values=r)
        msg = f"{len(rows)} rows · {sectors} sectors"
        if skipped:
            msg += f"  ·  {len(skipped)} skipped"
        self.status_var.set(msg)
        if skipped:
            preview = "\n".join(f"{c}  →  {why}" for c, why in skipped[:20])
            more = "" if len(skipped) <= 20 else f"\n… +{len(skipped) - 20} more"
            messagebox.showwarning("Skipped rows", preview + more)

    def generate(self):
        try:
            mapping, rows, skipped, sectors = self._gather()
            if not rows:
                messagebox.showwarning("Nothing to write", "No valid rows were produced.")
                return
            target_sheet = mapping.get("target", {}).get("sheet", "Internal")
            ret_core.write_output(
                self.template_var.get(), target_sheet, rows, self.output_var.get()
            )
        except Exception as e:
            messagebox.showerror("Generation failed", f"{e}\n\n{traceback.format_exc()}")
            return
        self._rows = rows
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert("", "end", values=r)
        self.status_var.set(f"Wrote {len(rows)} rows → {os.path.basename(self.output_var.get())}")
        extra = f"\n{len(skipped)} CDD row(s) skipped." if skipped else ""
        if messagebox.askyesno(
            "Done",
            f"Wrote {len(rows)} rows ({sectors} sectors) to:\n{self.output_var.get()}{extra}"
            "\n\nOpen containing folder?",
        ):
            self._open_folder(os.path.dirname(self.output_var.get()))

    @staticmethod
    def _open_folder(path):
        import subprocess
        import sys
        try:
            if sys.platform.startswith("linux"):
                subprocess.Popen(["xdg-open", path])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            elif sys.platform.startswith("win"):
                os.startfile(path)  # noqa: SLF001
        except Exception:
            pass


if __name__ == "__main__":
    App().mainloop()
