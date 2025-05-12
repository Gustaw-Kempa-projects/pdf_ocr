from __future__ import annotations

import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import pandas as pd

from .excel_writer import dataframe_to_excel
from .pdf_parser import parse_pdf


def run_gui() -> None:
    """Fire up the tiny Tkinter picker."""
    root = tk.Tk()
    root.geometry("380x140")
    root.title("Orders-OCR")

    tk.Label(root, text="Wybierz pliki PDF z zamówieniami").pack(pady=15)
    tk.Button(root, text="Przeglądaj…", command=_pick_and_process).pack()
    root.mainloop()


# ----------------------------------------------------------------------------- #
# internals                                                                     #
# ----------------------------------------------------------------------------- #
def _pick_and_process() -> None:
    paths = filedialog.askopenfilenames(
        title="PDF-y do wczytania",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    if not paths:
        messagebox.showinfo("Info", "Nic nie zaznaczono.")
        return

    frames = [parse_pdf(Path(p)) for p in paths]
    merged = pd.concat(frames, ignore_index=True)
    out_path = dataframe_to_excel(merged)

    # fire up Excel (Windows) or open with default app elsewhere
    try:
        subprocess.Popen(["start", "excel", str(out_path)], shell=True)
    except FileNotFoundError:
        subprocess.Popen(["open", str(out_path)])


if __name__ == "__main__":
    run_gui()