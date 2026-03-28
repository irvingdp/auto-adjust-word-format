"""
Thin GUI wrapper around format_docx.process().
Double-click the .exe → file dialog opens → pick a .docx → processed file saved next to it.
"""

import os
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

from format_docx import process
from rtf_to_docx import convert_rtf_to_docx


def main():
    root = tk.Tk()
    root.withdraw()

    src = filedialog.askopenfilename(
        title="選擇要轉換的檔案",
        filetypes=[
            ("Word / RTF", "*.docx;*.rtf"),
            ("Word", "*.docx"),
            ("RTF", "*.rtf"),
        ],
    )
    if not src:
        sys.exit(0)

    src_path = Path(src)
    dst_path = src_path.with_stem(src_path.stem + "_adjusted")

    try:
        if src_path.suffix.lower() == ".rtf":
            conv = src_path.parent / f"{src_path.stem}_converted.docx"
            convert_rtf_to_docx(src_path, conv)
            process(str(conv), str(dst_path))
        else:
            process(str(src_path), str(dst_path))
        messagebox.showinfo("完成", f"轉換完成！\n輸出檔案：\n{dst_path}")
    except Exception:
        messagebox.showerror("錯誤", traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()
