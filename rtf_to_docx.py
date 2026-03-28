"""
RTF to DOCX for the format_docx pipeline (Microsoft Word COM only, Windows).
"""
from __future__ import annotations

from pathlib import Path

_WD_FORMAT_XML_DOCUMENT = 12


def convert_rtf_to_docx(rtf_path: Path | str, docx_path: Path | str) -> None:
    """Convert RTF to DOCX using Microsoft Word via pywin32. No external fallback."""
    try:
        import win32com.client
    except ImportError as e:
        raise RuntimeError("需要 pywin32：pip install pywin32") from e
    rtf_path = Path(rtf_path).resolve()
    docx_path = Path(docx_path).resolve()
    if not rtf_path.is_file():
        raise FileNotFoundError(f"找不到 RTF：{rtf_path}")
    docx_path.parent.mkdir(parents=True, exist_ok=True)
    if docx_path.exists():
        docx_path.unlink()
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0
        except Exception:
            pass
        doc = word.Documents.Open(str(rtf_path), ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs2(str(docx_path), FileFormat=_WD_FORMAT_XML_DOCUMENT)
    except Exception as e:
        raise RuntimeError(f"Word 轉檔失敗：{e}") from e
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
