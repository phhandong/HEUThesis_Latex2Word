from __future__ import annotations

import sys
from pathlib import Path

from .report import ConversionReport


def _update_header_footer_fields(document) -> None:
    for section_index in range(1, document.Sections.Count + 1):
        section = document.Sections(section_index)
        for collection in (section.Headers, section.Footers):
            for kind in (1, 2, 3):
                try:
                    collection(kind).Range.Fields.Update()
                except Exception:
                    continue


def finalize_word_fields(path: str | Path, report: ConversionReport) -> bool:
    """Update and save layout-dependent Word fields when desktop Word is available."""
    if sys.platform != "win32":
        report.warn("当前平台没有 Word COM；目录域已保留，但未缓存实际目录结果。")
        return False

    try:
        import pythoncom
        import win32com.client
    except ImportError:
        report.warn("未安装 pywin32；目录域已保留，但未缓存实际目录结果。")
        return False

    document_path = Path(path).resolve()
    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.ScreenUpdating = False
        try:
            word.AutomationSecurity = 3
            word.Options.UpdateLinksAtOpen = False
        except Exception:
            pass

        document = word.Documents.Open(
            str(document_path),
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=False,
            NoEncodingDialog=True,
        )

        # First pass materializes TOC entries. Second pass refreshes page numbers
        # after the populated TOC changes pagination.
        for _ in range(2):
            for toc_index in range(1, document.TablesOfContents.Count + 1):
                document.TablesOfContents(toc_index).Update()
            document.Fields.Update()
            _update_header_footer_fields(document)
            document.Repaginate()
            for toc_index in range(1, document.TablesOfContents.Count + 1):
                document.TablesOfContents(toc_index).UpdatePageNumbers()

        document.Repaginate()
        document.Save()
        report.note("已由 Microsoft Word 更新并缓存目录、页码和交叉引用域。")
        return True
    except Exception as exc:
        report.warn(f"Word 域缓存失败，文档仍保留可更新域：{exc}")
        return False
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.NormalTemplate.Saved = True
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
