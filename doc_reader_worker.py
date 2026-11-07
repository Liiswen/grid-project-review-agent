# -*- coding: utf-8 -*-
"""在独立进程中读取旧版 Word DOC，避免 Office COM 阻塞主审查进程。"""
import sys


def main():
    if len(sys.argv) != 2:
        return 2
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            sys.argv[1], ConfirmConversions=False, ReadOnly=True,
            AddToRecentFiles=False, Visible=False, OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        # Content 覆盖段落及表格正文；表单型审批单通常把字段放在表格中。
        text = document.Content.Text or ''
        if text.strip():
            sys.stdout.write(text)
            return 0
        return 3
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    raise SystemExit(main())
