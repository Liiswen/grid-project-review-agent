# -*- coding: utf-8 -*-
"""Windows executable entry point for the local review service."""
import sys
import threading
import webbrowser


def main():
    if len(sys.argv) > 1 and sys.argv[1] == '--doc-reader':
        from doc_reader_worker import main as read_doc
        sys.argv = [sys.argv[0], *sys.argv[2:]]
        return read_doc()

    from app import app
    threading.Timer(1.0, lambda: webbrowser.open('http://127.0.0.1:5000')).start()
    app.run(debug=False, host='127.0.0.1', port=5000)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
