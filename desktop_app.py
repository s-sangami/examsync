import sys
import os
import threading
import time
import socket

# Note: uploads/outputs are stored by app.py (%LOCALAPPDATA%\ExamSync when
# frozen), so the working directory no longer matters.

# ── Find a free port ──────────────────────────────────────────────────────────
def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]

# ── Start Flask in background thread ─────────────────────────────────────────
def run_flask(port):
    from app import app
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import webview

    port = find_free_port()

    # Start Flask thread (daemon = stops when window closes)
    flask_thread = threading.Thread(target=run_flask, args=(port,), daemon=True)
    flask_thread.start()

    # Wait for Flask to be ready before opening window
    url = f"http://127.0.0.1:{port}"
    for _ in range(20):           # try for up to 4 seconds
        try:
            import urllib.request
            urllib.request.urlopen(url, timeout=0.3)
            break
        except Exception:
            time.sleep(0.2)

    # Create and open the desktop window (NO BROWSER — native window)
    window = webview.create_window(
        title   = "ExamSync — Attendance Management",
        url     = url,
        width   = 1100,
        height  = 720,
        min_size= (900, 600),
        resizable = True,
    )

    webview.start()
