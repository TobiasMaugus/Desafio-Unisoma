"""Windows launcher for the local ABC Aprendiz Streamlit application."""

from __future__ import annotations

import ctypes
import os
from pathlib import Path
import subprocess
import sys
import time
import urllib.error
import urllib.request
import webbrowser

PORT = 8517
URL = f"http://127.0.0.1:{PORT}"
APP_TITLE = "ABC Aprendiz · Ensalamento"
IDENTITY_FILE = "abc-aprendiz-identity.txt"
IDENTITY_VALUE = "unisoma-abc-aprendiz-ensalamento-v1"
MUTEX_NAME = "Local\\ABC_Aprendiz_Ensalamento_8517"


def base_directory() -> Path:
    if getattr(sys, "frozen", False):
        unpacked = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        if (unpacked / "streamlit_app.py").exists():
            return unpacked
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def probe_server() -> str:
    """Return app, other, or unavailable after checking both health and identity."""
    try:
        with urllib.request.urlopen(f"{URL}/_stcore/health", timeout=0.45) as response:
            if response.status != 200:
                return "other"
    except (urllib.error.URLError, TimeoutError, OSError):
        return "unavailable"
    try:
        with urllib.request.urlopen(f"{URL}/app/static/{IDENTITY_FILE}", timeout=0.6) as response:
            identity = response.read(256).decode("ascii", errors="replace").strip()
        return "app" if identity == IDENTITY_VALUE else "other"
    except urllib.error.HTTPError:
        return "other"
    except (urllib.error.URLError, TimeoutError, OSError):
        return "unavailable"


def acquire_mutex():
    if os.name != "nt":
        return None, True
    kernel = ctypes.windll.kernel32
    kernel.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
    kernel.CreateMutexW.restype = ctypes.c_void_p
    kernel.CloseHandle.argtypes = [ctypes.c_void_p]
    handle = kernel.CreateMutexW(None, True, MUTEX_NAME)
    if not handle:
        return None, False
    already_exists = kernel.GetLastError() == 183
    if already_exists:
        kernel.CloseHandle(handle)
        return None, False
    return handle, not already_exists


def release_mutex(handle):
    if handle and os.name == "nt":
        ctypes.windll.kernel32.ReleaseMutex(handle)
        ctypes.windll.kernel32.CloseHandle(handle)


def run_server_child():
    base = base_directory()
    app = base / "streamlit_app.py"
    if not app.exists():
        raise FileNotFoundError("Não encontrei streamlit_app.py junto ao launcher.")
    from streamlit.web.cli import main as streamlit_main

    sys.argv = ["streamlit", "run", str(app),
                "--global.developmentMode=false",
                "--server.address=127.0.0.1", f"--server.port={PORT}",
                "--server.headless=true", "--server.fileWatcherType=none",
                "--browser.gatherUsageStats=false"]
    os.chdir(base)
    streamlit_main()


def child_command() -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, "--streamlit-child"]
    return [sys.executable, str(Path(__file__).resolve()), "--streamlit-child"]


def show_error(message: str):
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("ABC Aprendiz · Ensalamento", message, parent=root)
    root.destroy()


def launch():
    initial = probe_server()
    if initial == "app":
        webbrowser.open(URL)
        return
    if initial == "other":
        show_error(f"A porta {PORT} está ocupada por outro serviço. Feche esse programa e tente novamente.")
        return

    mutex, owner = acquire_mutex()
    if not owner:
        show_error("A aplicação já está sendo iniciada. Aguarde alguns segundos e tente novamente.")
        return

    import tkinter as tk
    from tkinter import ttk

    root = tk.Tk()
    root.title("ABC Aprendiz")
    root.geometry("420x150")
    root.resizable(False, False)
    root.configure(background="#f5f7f9")
    ttk.Label(root, text="ABC Aprendiz", font=("Segoe UI", 18, "bold")).pack(pady=(22, 4))
    ttk.Label(root, text="Iniciando o sistema de ensalamento…", font=("Segoe UI", 10)).pack(pady=3)
    progress = ttk.Progressbar(root, mode="indeterminate", length=340)
    progress.pack(pady=14)
    progress.start(10)

    child = None
    started = time.monotonic()
    try:
        if probe_server() == "unavailable":
            environment = os.environ.copy()
            if getattr(sys, "frozen", False):
                # The server outlives this launcher and needs its own ONEFILE extraction.
                environment["PYINSTALLER_RESET_ENVIRONMENT"] = "1"
            child = subprocess.Popen(
                child_command(), cwd=str(Path(sys.executable).parent if getattr(sys, "frozen", False)
                                         else base_directory()), env=environment, stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0)

        def poll():
            state = probe_server()
            if state == "app":
                progress.stop()
                root.destroy()
                release_mutex(mutex)
                webbrowser.open(URL)
                return
            if state == "other":
                progress.stop()
                root.destroy()
                release_mutex(mutex)
                show_error(f"A porta {PORT} foi ocupada por outro serviço. O launcher não irá abri-lo.")
                return
            if child is not None and child.poll() is not None:
                progress.stop()
                root.destroy()
                release_mutex(mutex)
                show_error("O servidor não iniciou. Confira os arquivos do aplicativo ou tente executá-lo novamente.")
                return
            if time.monotonic() - started > 60:
                if child is not None:
                    child.terminate()
                progress.stop()
                root.destroy()
                release_mutex(mutex)
                show_error("O sistema demorou demais para iniciar. Tente novamente.")
                return
            root.after(500, poll)

        root.after(200, poll)
        root.mainloop()
    except Exception as exc:
        if child is not None and child.poll() is None:
            child.terminate()
        release_mutex(mutex)
        try:
            root.destroy()
        except tk.TclError:
            pass
        show_error(f"Não foi possível iniciar a aplicação: {exc}")


if __name__ == "__main__":
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    if "--streamlit-child" in sys.argv:
        run_server_child()
    else:
        launch()
