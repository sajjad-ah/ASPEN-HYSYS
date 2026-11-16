import os, sys, webbrowser, socket, time, threading
from pathlib import Path

BASE = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)).resolve()
APP = BASE / "app_entrypoint.py"

def wait_for_server(host: str, port: int, timeout_s: int = 45) -> bool:
    t0 = time.time()
    while time.time() - t0 < timeout_s:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1.0)
            try:
                s.connect((host, port))
                return True
            except OSError:
                time.sleep(0.3)
    return False

def open_once_when_ready(host: str, ports):
    if os.environ.get("THERMOWELL_BROWSER_OPENED") == "1":
        return
    
    deadline = time.time() + 60
    while time.time() < deadline:
        for p in ports:
            if wait_for_server(host,p,timeout_s=1):
                webbrowser.open_new_tab(f"http://{host}:{p}")

                os.environ["THERMOWELL_BROWSER_OPENED"] = "1"
                return
        time.sleep(0.5)

def main():
    host = "127.0.0.1"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_SERVER_ADDRESS"] = host
    os.environ["STREAMLIT_SERVER_FILEWATCHERTYPE"] = "none"
    os.environ["STREAMLIT_SERVER_GATHERUSAGESTATS"] = "false"
    os.environ["STREAMLIT_SERVER_ENABLEXSRFPROTECTION"] = "false"
    os.environ["STREAMLIT_SERVER_ENABLECORS"] = "false"

    from streamlit.web.cli import main as st_main 

    threading.Thread(target=open_once_when_ready, args=(host, list(range(8501, 8511))), daemon=True).start()

    sys.argv = ["streamlit", "run", APP, "--server.headless", "true", "--server.fileWatcherType", "none"]

    st_main()

if __name__ == "__main__":
    main()