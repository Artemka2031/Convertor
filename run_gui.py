import subprocess, sys, time, webbrowser, urllib.request, importlib.util
from pathlib import Path

ROOT   = Path(__file__).resolve().parent
VENV_PY = ROOT / "venv" / "Scripts" / "python.exe"
PYTHON  = str(VENV_PY if VENV_PY.exists() else sys.executable)

REQUIRED = ["fastapi", "uvicorn", "pandas", "lxml", "openpyxl"]
missing = [p for p in REQUIRED if importlib.util.find_spec(p) is None]
if missing:
    print("⇢ Устанавливаю отсутствующие пакеты:", *missing)
    subprocess.check_call([PYTHON, "-m", "pip", "install", *missing])

CMD = [PYTHON, "-m", "uvicorn", "api.server:app", "--host", "127.0.0.1", "--port", "8000"]
print("⇢ стартую сервер:", " ".join(CMD))
server = subprocess.Popen(CMD, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)

def _up() -> bool:
    try: urllib.request.urlopen("http://127.0.0.1:8000", timeout=1); return True
    except Exception: return False

for _ in range(20):
    if _up(): break
    time.sleep(0.5)
else:
    print("❌ сервер не поднялся."); server.terminate(); sys.exit(1)

print("✅ сервер готов — открываю браузер."); webbrowser.open("http://127.0.0.1:8000")

try:
    for line in server.stdout:  # транслируем логи uvicorn
        print(line, end="")
except KeyboardInterrupt:
    print("⏹  Ctrl-C — останавливаю uvicorn…")
    server.terminate(); server.wait()
