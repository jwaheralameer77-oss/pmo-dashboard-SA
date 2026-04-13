import subprocess, sys, os, shutil, webbrowser, threading

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PYTHON = sys.executable

def install_deps():
    print("Installing dependencies...")
    subprocess.check_call([PYTHON, "-m", "pip", "install", "--default-timeout=300", "-q", "flask", "flask-sqlalchemy"], timeout=300)
    print("Dependencies installed.")

def copy_logo():
    src = r"C:\Users\ameer\.verdent\artifacts\buckets\4bae9bb5-302f-4f09-9754-dbb0067de45e\images\1775991837070_eb8e8585.jpeg"
    dst = os.path.join(BASE_DIR, "static", "img", "logo.jpeg")
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if not os.path.exists(dst) and os.path.exists(src):
        shutil.copy2(src, dst)
        print("Logo copied.")

def open_browser():
    webbrowser.open("http://localhost:5000")

if __name__ == "__main__":
    os.chdir(BASE_DIR)
    install_deps()
    copy_logo()

    threading.Timer(2, open_browser).start()

    import subprocess
    subprocess.call([PYTHON, os.path.join(BASE_DIR, "app.py")])
