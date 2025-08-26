import os
import PyInstaller.__main__

if __name__ == "__main__":
    PyInstaller.__main__.run([
        "app.py",
        "--name=ProposalGenerator",
        "--onefile",
        "--windowed",
        f"--add-data=Jost-Bold.ttf{os.pathsep}.",
        f"--add-data=Jost-Regular.ttf{os.pathsep}.",
        f"--add-data=logo.png{os.pathsep}.",
    ])
