"""图形界面入口（打包 exe 时以此为主程序）。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from gui.main_window import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
