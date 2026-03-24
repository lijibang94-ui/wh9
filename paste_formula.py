import argparse
import ctypes
import json
import sys
import time
from pathlib import Path
import tkinter as tk


VK_CONTROL = 0x11
VK_V = 0x56
VK_A = 0x41
KEYEVENTF_KEYUP = 0x0002
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004


def set_clipboard(text: str) -> None:
    root = tk.Tk()
    root.withdraw()
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update()
    root.destroy()


def tap_combo(*keys: int, pause: float = 0.05) -> None:
    user32 = ctypes.windll.user32
    for key in keys:
        user32.keybd_event(key, 0, 0, 0)
        time.sleep(pause)
    for key in reversed(keys):
        user32.keybd_event(key, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(pause)


def left_click(x: int, y: int, settle: float = 0.2) -> None:
    user32 = ctypes.windll.user32
    user32.SetCursorPos(x, y)
    time.sleep(settle)
    user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(settle)


def get_cursor_pos() -> tuple[int, int]:
    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    pt = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


def capture_position(label: str) -> dict:
    input(f"Move the mouse to [{label}] and press Enter here...")
    x, y = get_cursor_pos()
    print(f"Captured {label}: ({x}, {y})")
    return {"x": x, "y": y}


def load_config(path: Path) -> dict:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def save_config(path: Path, config: dict) -> None:
    path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


def calibrate(config_path: Path) -> int:
    print("Calibration will save three mouse positions:")
    print("1. formula editor text area")
    print("2. syntax check button")
    print("3. update main chart button")
    config = {
        "editor": capture_position("editor"),
        "syntax_check": capture_position("syntax_check"),
        "update_main_chart": capture_position("update_main_chart"),
    }
    save_config(config_path, config)
    print(f"Saved calibration to: {config_path}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Paste a Wenhua formula into the active editor, then click syntax check and update main chart."
    )
    parser.add_argument(
        "formula",
        nargs="?",
        default="hour_boll_range_lower_alert.txt",
        help="Formula file path. Defaults to hour_boll_range_lower_alert.txt in the current directory.",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=3.0,
        help="Seconds to wait after you press Enter so you can focus the Wenhua editor.",
    )
    parser.add_argument(
        "--check-wait",
        type=float,
        default=1.2,
        help="Seconds to wait after clicking syntax check before clicking update main chart.",
    )
    parser.add_argument(
        "--config",
        default="paste_formula_config.json",
        help="Path to the JSON config file that stores mouse positions.",
    )
    parser.add_argument(
        "--calibrate",
        action="store_true",
        help="Capture and save the mouse positions for the editor and the two buttons.",
    )
    args = parser.parse_args()

    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = Path.cwd() / config_path

    if args.calibrate:
        return calibrate(config_path)

    config = load_config(config_path)
    for key in ("editor", "syntax_check", "update_main_chart"):
        if key not in config:
            print(f"Missing [{key}] in config. Run with --calibrate first.")
            return 1

    formula_path = Path(args.formula)
    if not formula_path.is_absolute():
        formula_path = Path.cwd() / formula_path

    if not formula_path.exists():
        print(f"Formula file not found: {formula_path}")
        return 1

    text = formula_path.read_text(encoding="utf-8")
    if not text.strip():
        print(f"Formula file is empty: {formula_path}")
        return 1

    set_clipboard(text)
    print(f"Loaded formula into clipboard: {formula_path}")
    print("Switch to the Wenhua formula editor now.")
    input(f"Press Enter here, then you will have {args.delay:.1f} seconds to focus the editor...")
    time.sleep(args.delay)

    left_click(config["editor"]["x"], config["editor"]["y"])
    tap_combo(VK_CONTROL, VK_A)
    time.sleep(0.1)
    tap_combo(VK_CONTROL, VK_V)
    time.sleep(0.2)

    left_click(config["syntax_check"]["x"], config["syntax_check"]["y"])
    time.sleep(args.check_wait)
    left_click(config["update_main_chart"]["x"], config["update_main_chart"]["y"])

    print("Paste, syntax check, and update main chart completed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
