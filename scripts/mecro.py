import tkinter as tk
from tkinter import messagebox
import threading
import time

import pyautogui
import keyboard  # 전역 단축키(F11/F12)

# -----------------------
# 상태 변수
# -----------------------
running = False
worker_thread = None
stop_event = threading.Event()

last_mouse_pos = None
lock = threading.Lock()


# -----------------------
# 유틸
# -----------------------
def safe_set_status(text: str):
    # tkinter는 메인 스레드에서만 UI 업데이트 권장
    root.after(0, lambda: status_label.config(text=text))


def get_interval_sec() -> float:
    s = interval_var.get().strip()
    if not s:
        raise ValueError("주기를 입력하세요. (초 단위)")
    v = float(s)
    if v <= 0:
        raise ValueError("주기는 0보다 커야 합니다.")
    return v


def stop_clicking(reason: str = ""):
    global running
    with lock:
        if not running:
            return
        running = False
        stop_event.set()

    if reason:
        print(f"[STOP] {reason}")
    safe_set_status(f"정지{(' - ' + reason) if reason else ''}")


def start_clicking():
    global running, worker_thread, last_mouse_pos

    with lock:
        if running:
            return

        try:
            interval = get_interval_sec()
        except Exception as e:
            messagebox.showerror("입력 오류", str(e))
            return

        running = True
        stop_event.clear()
        last_mouse_pos = pyautogui.position()

    safe_set_status("동작중 (F11 정지 / 마우스 움직이면 정지)")
    print(f"[START] interval={interval}s")

    worker_thread = threading.Thread(target=click_loop, daemon=True)
    worker_thread.start()


# -----------------------
# 메인 루프
# -----------------------
def click_loop():
    """
    - interval마다 좌클릭 1회
    - 마우스가 움직이면 즉시 정지
    """
    global last_mouse_pos

    try:
        interval = get_interval_sec()
    except Exception:
        # 시작 시점에 검증했지만, 혹시 UI에서 변경되는 경우 대비
        stop_clicking("주기 값 오류")
        return

    next_click_time = time.time()

    while not stop_event.is_set():
        # 1) 마우스 움직임 감지 -> 정지
        pos = pyautogui.position()
        if last_mouse_pos is not None and pos != last_mouse_pos:
            stop_clicking("마우스 움직임 감지")
            break

        # 2) 클릭 타이밍
        now = time.time()
        if now >= next_click_time:
            # 클릭 직전에 다시 마우스 위치 기록(“내가 클릭해서” 움직였다고 오탐 방지)
            # 일반 좌클릭은 위치가 안 변하므로 큰 의미는 없지만 안정적으로 유지
            last_mouse_pos = pyautogui.position()

            pyautogui.click(button="left")
            print("[CLICK]")
            next_click_time = now + interval

        # CPU 과점유 방지
        time.sleep(0.01)


# -----------------------
# 단축키 핸들러
# -----------------------
def on_hotkey_start():
    # F12
    root.after(0, start_clicking)


def on_hotkey_stop():
    # F11
    root.after(0, lambda: stop_clicking("F11"))


# -----------------------
# UI
# -----------------------
root = tk.Tk()
root.title("Auto Clicker (F12 Start / F11 Stop)")
root.geometry("360x190")
root.resizable(False, False)

title = tk.Label(root, text="자동 좌클릭", font=("Arial", 14, "bold"))
title.pack(pady=(12, 6))

frame = tk.Frame(root)
frame.pack(pady=6)

tk.Label(frame, text="주기(초):", font=("Arial", 11)).grid(row=0, column=0, padx=6, pady=6, sticky="e")

interval_var = tk.StringVar(value="60")
interval_entry = tk.Entry(frame, textvariable=interval_var, width=10, font=("Arial", 11))
interval_entry.grid(row=0, column=1, padx=6, pady=6, sticky="w")

hint = tk.Label(root, text="F12 시작 / F11 정지 / 마우스 움직이면 자동 정지", font=("Arial", 10))
hint.pack(pady=(0, 8))

btns = tk.Frame(root)
btns.pack(pady=4)

start_btn = tk.Button(btns, text="시작(F12)", width=12, command=start_clicking)
start_btn.grid(row=0, column=0, padx=8)

stop_btn = tk.Button(btns, text="정지(F11)", width=12, command=lambda: stop_clicking("버튼"))
stop_btn.grid(row=0, column=1, padx=8)

status_label = tk.Label(root, text="정지", font=("Arial", 11))
status_label.pack(pady=10)


def on_close():
    # 종료 시 단축키 해제 + 스레드 종료
    try:
        keyboard.unhook_all_hotkeys()
    except Exception:
        pass
    stop_clicking("프로그램 종료")
    root.destroy()


root.protocol("WM_DELETE_WINDOW", on_close)

# 전역 단축키 등록
keyboard.add_hotkey("F12", on_hotkey_start)
keyboard.add_hotkey("F11", on_hotkey_stop)

root.mainloop()