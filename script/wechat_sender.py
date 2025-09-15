import os
import sys
import time
import subprocess
from typing import List, Optional
import threading
import psutil

from pywinauto import Application, keyboard
from pywinauto import Desktop
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.timings import wait_until_passes
from pywinauto import timings
from pywinauto import mouse
try:
    import pyperclip  # 可选：用于剪贴板粘贴
except Exception:
    pyperclip = None


def _possible_wechat_paths() -> List[str]:
    paths = [
        os.path.expandvars(r"%PROGRAMFILES%\Tencent\Weixin\Weixin.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\WeChat\WeChat.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\Tencent\WeChat\WeChat.exe"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\Tencent\WeChat\WeChat.exe"),
    ]
    return [p for p in paths if p and os.path.isfile(p)]


VERBOSE = False
BACKEND = "uia"  # 默认使用 UIA，可切换为 win32 作为兜底


def _log(msg: str) -> None:
    if VERBOSE:
        try:
            print(msg)
        except Exception:
            try:
                print(str(msg).encode("utf-8", errors="ignore").decode("utf-8", errors="ignore"))
            except Exception:
                pass


def _ensure_utf8_console() -> None:
    # 尽量把标准输出/错误设置为 UTF-8，便于中文显示
    try:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        if hasattr(sys.stderr, "reconfigure"):
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def _window_area(rect) -> int:
    try:
        return max(0, rect.width() * rect.height())
    except Exception:
        return 0


def _safe_enum_windows(backend: str, timeout: float = 2.0):
    result = {"windows": None}

    def _worker():
        try:
            result["windows"] = Desktop(backend=backend).windows()
        except Exception:
            result["windows"] = []

    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    t.join(timeout)
    if t.is_alive():
        _log(f"枚举顶层窗口超时（backend={backend}）")
        return []
    return result["windows"] or []


def _find_weixin_main_window() -> Optional[object]:
    """在桌面枚举顶层窗口，选择最可能的微信主窗口。返回 WindowSpecification 或 None。"""
    top_windows = _safe_enum_windows(BACKEND, timeout=2.0)
    if not top_windows:
        # 再尝试另一种 backend 兜底
        alt_backend = "win32" if BACKEND == "uia" else "uia"
        top_windows = _safe_enum_windows(alt_backend, timeout=2.0)
        if not top_windows:
            return None

    candidates = []
    for w in top_windows:
        try:
            ei = w.element_info
            name = (ei.name or "")
            class_name = (ei.class_name or "")
            pid = getattr(ei, "process_id", None)
            proc_name = ""
            try:
                if pid:
                    proc_name = (psutil.Process(pid).name() or "")
            except Exception:
                proc_name = ""
            if not any(x in name for x in ("微信", "WeChat", "Weixin")):
                continue
            # 仅保留微信进程窗口，排除资源管理器等
            if proc_name.lower() not in ("weixin.exe", "wechat.exe"):
                continue
            # 常见主窗体类名（不同版本可能不同）
            score = 0
            if class_name in ("WeChatMainWndForPC", "WeChatMainWndForPC64", "WeChatMainWnd"):
                score += 5
            # 可见的大窗口更像主窗口
            area = _window_area(ei.rectangle)
            score += min(5, area // (800*600))  # 粗略加分
            # 登录窗口/小弹窗一般名字较短，面积较小
            candidates.append((score, area, w))
        except Exception:
            continue

    if not candidates:
        return None
    # 先按分数，再按面积排序，取最可能的
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    chosen = candidates[0][2]
    try:
        ei = chosen.element_info
        _log(f"选择窗口：title='{ei.name}' class='{ei.class_name}' area={_window_area(ei.rectangle)} pid={getattr(ei, 'process_id', '')}")
    except Exception:
        pass
    return chosen


def ensure_wechat_running(start_if_needed: bool = True, timeout: float = 20.0) -> None:
    # 降低 pywinauto 全局等待，避免卡住
    timings.Timings.window_find_timeout = 2
    timings.Timings.exists_timeout = 2
    timings.Timings.app_connect_timeout = 2
    """Ensure WeChat is running; optionally launch if not."""
    # 优先用枚举方式避免多窗口二义性
    win = _find_weixin_main_window()
    if win is not None:
        _log("已通过枚举找到主窗口")
        return

    if not start_if_needed:
        raise RuntimeError("WeChat is not running. Please start WeChat manually.")

    candidates = _possible_wechat_paths()
    fallbacks = ["Weixin.exe", "WeChat.exe"]
    tried_errors = []
    if not candidates:
        candidates = fallbacks

    started = False
    for exe in candidates:
        try:
            _log(f"尝试启动：{exe}")
            subprocess.Popen([exe], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            started = True
            break
        except FileNotFoundError as exc:
            tried_errors.append((exe, exc))
            continue
    if not started:
        raise RuntimeError(
            "找不到 Weixin/WeChat 可执行文件，请安装或调整路径。已尝试："
            + ", ".join(x for x, _ in tried_errors)
        )

    # Wait for window to appear
    def _connected():
        app = Application(backend="uia")
        # 直接优先用枚举选择主窗口并通过句柄附着
        chosen = _find_weixin_main_window()
        if chosen is not None:
            app.connect(handle=chosen.handle, timeout=2)
            _log("启动后已通过句柄连接")
            return True
        # 退回短超时标题/路径匹配
        try:
            app.connect(title_re="微信|WeChat|Weixin", timeout=2)
            _log("启动后已通过标题连接")
        except Exception:
            app.connect(path_re=r"Weixin\.exe|WeChat\.exe", timeout=2)
            _log("启动后已通过进程路径连接")
        return True

    wait_until_passes(timeout, 1.0, _connected)


def attach_wechat(timeout: float = 20.0):
    """Attach to WeChat main window and return (app, main_window)."""
    app = Application(backend=BACKEND)
    def _get_window():
        # 首选：通过枚举选择主窗口后用句柄附着
        chosen = _find_weixin_main_window()
        if chosen is not None:
            app.connect(handle=chosen.handle, timeout=2)
        else:
            # 退回：短超时标题或路径匹配
            try:
                app.connect(title_re="微信|WeChat|Weixin", timeout=2)
            except Exception:
                app.connect(path_re=r"Weixin\.exe|WeChat\.exe", timeout=2)
        win = app.top_window()
        # Ensure it's really WeChat/Weixin
        name = (win.element_info.name or "").lower()
        if not ("微信" in name or "wechat" in name or "weixin" in name):
            # Try to find by title
            win = app.window(title_re="微信|WeChat|Weixin")
        return win

    main_win = wait_until_passes(timeout, 1.0, _get_window)
    main_win.wait("ready", timeout=timeout)
    wrapper = main_win.wrapper_object()
    try:
        if getattr(wrapper, "is_minimized", None) and wrapper.is_minimized():
            _log("正在还原已最小化的窗口 ...")
            wrapper.restore()
    except Exception:
        pass
    _log("正在将焦点置于 Weixin 窗口 ...")
    wrapper.set_focus()
    try:
        if hasattr(wrapper, "set_keyboard_focus"):
            wrapper.set_keyboard_focus()
    except Exception:
        pass
    # Wait a moment to ensure foreground
    time.sleep(0.3)
    return app, main_win


def _try_focus_search_edit(main_win) -> bool:
    # 尝试直接聚焦看起来像“全局搜索”的输入框（Edit 控件）
    try:
        edits = main_win.descendants(control_type="Edit")
    except Exception:
        edits = []
    for edit in edits[:5]:
        name = (getattr(edit.element_info, "name", "") or "").lower()
        if ("search" in name) or ("搜索" in name) or ("查找" in name) or name == "":
            try:
                edit.set_focus()
                _log(f"已聚焦搜索框（Edit）：name='{name}'")
                return True
            except Exception:
                continue
    return False


def focus_search_and_open_chat(main_win, friend_name: str, delay: float = 0.25) -> None:
    """聚焦全局搜索（Ctrl+F / Ctrl+K / 直接控件），输入好友名并回车打开聊天。"""
    main_win.set_focus()
    time.sleep(0.2)
    # Try common shortcuts first
    for combo in ("^f", "^k", "^f"):
        _log(f"发送快捷键 {combo} 以聚焦搜索框 ...")
        keyboard.send_keys(combo)
        time.sleep(delay)
        keyboard.send_keys("^a{BACKSPACE}")
        time.sleep(delay)

    # If shortcuts didn't place focus into a text box, try direct Edit focus
    if not _try_focus_search_edit(main_win):
        _log("未能确认已聚焦搜索框，将直接向前台窗口输入。")

    # Now type and open the first result
    _log(f"输入好友名称：{friend_name}")
    # 再次清空，避免任何误输入字符
    keyboard.send_keys("^a{BACKSPACE}")
    time.sleep(0.1)
    keyboard.send_keys(friend_name, with_spaces=True)
    time.sleep(delay)
    _log("回车打开第一个搜索结果 ...")
    keyboard.send_keys("{ENTER}")
    time.sleep(delay)


def _focus_message_input(main_win) -> bool:
    """尝试聚焦聊天输入框。返回是否成功。"""
    try:
        ctrls = main_win.descendants()
    except Exception:
        ctrls = []

    scored = []
    try:
        rect_win = main_win.element_info.rectangle
        bottom_win = rect_win.bottom
    except Exception:
        bottom_win = 99999

    for c in ctrls:
        try:
            ei = c.element_info
            ct = getattr(ei, 'control_type', '') or ''
            cn = getattr(ei, 'class_name', '') or ''
            nm = getattr(ei, 'name', '') or ''
            rect = getattr(ei, 'rectangle', None)
            if ct not in ('Edit', 'Document', 'Text') and 'RichEdit' not in cn:
                continue
            if rect is None:
                continue
            area = _window_area(rect)
            # 越靠下越像输入框，给予更高分；面积也应该适中偏大
            distance_from_bottom = max(0, bottom_win - rect.bottom)
            score = area - distance_from_bottom * 10
            scored.append((score, rect, c, ct, cn, nm))
        except Exception:
            continue

    if not scored:
        _log("未找到候选输入控件")
        return False

    scored.sort(key=lambda x: x[0], reverse=True)
    for score, rect, ctrl, ct, cn, nm in scored[:5]:
        try:
            _log(f"尝试聚焦输入控件 type={ct} class={cn} name={nm}")
            try:
                ctrl.set_focus()
            except Exception:
                # 如果 set_focus 失败，尝试点击控件中心
                x = int((rect.left + rect.right) / 2)
                y = int((rect.top + rect.bottom) / 2)
                mouse.click(button='left', coords=(x, y))
            return True
        except Exception:
            continue
    _log("无法聚焦任一输入控件")
    return False


def _click_bottom_chat_area(main_win, clicks: int = 3) -> None:
    try:
        rect = main_win.element_info.rectangle
        cx = int((rect.left + rect.right) / 2)
        # 从底部往上多次点击，尽量点到输入区域
        for i in range(clicks):
            y = int(rect.bottom - 80 - i * 40)
            _log(f"尝试点击聊天窗口底部区域以获取焦点：({cx}, {y})")
            mouse.click(button='left', coords=(cx, y))
            time.sleep(0.1)
    except Exception:
        pass


def send_message_to_current_chat(main_win, message: str, delay: float = 0.12, press_enter_to_send: bool = True, use_paste: bool = False) -> None:
    """
    Type message into the current chat input and send.

    If your WeChat setting is "Enter to send", keep press_enter_to_send=True.
    If it's set to "Ctrl+Enter to send", set press_enter_to_send=False and it will use Ctrl+Enter.
    """
    # Ensure focus is in the message input area: usually Alt+s focuses input, but it's not universal.
    # Rely on default: once a chat is opened, input is focused. We still click End to ensure caret.
    # 多次尝试聚焦，避免聚焦后又被抢走
    attempts = 3
    for i in range(attempts):
        focused = _focus_message_input(main_win)
        if not focused:
            _click_bottom_chat_area(main_win, clicks=2)
            time.sleep(0.1)
            focused = _focus_message_input(main_win)
        time.sleep(0.1)
        _log("确保光标在输入框末尾 ...")
        keyboard.send_keys("{END}")
        time.sleep(0.08)
        # 在真正输入前再次尝试设置键盘焦点
        try:
            w = main_win.wrapper_object()
            if hasattr(w, "set_keyboard_focus"):
                w.set_keyboard_focus()
        except Exception:
            pass

        # 输入消息（可选用粘贴，降低 IME 干扰概率）
        _log(f"输入消息：{message}")
        if use_paste and pyperclip is not None:
            try:
                pyperclip.copy(message)
                keyboard.send_keys("^v")
            except Exception:
                keyboard.send_keys(message, with_spaces=True)
        else:
            keyboard.send_keys(message, with_spaces=True)
        time.sleep(delay)
        # 若输入后可能又失焦，外层循环会再次尝试
        break
    if press_enter_to_send:
        _log("使用 Enter 发送 ...")
        keyboard.send_keys("{ENTER}")
    else:
        _log("使用 Ctrl+Enter 发送 ...")
        keyboard.send_keys("^{{ENTER}}")
    time.sleep(delay)


def send_messages_to_friends(
    friends: List[str],
    messages: List[str],
    start_if_needed: bool = True,
    per_friend_pause: float = 0.5,
    per_message_pause: float = 0.2,
    press_enter_to_send: bool = True,
) -> None:
    _log("确保 Weixin/WeChat 已启动 ...")
    ensure_wechat_running(start_if_needed=start_if_needed)
    _log("正在附着到窗口 ...")
    _, main_win = attach_wechat()

    for friend in friends:
        _log(f"打开与 {friend} 的聊天 ...")
        focus_search_and_open_chat(main_win, friend)
        time.sleep(per_friend_pause)
        for msg in messages:
            send_message_to_current_chat(
                main_win,
                msg,
                delay=per_message_pause,
                press_enter_to_send=press_enter_to_send,
            )


def _parse_cli_args(argv: List[str]):
    import argparse

    parser = argparse.ArgumentParser(description="使用 pywinauto 向微信/Weixin 发送消息")
    parser.add_argument(
        "--friends",
        type=str,
        help="以英文逗号分隔的好友名称，例如：张三,李四",
    )
    parser.add_argument(
        "--messages",
        type=str,
        help="以分号分隔的多条消息，例如：早上好;开会见",
    )
    parser.add_argument(
        "--no-launch",
        action="store_true",
        help="若未运行则不要自动启动微信/Weixin",
    )
    parser.add_argument(
        "--ctrl-enter",
        action="store_true",
        help="使用 Ctrl+Enter 发送（当微信设置为该模式时）",
    )
    parser.add_argument(
        "--friend-delay",
        type=float,
        default=0.5,
        help="切换到好友聊天后等待的秒数",
    )
    parser.add_argument(
        "--message-delay",
        type=float,
        default=0.2,
        help="每条消息输入/发送时的等待秒数",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="开启详细日志（中文）",
    )
    parser.add_argument(
        "--dump-controls",
        action="store_true",
        help="仅导出前若干个控件信息到控制台（用于诊断）",
    )
    parser.add_argument(
        "--backend",
        type=str,
        choices=["uia", "win32"],
        default="uia",
        help="选择后端（默认 uia，可尝试 win32 兜底）",
    )

    args = parser.parse_args(argv)
    friends = []
    messages = []
    if args.friends:
        friends = [s.strip() for s in args.friends.split(",") if s.strip()]
    if args.messages:
        messages = [s.strip() for s in args.messages.split(";") if s.strip()]
    return {
        "friends": friends,
        "messages": messages,
        "start_if_needed": not args.no_launch,
        "press_enter_to_send": not args.ctrl_enter,
        "per_friend_pause": args.friend_delay,
        "per_message_pause": args.message_delay,
        "verbose": args.verbose,
        "dump_controls": args.dump_controls,
        "backend": args.backend,
    }


def _dump_some_controls(main_win, limit: int = 50) -> None:
    try:
        ctrls = main_win.descendants()
    except Exception:
        _log("枚举子控件失败。")
        return
    print("-- 控件导出开始 --")
    count = 0
    for c in ctrls:
        try:
            ei = c.element_info
            print(f"[{count}] 控件类型={getattr(ei, 'control_type', '')} 名称={getattr(ei, 'name', '')} 类名={getattr(ei, 'class_name', '')}")
        except Exception:
            pass
        count += 1
        if count >= limit:
            break
    print("-- 控件导出结束 --")


def main(argv: Optional[List[str]] = None) -> None:
    if argv is None:
        argv = sys.argv[1:]
    cfg = _parse_cli_args(argv)
    global VERBOSE
    VERBOSE = cfg.get("verbose", False)
    global BACKEND
    BACKEND = cfg.get("backend", "uia")
    _ensure_utf8_console()

    # Defaults for quick demo if nothing is passed
    if not cfg["friends"]:
        cfg["friends"] = ["文件传输助手"]
    if not cfg["messages"]:
        cfg["messages"] = ["这是一条来自 pywinauto 的自动消息。"]

    if cfg.get("dump_controls"):
        ensure_wechat_running(start_if_needed=True)
        _, main_win = attach_wechat()
        _dump_some_controls(main_win)
        return

    send_messages_to_friends(
        friends=cfg["friends"],
        messages=cfg["messages"],
        start_if_needed=cfg["start_if_needed"],
        per_friend_pause=cfg["per_friend_pause"],
        per_message_pause=cfg["per_message_pause"],
        press_enter_to_send=cfg["press_enter_to_send"],
    )


if __name__ == "__main__":
    main()
