"""
wechat_send.py  --  微信连续发送消息脚本
版本    : v2.5
日期    : 2026/04/14

修改记录:
    v2.5  新增 focus_wechat()，click_input_box 前强制置顶微信，修复 Ctrl+V 打到终端的问题
    v2.4  改用向托盘消息窗口发送双击消息唤出微信，彻底避免白屏和登录界面问题
    v2.3  改用重启 weixin.exe 唤出窗口（已废弃，会弹出登录界面）
"""

import sys
import time
import argparse
import ctypes
import psutil          # type: ignore
import pyperclip
import pyautogui
import win32gui        # type: ignore
import win32con        # type: ignore
import win32process    # type: ignore


def activate_wechat():
    """
    @brief  向微信托盘消息窗口发送双击消息，让微信自己唤出主窗口
    @note   类名以 WxTrayIconMessageWindowClass 结尾，Qt 版本号前缀可能随版本变化
    """
    tray_hwnd = 0

    def cb(hwnd, _):
        nonlocal tray_hwnd
        if (win32gui.GetClassName(hwnd) or '').endswith('WxTrayIconMessageWindowClass'):
            tray_hwnd = hwnd

    win32gui.EnumWindows(cb, None)
    if not tray_hwnd:
        print("错误：未找到微信托盘窗口，请先打开微信")
        sys.exit(1)

    # Qt QSystemTrayIcon 托盘消息 = WM_APP+1，图标ID=0，lParam=双击
    win32gui.PostMessage(tray_hwnd, win32con.WM_APP + 1, 0, win32con.WM_LBUTTONDBLCLK)
    time.sleep(1.5)   # 等待微信主窗口完全显示和渲染


def get_wechat_hwnd() -> int:
    """
    @brief  获取微信主窗口句柄（QWindowIcon 类，尺寸 ≥400×300）
    @return 微信主窗口句柄
    """
    weixin_pids = {p.pid for p in psutil.process_iter(['pid', 'name'])
                   if p.info['name'].lower() == 'weixin.exe'}
    if not weixin_pids:
        print("错误：未找到微信进程，请先打开微信")
        sys.exit(1)

    candidates = []

    def cb(hwnd, _):
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid not in weixin_pids:
                return
            if 'QWindowIcon' not in (win32gui.GetClassName(hwnd) or ''):
                return
            r = win32gui.GetWindowRect(hwnd)
            w, h = r[2] - r[0], r[3] - r[1]
            if w >= 400 and h >= 300:
                candidates.append((w * h, hwnd))
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    if not candidates:
        print("错误：未找到微信主窗口")
        sys.exit(1)
    candidates.sort(reverse=True)
    return candidates[0][1]


def focus_wechat(hwnd: int):
    """
    @brief  强制将微信窗口置于前台，绕过 Windows 焦点限制
    @param  hwnd  微信主窗口句柄
    """
    ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)       # Alt 按下
    win32gui.SetForegroundWindow(hwnd)
    ctypes.windll.user32.keybd_event(0x12, 0, 0x0002, 0)  # Alt 释放
    time.sleep(0.3)


def click_input_box(hwnd: int):
    """
    @brief  将微信置于前台并点击聊天输入框
    @param  hwnd  微信主窗口句柄
    """
    focus_wechat(hwnd)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    x = left + int((right - left) * 0.65)   # 右侧聊天区域水平中点
    y = top  + int((bottom - top) * 0.88)   # 底部输入框位置
    pyautogui.click(x, y)
    time.sleep(0.2)


def open_contact(name: str):
    """
    @brief  在微信搜索框中搜索联系人并打开对话窗口
    @param  name  联系人名称（昵称 / 备注）
    """
    pyautogui.hotkey('ctrl', 'f')
    time.sleep(0.5)
    pyperclip.copy(name)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1.0)   # 等待搜索结果出现
    pyautogui.press('enter')
    time.sleep(0.8)   # 等待对话窗口打开


def send_wechat_message(message: str, count: int, interval: float):
    """
    @brief  连续向微信当前聊天窗口发送消息
    @param  message   要发送的消息内容（支持 emoji）
    @param  count     发送次数
    @param  interval  每次发送间隔（秒）
    """
    for i in range(1, count + 1):
        pyperclip.copy(message)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)   # 等待微信完成粘贴
        pyautogui.press('enter')
        time.sleep(0.3)   # 等待微信完成发送
        print(f"[{i}/{count}] 已发送：{message}")
        if i < count:
            time.sleep(interval)
    print("发送完成。")


def main():
    parser = argparse.ArgumentParser(
        description="微信连续发送消息脚本（支持 emoji，如：py wechat_send.py 😄 3）"
    )
    parser.add_argument("message", help="要发送的消息内容（支持 emoji）")
    parser.add_argument("count", type=int, help="发送次数")
    parser.add_argument("interval", type=float, nargs="?", default=1.0, help="发送间隔（秒），默认 1.0")
    parser.add_argument("-t", "-to", dest="to", metavar="联系人", default=None,
                        help="联系人名称（昵称/备注），不填则发到当前对话")
    args = parser.parse_args()

    if args.count <= 0:
        print("错误：发送次数必须大于 0")
        return

    activate_wechat()           # 通过托盘双击消息唤出微信
    hwnd = get_wechat_hwnd()    # 获取已显示的主窗口句柄

    if args.to:
        print(f"正在搜索联系人：{args.to}")
        open_contact(args.to)

    click_input_box(hwnd)
    send_wechat_message(args.message, args.count, args.interval)


if __name__ == "__main__":
    main()
