"""
wechat_send.py  --  微信连续发送消息脚本
版本    : v2.0
日期    : 2026/04/14

修改记录:
    v2.0  重写窗口查找逻辑：枚举所有窗口，按 QWindowIcon 类+Weixin 标题定位主窗口
    v1.5  发送前点击输入框获取焦点，修复消息未发出的问题
    v1.4  改用 win32gui 激活微信窗口，修复窗口标题应为 Weixin 的问题
"""

import sys
import time
import argparse
import ctypes
import pyperclip
import pyautogui
import win32gui  # type: ignore
import win32con  # type: ignore


def get_wechat_hwnd() -> int:
    """
    @brief  枚举所有窗口，找标题为 Weixin、类名含 QWindowIcon 的最大窗口（微信主窗口）
    @return 微信主窗口句柄
    """
    candidates = []

    def cb(hwnd, _):
        try:
            if win32gui.GetWindowText(hwnd) != 'Weixin':
                return
            if 'QWindowIcon' not in (win32gui.GetClassName(hwnd) or ''):
                return
            r = win32gui.GetWindowRect(hwnd)
            w, h = r[2] - r[0], r[3] - r[1]
            candidates.append((w * h, hwnd))
        except Exception:
            pass

    win32gui.EnumWindows(cb, None)
    if not candidates:
        print("错误：未找到微信主窗口，请先打开微信")
        sys.exit(1)
    candidates.sort(reverse=True)
    return candidates[0][1]


def activate_wechat(hwnd: int):
    """
    @brief  显示并激活微信主窗口
    @param  hwnd  微信主窗口句柄
    """
    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
    ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)       # Alt 按下（绕过焦点限制）
    win32gui.SetForegroundWindow(hwnd)
    ctypes.windll.user32.keybd_event(0x12, 0, 0x0002, 0)  # Alt 释放
    time.sleep(0.5)


def click_input_box(hwnd: int):
    """
    @brief  点击微信聊天输入框，确保后续粘贴落入输入区域
    @param  hwnd  微信主窗口句柄
    """
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

    hwnd = get_wechat_hwnd()
    activate_wechat(hwnd)

    if args.to:
        print(f"正在搜索联系人：{args.to}")
        open_contact(args.to)

    click_input_box(hwnd)
    send_wechat_message(args.message, args.count, args.interval)


if __name__ == "__main__":
    main()
