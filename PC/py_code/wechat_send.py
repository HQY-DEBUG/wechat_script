"""
wechat_send.py  --  微信连续发送消息脚本
版本    : v1.3
日期    : 2026/04/14

修改记录:
    v1.3  自动聚焦微信窗口；支持 --to 指定联系人；emoji 通过剪贴板直接支持
    v1.2  粘贴后等待延长至 0.3s，Enter 后增加 0.3s 等待，确保每条消息都能发出
    v1.1  interval 改为可选位置参数，支持 py wechat_send.py <msg> <count> [interval]
"""

import sys
import time
import argparse
import pyperclip
import pyautogui
import pygetwindow


def focus_wechat():
    """
    @brief  查找并激活微信主窗口
    @return 成功返回 True，未找到窗口则打印错误并退出
    """
    windows = pygetwindow.getWindowsWithTitle('微信')
    if not windows:
        print("错误：未找到微信窗口，请先打开微信")
        sys.exit(1)
    win = windows[0]
    win.activate()
    time.sleep(0.5)  # 等待窗口激活完成


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
    @param  message   要发送的消息内容（支持 emoji，如 😄）
    @param  count     发送次数
    @param  interval  每次发送间隔（秒）
    """
    for i in range(1, count + 1):
        # 将消息写入剪贴板，避免中文输入法和 emoji 输入问题
        pyperclip.copy(message)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)   # 等待微信完成粘贴
        pyautogui.press('enter')
        time.sleep(0.3)   # 等待微信完成发送，避免下次粘贴时输入框未就绪
        print(f"[{i}/{count}] 已发送：{message}")
        if i < count:
            time.sleep(interval)

    print("发送完成。")


def main():
    parser = argparse.ArgumentParser(
        description="微信连续发送消息脚本。消息内容支持 emoji，如：py wechat_send.py 😄 3"
    )
    parser.add_argument("message", help="要发送的消息内容（支持 emoji）")
    parser.add_argument("count", type=int, help="发送次数")
    parser.add_argument("interval", type=float, nargs="?", default=1.0, help="每次发送间隔（秒），默认 1.0")
    parser.add_argument("--to", metavar="联系人", default=None, help="联系人名称（昵称/备注），不填则发送到当前打开的对话")
    args = parser.parse_args()

    if args.count <= 0:
        print("错误：发送次数必须大于 0")
        return

    focus_wechat()

    if args.to:
        print(f"正在搜索联系人：{args.to}")
        open_contact(args.to)

    send_wechat_message(args.message, args.count, args.interval)


if __name__ == "__main__":
    main()
