"""
wechat_send.py  --  微信连续发送消息脚本
版本    : v1.2
日期    : 2026/04/14

修改记录:
    v1.2  粘贴后等待延长至 0.3s，Enter 后增加 0.3s 等待，确保每条消息都能发出
    v1.1  interval 改为可选位置参数，支持 py wechat_send.py <msg> <count> [interval]
    v1.0  创建文件，实现微信连续发送消息功能
"""

import time
import argparse
import pyperclip
import pyautogui


def send_wechat_message(message: str, count: int, interval: float):
    """
    连续向微信当前聊天窗口发送消息。

    Args:
        message:  要发送的消息内容
        count:    发送次数
        interval: 每次发送间隔（秒）
    """
    print(f"请在 3 秒内切换到微信聊天窗口...")
    time.sleep(3)

    for i in range(1, count + 1):
        # 将消息写入剪贴板，避免中文输入法问题
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
    parser = argparse.ArgumentParser(description="微信连续发送消息脚本")
    parser.add_argument("message", help="要发送的消息内容")
    parser.add_argument("count", type=int, help="发送次数")
    parser.add_argument("interval", type=float, nargs="?", default=1.0, help="每次发送间隔（秒），默认 1.0")
    args = parser.parse_args()

    if args.count <= 0:
        print("错误：发送次数必须大于 0")
        return

    send_wechat_message(args.message, args.count, args.interval)


if __name__ == "__main__":
    main()
