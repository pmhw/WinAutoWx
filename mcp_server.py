import asyncio
import json
import os
from typing import List

from fastmcp import FastMCP
import httpx

API_URL = os.environ.get("WEIXIN_API_URL", "http://127.0.0.1:8000")

app = FastMCP(
    name="weixin-auto-sender",
    version="2.0.0",
)

@app.tool()
async def send_messages(
    friends: List[str],
    messages: List[str],
    backend: str = "win32",
    ctrl_enter: bool = False,
    friend_delay: float = 0.5,
    message_delay: float = 0.2,
    no_launch: bool = False,
    verbose: bool = False,
) -> dict:
    """向好友/群聊发送消息（通过本地 HTTP 自动化服务）。

    参数：
    - friends: 好友或群聊名称列表
    - messages: 要发送的消息列表
    - backend: 自动化后端，'uia' 或 'win32'
    - ctrl_enter: 是否使用 Ctrl+Enter 发送
    - friend_delay: 切换到聊天后的等待秒数
    - message_delay: 每条消息输入/发送时的等待秒数
    - no_launch: 若未运行则不要自动启动 Weixin
    - verbose: 是否输出详细日志

    返回：JSON 结果
    """
    async with httpx.AsyncClient(timeout=60.0) as client:
        resp = await client.post(
            f"{API_URL}/send",
            json={
                "friends": friends,
                "messages": messages,
                "backend": backend,
                "ctrl_enter": ctrl_enter,
                "friend_delay": friend_delay,
                "message_delay": message_delay,
                "no_launch": no_launch,
                "verbose": verbose,
            },
        )
        resp.raise_for_status()
        return resp.json()

@app.tool()
async def dump_controls(
    backend: str = "win32",
    verbose: bool = True,
) -> dict:
    """导出 Weixin 主窗口的前若干个控件信息（通过本地 HTTP 自动化服务）。"""
    # proxies=None, trust_env=False,
    async with httpx.AsyncClient(timeout=60.0, proxies=None, trust_env=False) as client:
        resp = await client.post(
            f"{API_URL}/dump",
            json={
                "backend": backend,
                "verbose": verbose,
            },
        )
        resp.raise_for_status()
        return resp.json()

if __name__ == "__main__":
    # 以 STDIO 方式运行 MCP 服务器
    app.run()
