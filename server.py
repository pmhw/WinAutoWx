from fastapi import FastAPI
from pydantic import BaseModel, Field
from typing import List, Optional

# Import the script module without changing it
from script import wechat_sender as ws

app = FastAPI(title="Weixin Auto Sender API", version="2.0")

class SendRequest(BaseModel):
    friends: List[str] = Field(..., description="好友/群聊名称列表")
    messages: List[str] = Field(..., description="要发送的消息列表")
    backend: str = Field("uia", description="后端：uia 或 win32")
    ctrl_enter: bool = Field(False, description="是否使用 Ctrl+Enter 发送")
    friend_delay: float = Field(0.5, description="切换聊天后的等待秒数")
    message_delay: float = Field(0.2, description="每条消息的等待秒数")
    no_launch: bool = Field(False, description="不自动启动微信/Weixin")
    verbose: bool = Field(False, description="中文详细日志")

class DumpRequest(BaseModel):
    backend: str = Field("uia", description="后端：uia 或 win32")
    verbose: bool = Field(True, description="中文详细日志")

@app.post("/send")
async def send_messages(req: SendRequest):
    # Configure globals as the script's CLI would
    ws.BACKEND = req.backend
    ws.VERBOSE = req.verbose

    # Ensure running and attach
    ws.ensure_wechat_running(start_if_needed=not req.no_launch)
    _, main_win = ws.attach_wechat()

    # Send
    for friend in req.friends:
        ws.focus_search_and_open_chat(main_win, friend)
        ws.time.sleep(req.friend_delay)
        for msg in req.messages:
            ws.send_message_to_current_chat(
                main_win,
                msg,
                delay=req.message_delay,
                press_enter_to_send=(not req.ctrl_enter),
            )
    return {"ok": True}

@app.post("/dump")
async def dump_controls(req: DumpRequest):
    ws.BACKEND = req.backend
    ws.VERBOSE = req.verbose
    ws.ensure_wechat_running(start_if_needed=True)
    _, main_win = ws.attach_wechat()

    # Collect a small dump
    out = []
    try:
        ctrls = main_win.descendants()
    except Exception:
        return {"ok": False, "error": "enumerate_failed"}

    count = 0
    for c in ctrls:
        try:
            ei = c.element_info
            out.append({
                "type": getattr(ei, 'control_type', ''),
                "name": getattr(ei, 'name', ''),
                "class": getattr(ei, 'class_name', ''),
            })
        except Exception:
            pass
        count += 1
        if count >= 80:
            break

    return {"ok": True, "controls": out}

# Run with: uvicorn server:app --host 127.0.0.1 --port 8000
