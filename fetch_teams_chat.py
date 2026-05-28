#!/usr/bin/env python3
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv
from msal import PublicClientApplication


GRAPH_BASE = "https://graph.microsoft.com/v1.0"


@dataclass
class Config:
    client_id: str
    tenant_id: str
    scopes: list[str]
    output_dir: str
    login_port: int


def ensure_output_dir(path_str: str) -> Path:
    path = Path(path_str)
    path.mkdir(parents=True, exist_ok=True)
    return path


def save_json(path: Path, data: Any) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def pretty_http_error(resp: requests.Response, label: str) -> None:
    print(f"[ERROR] {label} failed")
    print(f"Status code: {resp.status_code}")
    print("Response text:")
    print(resp.text)


def sanitize_chat_id(chat_id: str) -> str:
    safe_chars = []
    for ch in chat_id:
        if ch.isalnum() or ch in {"-", "_"}:
            safe_chars.append(ch)
        else:
            safe_chars.append("_")
    return "".join(safe_chars)


def load_config() -> Config:
    load_dotenv()

    client_id = os.getenv("CLIENT_ID")
    tenant_id = os.getenv("TENANT_ID")
    scopes_raw = os.getenv("SCOPES", "User.Read Chat.Read")
    output_dir = os.getenv("OUTPUT_DIR", "output")
    login_port_raw = os.getenv("MSAL_PORT", "53100")

    if not client_id or not tenant_id:
        raise ValueError("CLIENT_ID 或 TENANT_ID 未在 .env 設定")

    scopes = [s for s in scopes_raw.split() if s]
    if not scopes:
        raise ValueError("SCOPES 不可為空，至少需要 User.Read 與 Chat.Read")

    try:
        login_port = int(login_port_raw)
    except ValueError as exc:
        raise ValueError("MSAL_PORT 必須是整數") from exc

    return Config(
        client_id=client_id,
        tenant_id=tenant_id,
        scopes=scopes,
        output_dir=output_dir,
        login_port=login_port,
    )


def get_token(client_id: str, tenant_id: str, scopes: list[str], port: int) -> str:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = PublicClientApplication(
        client_id=client_id,
        authority=authority,
    )

    try:
        result = app.acquire_token_interactive(
            scopes=scopes,
            port=port,
        )
    except Exception as interactive_error:
        # 備援：互動式登入失敗時改走 device flow。
        print("[WARN] acquire_token_interactive 失敗，改用 device flow。")
        print(f"[WARN] {interactive_error}")
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(
                f"Device flow 初始化失敗: {json.dumps(flow, ensure_ascii=False, indent=2)}"
            )
        print(flow.get("message", "請依照終端機提示完成裝置登入。"))
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(
            f"Token acquisition failed: {json.dumps(result, ensure_ascii=False, indent=2)}"
        )
    return result["access_token"]


def graph_get(url: str, access_token: str) -> requests.Response:
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }
    try:
        return requests.get(url, headers=headers, timeout=30)
    except requests.RequestException as exc:
        raise RuntimeError(f"Request failed for {url}: {exc}") from exc


def main() -> int:
    try:
        cfg = load_config()
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return 1

    out = ensure_output_dir(cfg.output_dir)

    print("[INFO] Acquiring token...")
    try:
        token = get_token(cfg.client_id, cfg.tenant_id, cfg.scopes, cfg.login_port)
    except RuntimeError as exc:
        print("[ERROR] token 取得失敗")
        print(str(exc))
        return 1
    print("登入成功")

    # 1) GET /me
    me_url = f"{GRAPH_BASE}/me"
    try:
        me_resp = graph_get(me_url, token)
    except RuntimeError as exc:
        print(str(exc))
        return 1

    if me_resp.status_code != 200:
        pretty_http_error(me_resp, "GET /me")
        return 1

    me_data = me_resp.json()
    save_json(out / "me.json", me_data)
    user_name = me_data.get("displayName") or me_data.get("userPrincipalName") or "(unknown)"
    print(f"使用者名稱: {user_name}")

    # 2) GET /me/chats?$top=10
    chats_url = f"{GRAPH_BASE}/me/chats?$top=10"
    try:
        chats_resp = graph_get(chats_url, token)
    except RuntimeError as exc:
        print(str(exc))
        return 1

    if chats_resp.status_code != 200:
        pretty_http_error(chats_resp, "GET /me/chats")
        if chats_resp.status_code == 403:
            print("[HINT] 可能原因：")
            print(" - Chat.Read consent 尚未完成")
            print(" - 租戶政策阻止 user consent")
            print(" - App 未正確設為 public client (Allow public client flows = Yes)")
        return 1

    chats_data = chats_resp.json()
    save_json(out / "chats.json", chats_data)

    chats = chats_data.get("value", [])
    print(f"chat 數量: {len(chats)}")

    if not chats:
        print("[HINT] value 為空，請先到 Teams 建立測試 chat。")
        return 0

    first_chat = chats[0]
    chat_id = first_chat.get("id")
    if not chat_id:
        print("[ERROR] 第一個 chat 缺少 id 欄位。")
        return 1
    print(f"第一個 chat id: {chat_id}")

    # 3) GET /chats/{chat-id}/messages?$top=10
    messages_url = f"{GRAPH_BASE}/chats/{chat_id}/messages?$top=10"
    try:
        msg_resp = graph_get(messages_url, token)
    except RuntimeError as exc:
        print(str(exc))
        return 1

    if msg_resp.status_code != 200:
        pretty_http_error(msg_resp, f"GET /chats/{chat_id}/messages")
        print("messages 抓取成功: 否")
        return 1

    msg_data = msg_resp.json()
    safe_chat_id = sanitize_chat_id(chat_id)
    save_json(out / f"messages_{safe_chat_id}.json", msg_data)

    next_link = msg_data.get("@odata.nextLink")
    msg_count = len(msg_data.get("value", []))
    print(f"messages 抓取成功: 是 (count={msg_count})")
    print(f"是否有 @odata.nextLink: {'有' if next_link else '無'}")
    if next_link:
        print(next_link)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
