#!/usr/bin/env python3
"""TicNote Studio 本地服务：静态页面 + TicNote API 代理（绕过浏览器 CORS）。"""

from __future__ import annotations

import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from typing import Any

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)
from common import resolve_base_url  # noqa: E402

PORT = int(os.environ.get("PORT", "8765"))
HOST = os.environ.get("HOST", "127.0.0.1")


def _read_json(handler: SimpleHTTPRequestHandler) -> dict[str, Any]:
    length = int(handler.headers.get("Content-Length") or 0)
    if length <= 0:
        return {}
    raw = handler.rfile.read(length)
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def _send_json(handler: SimpleHTTPRequestHandler, payload: Any, status: int = 200) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    handler.wfile.write(body)


def _api_request(
    method: str,
    url: str,
    token: str | None = None,
    payload: Any | None = None,
    timeout: int = 60,
) -> dict[str, Any]:
    data = None
    headers = {"Accept": "application/json"}
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            text = resp.read().decode("utf-8", errors="replace")
            try:
                body = json.loads(text) if text else {}
            except json.JSONDecodeError:
                body = {"raw": text}
            return {"ok": True, "status": resp.status, "data": body}
    except urllib.error.HTTPError as e:
        err_text = e.read().decode("utf-8", errors="replace")
        try:
            err_body = json.loads(err_text) if err_text else {}
        except json.JSONDecodeError:
            err_body = {"raw": err_text}
        return {"ok": False, "status": e.code, "error": err_body, "message": str(e)}
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "status": 0, "error": str(e), "message": str(e)}


def _extract_transcript(detail: dict[str, Any]) -> str:
    """从 file-detail 响应中尽量抽出可读正文。"""
    for key in ("transcribeJson", "summaryJson"):
        raw = detail.get(key)
        if not raw:
            continue
        if isinstance(raw, dict):
            obj = raw
        else:
            try:
                obj = json.loads(raw)
            except (TypeError, json.JSONDecodeError):
                if isinstance(raw, str) and raw.strip():
                    return raw.strip()
                continue
        text = _walk_text(obj)
        if text:
            return text
    for key in ("content", "text", "preview", "title"):
        val = detail.get(key)
        if isinstance(val, str) and val.strip():
            return val.strip()
    return ""


def _walk_text(obj: Any) -> str:
    chunks: list[str] = []

    def walk(node: Any) -> None:
        if isinstance(node, str):
            s = node.strip()
            if s and len(s) > 1:
                chunks.append(s)
            return
        if isinstance(node, dict):
            # 优先常见字段
            for k in ("text", "content", "sentence", "transcript", "summary", "markdown", "value"):
                if k in node:
                    walk(node[k])
            for k, v in node.items():
                if k in ("text", "content", "sentence", "transcript", "summary", "markdown", "value"):
                    continue
                if k.lower() in ("id", "ids", "time", "start", "end", "offset", "speakerid"):
                    continue
                walk(v)
            return
        if isinstance(node, list):
            for item in node:
                walk(item)

    walk(obj)
    # 去重保序
    seen = set()
    out = []
    for c in chunks:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return "\n".join(out)


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("[studio] " + (fmt % args) + "\n")

    def do_OPTIONS(self) -> None:  # noqa: N802
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.end_headers()

    def do_GET(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/health":
            return _send_json(self, {"ok": True, "service": "ticnote-studio", "ts": int(time.time())})
        if parsed.path == "/":
            self.path = "/index.html"
        return super().do_GET()

    def do_POST(self) -> None:  # noqa: N802
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        try:
            body = _read_json(self)
        except json.JSONDecodeError:
            return _send_json(self, {"ok": False, "error": "无效 JSON"}, 400)

        if path == "/api/ticnote/login":
            return self._login(body)
        if path == "/api/ticnote/projects":
            return self._projects(body)
        if path == "/api/ticnote/files":
            return self._files(body)
        if path == "/api/ticnote/file-detail":
            return self._file_detail(body)
        if path == "/api/ticnote/transcribe":
            return self._transcribe(body)
        if path == "/api/ticnote/translate":
            return self._translate(body)
        if path == "/api/ticnote/deep-research":
            return self._deep_research(body)
        if path == "/api/ticnote/sync-project":
            return self._sync_project(body)
        if path == "/api/fetch-url":
            return self._fetch_url(body)
        if path == "/api/ai/chat":
            return self._ai_chat(body)

        return _send_json(self, {"ok": False, "error": f"未知接口: {path}"}, 404)

    def _login(self, body: dict[str, Any]) -> None:
        appkey = (body.get("appkey") or "").strip()
        if not appkey:
            return _send_json(self, {"ok": False, "error": "请填写 TicNote AppKey"}, 400)
        try:
            base_url = resolve_base_url(appkey)
        except ValueError as e:
            return _send_json(self, {"ok": False, "error": str(e)}, 400)
        result = _api_request("POST", f"{base_url}/api/p1/appkey/login", payload={"appkey": appkey})
        if not result["ok"]:
            return _send_json(
                self,
                {
                    "ok": False,
                    "error": result.get("error") or result.get("message"),
                    "hint": "请到 https://www.ticnote.cn 个人中心「TicNote Key」获取 AppKey",
                },
                400,
            )
        data = result["data"]
        token = None
        if isinstance(data, dict):
            token = (data.get("data") or {}).get("token") or data.get("token")
        if not token:
            return _send_json(self, {"ok": False, "error": "登录成功但未返回 token", "raw": data}, 502)
        return _send_json(
            self,
            {
                "ok": True,
                "token": token,
                "base_url": base_url,
                "raw": data,
            },
        )

    def _require_auth(self, body: dict[str, Any]) -> tuple[str, str] | None:
        appkey = (body.get("appkey") or "").strip()
        token = (body.get("token") or "").strip()
        if not appkey or not token:
            _send_json(self, {"ok": False, "error": "缺少 appkey 或 token"}, 401)
            return None
        try:
            base_url = resolve_base_url(appkey)
        except ValueError as e:
            _send_json(self, {"ok": False, "error": str(e)}, 400)
            return None
        return base_url, token

    def _projects(self, body: dict[str, Any]) -> None:
        auth = self._require_auth(body)
        if not auth:
            return
        base_url, token = auth
        query = (body.get("query") or "").strip()
        url = f"{base_url}/api/v2/file-index/chats"
        if query:
            url += "?" + urllib.parse.urlencode({"query": query})
        result = _api_request("GET", url, token=token)
        if not result["ok"]:
            return _send_json(self, {"ok": False, "error": result.get("error")}, 400)
        data = result["data"]
        chats = data.get("chats") if isinstance(data, dict) else data
        projects = []
        for chat in chats or []:
            info = chat.get("projectInfo") or {}
            projects.append(
                {
                    "chatId": chat.get("id"),
                    "name": chat.get("name") or info.get("name") or "未命名项目",
                    "projectId": chat.get("project_id") or info.get("id"),
                    "fileNum": info.get("fileNum", 0),
                    "hasAgent": chat.get("has_agent", False),
                    "agentCount": chat.get("agent_count", 0),
                    "updatedAt": chat.get("updatedAt") or chat.get("lastMessageAt"),
                    "icon": info.get("icon") or "📁",
                    "color": info.get("color"),
                }
            )
        return _send_json(self, {"ok": True, "projects": projects})

    def _files(self, body: dict[str, Any]) -> None:
        auth = self._require_auth(body)
        if not auth:
            return
        base_url, token = auth
        root_id = body.get("projectId") or body.get("rootId") or body.get("chatId")
        if not root_id:
            return _send_json(self, {"ok": False, "error": "请提供 projectId"}, 400)
        url = f"{base_url}/api/v1/file-index/file-tree?" + urllib.parse.urlencode(
            {"rootId": str(root_id)}
        )
        result = _api_request("GET", url, token=token)
        if not result["ok"]:
            return _send_json(self, {"ok": False, "error": result.get("error")}, 400)
        data = result["data"]
        tree = data.get("fileTree") if isinstance(data, dict) else data
        flat = self._flatten(tree or [])
        return _send_json(self, {"ok": True, "files": flat, "raw": data})

    def _flatten(self, nodes: list, depth: int = 0) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        for node in nodes:
            remark = {}
            raw_remark = node.get("subRemark")
            if isinstance(raw_remark, str) and raw_remark:
                try:
                    remark = json.loads(raw_remark)
                except json.JSONDecodeError:
                    remark = {}
            elif isinstance(raw_remark, dict):
                remark = raw_remark
            item = {
                "id": node.get("id"),
                "fileId": node.get("fileId"),
                "name": node.get("name"),
                "type": node.get("type"),
                "fileType": node.get("fileType"),
                "path": node.get("path"),
                "depth": depth,
                "transcodeStatus": remark.get("transcodeStatus"),
                "transcribeId": remark.get("transcribeId"),
                "summaryId": remark.get("summaryId"),
                "deepResearchStatus": remark.get("deepResearchStatus"),
            }
            out.append(item)
            children = node.get("children") or []
            if children:
                out.extend(self._flatten(children, depth + 1))
        return out

    def _file_detail(self, body: dict[str, Any]) -> None:
        auth = self._require_auth(body)
        if not auth:
            return
        base_url, token = auth
        record_id = body.get("recordId")
        if not record_id:
            return _send_json(self, {"ok": False, "error": "请提供 recordId"}, 400)
        poll = bool(body.get("poll"))
        interval = int(body.get("interval") or 5)
        timeout = int(body.get("timeout") or 180)
        start = time.time()
        last = None
        while True:
            result = _api_request(
                "GET",
                f"{base_url}/api/v2/file-index/file-detail/{record_id}",
                token=token,
            )
            if not result["ok"]:
                return _send_json(self, {"ok": False, "error": result.get("error")}, 400)
            data = result["data"]
            detail = data.get("data") if isinstance(data, dict) and "data" in data else data
            last = detail if isinstance(detail, dict) else {"raw": detail}
            text = _extract_transcript(last)
            status = last.get("status")
            transcode = last.get("transcodeStatus")
            done = False
            if last.get("isVoice"):
                if transcode in ("fail", "no_rights"):
                    done = True
                elif transcode == "suc" and status in (2, 3, 4, 5) and text:
                    done = True
                elif transcode == "suc" and status in (2, 4, 5):
                    done = True
            else:
                if status in (2, 3, 4, 5) or text:
                    done = True
            if not poll or done or (time.time() - start) > timeout:
                return _send_json(
                    self,
                    {
                        "ok": True,
                        "detail": last,
                        "text": text,
                        "done": done,
                        "elapsed": int(time.time() - start),
                    },
                )
            time.sleep(max(2, interval))

    def _transcribe(self, body: dict[str, Any]) -> None:
        auth = self._require_auth(body)
        if not auth:
            return
        base_url, token = auth
        file_id = body.get("fileId")
        if not file_id:
            return _send_json(self, {"ok": False, "error": "请提供 fileId"}, 400)
        payload = {
            "fileId": int(file_id),
            "language": body.get("language") or "zh",
            "hasSpeakers": bool(body.get("hasSpeakers", True)),
            "detailLevel": body.get("detailLevel") or "more_details",
        }
        if body.get("model"):
            payload["model"] = body["model"]
        if body.get("template"):
            payload["template"] = body["template"]
        result = _api_request(
            "POST",
            f"{base_url}/api/v1/task/transcribe/commit",
            token=token,
            payload=payload,
        )
        if not result["ok"]:
            return _send_json(self, {"ok": False, "error": result.get("error")}, 400)
        return _send_json(self, {"ok": True, "data": result["data"]})

    def _translate(self, body: dict[str, Any]) -> None:
        auth = self._require_auth(body)
        if not auth:
            return
        base_url, token = auth
        tid = body.get("transcribeId")
        lang = body.get("targetLanguage") or "en"
        if not tid:
            return _send_json(self, {"ok": False, "error": "请提供 transcribeId"}, 400)
        result = _api_request(
            "POST",
            f"{base_url}/api/v1/translate",
            token=token,
            payload={"transcribeId": int(tid), "targetLanguage": lang},
        )
        if not result["ok"]:
            return _send_json(self, {"ok": False, "error": result.get("error")}, 400)
        return _send_json(self, {"ok": True, "data": result["data"]})

    def _deep_research(self, body: dict[str, Any]) -> None:
        auth = self._require_auth(body)
        if not auth:
            return
        base_url, token = auth
        question = (body.get("question") or "").strip()
        if not question:
            return _send_json(self, {"ok": False, "error": "请填写研究问题"}, 400)

        session_id = body.get("sessionId")
        session_type = body.get("sessionType")
        record_id = body.get("recordId")

        if record_id and (not session_id or session_type is None):
            detail_res = _api_request(
                "GET",
                f"{base_url}/api/v2/file-index/file-detail/{record_id}",
                token=token,
            )
            if not detail_res["ok"]:
                return _send_json(self, {"ok": False, "error": detail_res.get("error")}, 400)
            data = detail_res["data"]
            detail = data.get("data") if isinstance(data, dict) and "data" in data else data
            session_id = detail.get("dprSessionId") or session_id
            if session_type is None:
                session_type = 6 if detail.get("isVoice") else 9

        if not session_id:
            return _send_json(self, {"ok": False, "error": "缺少 sessionId / recordId"}, 400)
        if session_type is None:
            session_type = 9

        payload = {
            "sessionId": str(session_id),
            "sessionType": int(session_type),
            "question": question,
            "msgId": int(time.time() * 1000),
            "source": int(body.get("source") or 3),
        }
        if body.get("outline"):
            payload["outline"] = body["outline"]
        result = _api_request(
            "POST",
            f"{base_url}/api/v1/deep/research/query",
            token=token,
            payload=payload,
            timeout=120,
        )
        if not result["ok"]:
            return _send_json(self, {"ok": False, "error": result.get("error")}, 400)
        return _send_json(self, {"ok": True, "data": result["data"]})

    def _sync_project(self, body: dict[str, Any]) -> None:
        """同步项目文件：列表 → 逐个拉取详情与正文 → 返回时间线条目。"""
        auth = self._require_auth(body)
        if not auth:
            return
        base_url, token = auth
        project_id = body.get("projectId") or body.get("rootId") or body.get("chatId")
        if not project_id:
            return _send_json(self, {"ok": False, "error": "请提供 projectId 或 chatId"}, 400)

        auto_transcribe = bool(body.get("autoTranscribe", True))
        language = body.get("language") or "zh"
        limit = int(body.get("limit") or 40)

        files_url = f"{base_url}/api/v1/file-index/file-tree?" + urllib.parse.urlencode(
            {"rootId": str(project_id)}
        )
        files_res = _api_request("GET", files_url, token=token)
        if not files_res["ok"]:
            return _send_json(self, {"ok": False, "error": files_res.get("error")}, 400)
        tree = files_res["data"].get("fileTree") if isinstance(files_res["data"], dict) else []
        flat = [f for f in self._flatten(tree or []) if f.get("type") != "directory"]
        flat = flat[:limit]

        entries = []
        for f in flat:
            record_id = f.get("id")
            file_id = f.get("fileId")
            detail_res = _api_request(
                "GET",
                f"{base_url}/api/v2/file-index/file-detail/{record_id}",
                token=token,
            )
            detail = {}
            text = ""
            if detail_res["ok"]:
                data = detail_res["data"]
                detail = data.get("data") if isinstance(data, dict) and "data" in data else data
                if not isinstance(detail, dict):
                    detail = {}
                text = _extract_transcript(detail)

            is_voice = bool(detail.get("isVoice")) or str(f.get("fileType") or "").lower() in {
                "mp3",
                "wav",
                "m4a",
                "aac",
                "opus",
                "mp4",
                "mov",
                "caf",
                "upload_recording",
                "recording_file",
            }
            needs_transcribe = is_voice and not text and f.get("transcodeStatus") != "ing"
            transcribed = False
            if auto_transcribe and needs_transcribe and file_id:
                _api_request(
                    "POST",
                    f"{base_url}/api/v1/task/transcribe/commit",
                    token=token,
                    payload={
                        "fileId": int(file_id),
                        "language": language,
                        "hasSpeakers": True,
                        "detailLevel": "more_details",
                    },
                )
                # 短轮询一次
                for _ in range(6):
                    time.sleep(3)
                    poll = _api_request(
                        "GET",
                        f"{base_url}/api/v2/file-index/file-detail/{record_id}",
                        token=token,
                    )
                    if not poll["ok"]:
                        break
                    pdata = poll["data"]
                    detail = pdata.get("data") if isinstance(pdata, dict) and "data" in pdata else pdata
                    if not isinstance(detail, dict):
                        detail = {}
                    text = _extract_transcript(detail)
                    if text:
                        transcribed = True
                        break

            entries.append(
                {
                    "recordId": record_id,
                    "fileId": file_id,
                    "name": detail.get("fileName") or detail.get("title") or f.get("name"),
                    "fileType": detail.get("fileType") or f.get("fileType"),
                    "isVoice": is_voice,
                    "status": detail.get("status"),
                    "transcodeStatus": detail.get("transcodeStatus") or f.get("transcodeStatus"),
                    "duration": detail.get("duration"),
                    "language": detail.get("language"),
                    "text": text,
                    "summary": _extract_transcript({"summaryJson": detail.get("summaryJson")})
                    if detail.get("summaryJson")
                    else "",
                    "transcribeId": detail.get("transcribeId") or f.get("transcribeId"),
                    "summaryId": detail.get("summaryId") or f.get("summaryId"),
                    "dprSessionId": detail.get("dprSessionId"),
                    "updatedAt": detail.get("updateTime") or detail.get("createTime"),
                    "justTranscribed": transcribed,
                    "source": "ticnote",
                    "projectId": str(project_id),
                }
            )

        return _send_json(
            self,
            {
                "ok": True,
                "projectId": str(project_id),
                "count": len(entries),
                "entries": entries,
            },
        )

    def _fetch_url(self, body: dict[str, Any]) -> None:
        url = (body.get("url") or "").strip()
        if not url.startswith(("http://", "https://")):
            return _send_json(self, {"ok": False, "error": "请输入有效的 http(s) 网址"}, 400)
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": "TicNoteStudio/1.0 (+local; knowledge clipper)",
                "Accept": "text/html,application/xhtml+xml,text/plain;q=0.9,*/*;q=0.8",
            },
            method="GET",
        )
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read()
                ctype = resp.headers.get("Content-Type", "")
                charset = "utf-8"
                if "charset=" in ctype:
                    charset = ctype.split("charset=")[-1].split(";")[0].strip() or "utf-8"
                html = raw.decode(charset, errors="replace")
        except Exception as e:  # noqa: BLE001
            return _send_json(self, {"ok": False, "error": f"抓取失败: {e}"}, 400)

        title = ""
        import re

        m = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
        if m:
            title = re.sub(r"\s+", " ", m.group(1)).strip()
        text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
        text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
        text = re.sub(r"(?is)<noscript[^>]*>.*?</noscript>", " ", text)
        text = re.sub(r"(?is)<[^>]+>", " ", text)
        text = re.sub(r"&nbsp;", " ", text)
        text = re.sub(r"&amp;", "&", text)
        text = re.sub(r"&lt;", "<", text)
        text = re.sub(r"&gt;", ">", text)
        text = re.sub(r"\s+", " ", text).strip()
        if len(text) > 20000:
            text = text[:20000] + "…"
        return _send_json(self, {"ok": True, "title": title or url, "text": text, "url": url})

    def _ai_chat(self, body: dict[str, Any]) -> None:
        """可选 OpenAI 兼容接口；用于本地知识库问答增强。"""
        endpoint = (body.get("endpoint") or "").rstrip("/")
        api_key = body.get("apiKey") or ""
        model = body.get("model") or "gpt-4o-mini"
        messages = body.get("messages") or []
        if not endpoint or not api_key:
            return _send_json(self, {"ok": False, "error": "未配置 AI 端点"}, 400)
        url = endpoint if endpoint.endswith("/chat/completions") else endpoint + "/chat/completions"
        result = _api_request(
            "POST",
            url,
            token=api_key,
            payload={"model": model, "messages": messages, "temperature": 0.3},
            timeout=90,
        )
        if not result["ok"]:
            return _send_json(self, {"ok": False, "error": result.get("error")}, 400)
        data = result["data"]
        try:
            content = data["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError):
            return _send_json(self, {"ok": False, "error": "AI 响应格式异常", "raw": data}, 502)
        return _send_json(self, {"ok": True, "content": content, "raw": data})


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"TicNote Studio 已启动: http://{HOST}:{PORT}")
    print("在浏览器打开上述地址，填入 TicNote AppKey 即可同步知识库。")
    print("AppKey 获取: https://www.ticnote.cn → 个人中心 → TicNote Key")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")


if __name__ == "__main__":
    main()
