#!/usr/bin/env python3
"""Local server and SQLite persistence for TradeFlow Atlas."""
from __future__ import annotations

import json
import os
import re
import sqlite3
import ssl
import subprocess
import sys
import tempfile
from html import unescape
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlencode, urlparse
from urllib.error import HTTPError
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent
DATABASE = ROOT / "tradeflow-atlas.sqlite3"
NODE = "/Users/abhyudayamarya/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node"
EXPORTER = ROOT.parent / "work" / "export_tradeflow_xlsx.mjs"
CURRICULUM = ROOT.parent / "work" / "curriculum-source" / "curriculum"
ENV_FILE = ROOT / ".env"
SIGNAL_CACHE: dict[str, object] = {"created_at": 0.0, "items": []}
IAM_TOKEN_CACHE: dict[str, object] = {"token": "", "expires_at": 0.0}


def load_local_env() -> None:
    """Load only this app's local .env file; explicitly exported shell values win."""
    if not ENV_FILE.is_file():
        return
    for raw_line in ENV_FILE.read_text(encoding="utf-8", errors="ignore").splitlines():
        match = re.match(r"\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$", raw_line)
        if not match:
            continue
        key, value = match.groups()
        if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
            value = value[1:-1]
        else:
            value = value.split(" #", 1)[0].strip()
        if value and not os.environ.get(key):
            os.environ[key] = value


load_local_env()


def connection() -> sqlite3.Connection:
    db = sqlite3.connect(DATABASE)
    db.row_factory = sqlite3.Row
    db.execute("CREATE TABLE IF NOT EXISTS notes (zone TEXT PRIMARY KEY, content TEXT NOT NULL, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)")
    db.execute("CREATE TABLE IF NOT EXISTS note_entries (id INTEGER PRIMARY KEY AUTOINCREMENT, zone TEXT NOT NULL, content TEXT NOT NULL, created_at TEXT DEFAULT CURRENT_TIMESTAMP)")
    db.execute("CREATE TABLE IF NOT EXISTS capsules (zone TEXT PRIMARY KEY, content TEXT NOT NULL, unlock_date TEXT NOT NULL, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)")
    db.execute("CREATE TABLE IF NOT EXISTS progress (zone TEXT PRIMARY KEY, checked_keys TEXT NOT NULL, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)")
    db.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)")
    columns = {row[1] for row in db.execute("PRAGMA table_info(progress)")}
    if "checked_keys" not in columns:
        db.execute("ALTER TABLE progress ADD COLUMN checked_keys TEXT NOT NULL DEFAULT '[]'")
    db.execute("INSERT INTO note_entries(zone, content, created_at) SELECT zone, content, updated_at FROM notes WHERE NOT EXISTS (SELECT 1 FROM note_entries n WHERE n.zone = notes.zone AND n.content = notes.content)")
    return db


def plain_note(value: str) -> str:
    """Keep exports readable when rich-text notes contain browser HTML."""
    value = re.sub(r"<(br|/p|/div|/li|/h[1-6])[^>]*>", "\n", value, flags=re.I)
    value = re.sub(r"<[^>]+>", "", value)
    return re.sub(r"\n{3,}", "\n\n", value).strip()


def progress_text(value: str) -> str:
    try:
        entries = json.loads(value)
    except json.JSONDecodeError:
        return value
    return "\n".join(f"✓ {x.get('label', x.get('key', 'Completed task'))}" if isinstance(x, dict) else f"✓ {x}" for x in entries)


def trusted_urlopen(request: Request, timeout: int | None):
    """Use certifi in the project virtual environment on macOS, where system roots can be unavailable."""
    try:
        import certifi
        context = ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        context = ssl.create_default_context()
    return urlopen(request, timeout=timeout, context=context)


def relevant_curriculum(question: str, limit: int = 3) -> list[dict[str, str]]:
    """Retrieve a small, relevant slice of the local curriculum without notes."""
    tokens = {item.lower() for item in re.findall(r"[A-Za-z][A-Za-z0-9&/-]{2,}", question)}
    candidates: list[tuple[int, str, str]] = []
    for file in CURRICULUM.glob("*.md"):
        text = file.read_text(encoding="utf-8", errors="ignore")
        chunks = re.split(r"\n(?=#{1,3}\s)", text)
        for chunk in chunks:
            cleaned = re.sub(r"\s+", " ", chunk).strip()
            # Code-heavy notebook cells are poor retrieval context for a learning answer.
            if len(cleaned) < 70 or cleaned.startswith("# ===") or re.search(r"\bdef\s+\w+\s*\(|client\.chat|response\.choices", cleaned):
                continue
            score = sum(cleaned.lower().count(token) for token in tokens)
            if score:
                title_match = re.search(r"^#{1,3}\s+(.+)", chunk, re.M)
                title = title_match.group(1).strip("# =-") if title_match else file.stem.replace("-", " ")
                candidates.append((score, title, cleaned[:700]))
    candidates.sort(key=lambda item: item[0], reverse=True)
    return [{"title": title, "excerpt": excerpt, "source": f"Your curriculum · {title}"} for _, title, excerpt in candidates[:limit]]


def local_guidance(question: str, curriculum: list[dict[str, str]]) -> str:
    """Server-side guidance when live model inference is unavailable."""
    rules = (
        (r"promo|promotion|trade spend|ropi|roi|net revenue", "For promotion economics, start with baseline versus incremental volume, then test trade spend and margin sensitivity. Open Zone 01 for trade-investment foundations and Zone 03 for the Promo ROI Analyzer."),
        (r"claim|deduction|chargeback|debit memo", "Frame the claims question as: reason-code pattern → financial exposure → root cause → owner and corrective action. Use Zone 01 for governance and Zone 03 for the Claims Analysis Agent."),
        (r"forecast|demand|uplift|s&op|ibp|planning", "Separate statistical baseline, promotional uplift, constraints, and the decision needed before asking an AI tool for interpretation. Start in Zone 01 and continue to Zone 02 for planning GenAI patterns."),
        (r"risk|control tower|inventory|supply|service", "Prioritise risk by customer/service impact, financial exposure, time sensitivity, root cause, and escalation owner. Zone 04 contains the Supply Risk Copilot."),
        (r"watson|api|granite|model|rag|agent", "Use the watsonx.ai build standard: API key, regional URL, project ID, source grounding, and functional expert review before recommendations are used."),
        (r"capstone|portfolio|interview|consultant", "Lead the portfolio story with the business problem, analytical method and control points, then the client decision or measurable operating impact. Zone 05 is your consulting portfolio zone."),
    )
    for pattern, answer in rules:
        if re.search(pattern, question, re.I):
            return answer
    if curriculum:
        excerpt = re.sub(r"\s+", " ", curriculum[0]["excerpt"]).strip()
        return f"From your curriculum: {excerpt[:420]}\n\nFor a sharper answer, name the planning area, retailer/CPG context, and decision you are trying to make."
    return "I can help with promotion ROI, claims, S&OP, supply risk, watsonx.ai, or your capstone portfolio. Tell me the decision you are trying to make."


def web_search(question: str, limit: int = 2) -> list[dict[str, str]]:
    """Best-effort public web retrieval. It sends only the question, never saved notes."""
    try:
        url = "https://html.duckduckgo.com/html/?" + urlencode({"q": question})
        request = Request(url, headers={"User-Agent": "TradeFlow-Atlas/1.0 learning research"})
        with trusted_urlopen(request, timeout=3) as response:
            page = response.read().decode("utf-8", errors="ignore")
        matches = re.findall(r'<a[^>]+class="result__a"[^>]+href="([^"]+)"[^>]*>(.*?)</a>', page, re.S)
        results = []
        for href, title in matches:
            clean_title = re.sub(r"<[^>]+>", "", unescape(title)).strip()
            parsed = urlparse(unescape(href))
            query = parse_qs(parsed.query)
            link = unquote(query.get("uddg", [href])[0])
            source_host = (urlparse(link).hostname or "").lower()
            if clean_title and link.startswith(("https://", "http://")) and source_host != "biztoc.com" and not source_host.endswith(".biztoc.com"):
                results.append({"title": clean_title[:150], "url": link})
            if len(results) >= limit:
                break
        return results
    except Exception:
        return []


def news_api_key() -> str:
    """Allow the simple NEWS_API name used in this project's .env file."""
    return os.environ.get("NEWS_API") or os.environ.get("NEWS_API_KEY") or os.environ.get("NEWSAPI_KEY", "")


def industry_signals(force: bool = False) -> dict[str, object]:
    """Fetch a narrow, public learning feed; no notes or user queries leave the device."""
    import time

    cached = SIGNAL_CACHE.get("items", [])
    if cached and not force and time.time() - float(SIGNAL_CACHE["created_at"]) < 1800:
        return {"items": cached, "cached": True}
    api_key = news_api_key()
    if not api_key:
        return {"items": [], "error": "Add NEWS_API to outputs/.env to enable Industry Signals."}
    queries = ("supply chain", "retail supply chain")
    try:
        articles = []
        for query in queries:
            url = "https://newsapi.org/v2/everything?" + urlencode({
                "q": query,
                "language": "en",
                "sortBy": "publishedAt",
                "pageSize": "100",
            })
            request = Request(url, headers={"X-Api-Key": api_key, "User-Agent": "TradeFlow-Atlas/1.0"})
            with trusted_urlopen(request, timeout=12) as response:
                articles.extend(json.loads(response.read().decode("utf-8")).get("articles", []))
        items = []
        seen_urls: set[str] = set()
        seen_titles: set[str] = set()
        block_terms = ("arrest", "assault", "missile", "drone strike", "crypto", "celebrity", "luxury car")
        for article in articles:
            title = str(article.get("title") or "").strip()
            description = str(article.get("description") or "").strip()
            article_url = str(article.get("url") or "").strip()
            source_name = str((article.get("source") or {}).get("name") or "Industry source").strip()
            source_host = (urlparse(article_url).hostname or "").lower()
            combined = (title + " " + description + " " + source_name).lower()
            normalized_title = re.sub(r"\W+", "", title.lower())
            title_context = any(term in title.lower() for term in ("supply chain", "logistics", "trade promotion", "trade spend", "net revenue", "demand planning", "inventory", "merchandise planning"))
            if not title or not article_url.startswith(("https://", "http://")) or article_url in seen_urls or normalized_title in seen_titles or not title_context or any(term in combined for term in block_terms) or source_host == "biztoc.com" or source_host.endswith(".biztoc.com") or "biztoc" in source_name.lower():
                continue
            seen_urls.add(article_url)
            seen_titles.add(normalized_title)
            items.append({
                "title": title[:220],
                "description": description[:340],
                "url": article_url,
                "source": source_name[:100],
                "publishedAt": str(article.get("publishedAt") or ""),
            })
            if len(items) >= 12:
                break
        SIGNAL_CACHE.update({"created_at": time.time(), "items": items})
        return {"items": items, "cached": False}
    except HTTPError as error:
        return {"items": [], "error": f"NewsAPI returned HTTP {error.code}. Try refreshing later."}
    except Exception:
        return {"items": [], "error": "Industry Signals could not refresh right now. Try again later."}


def watsonx_config() -> tuple[dict[str, str] | None, list[str]]:
    values = {
        "apikey": os.environ.get("WATSONX_APIKEY", ""),
        "project_id": os.environ.get("WATSONX_PROJECT_ID", ""),
        "url": os.environ.get("WATSONX_URL", ""),
        "model_id": os.environ.get("WATSONX_MODEL_ID", "ibm/granite-4-h-small"),
    }
    missing = [name for name in ("WATSONX_APIKEY", "WATSONX_PROJECT_ID", "WATSONX_URL") if not os.environ.get(name)]
    return (values if not missing else None), missing


def generate_watsonx_answer(question: str, curriculum: list[dict[str, str]], sources: list[dict[str, str]]) -> str:
    config, missing = watsonx_config()
    if not config:
        raise RuntimeError("Configure " + ", ".join(missing) + " locally before enabling live AI.")
    local_context = "\n\n".join(f"[{item['source']}]\n{item['excerpt']}" for item in curriculum) or "No directly matching local excerpt was found."
    web_context = "\n".join(f"- {item['title']}: {item['url']}" for item in sources) or "(None.)"
    system = """You are Anusha, a practical strategic supply-chain planning and retail TPM learning assistant. Use the local curriculum as the primary grounding. Use web sources only as supplemental reading leads; do not claim facts not present in the supplied source list. Do not invent client data, citations, or metrics. Distinguish fact from recommendation. Answer in 120–180 words maximum, using at most 4 short bullets when useful. Finish every sentence and return only the helpful answer."""
    user_content = f"""USER QUESTION:
{question}

LOCAL CURRICULUM (private to this local app; never include saved notes):
{local_context}

WEB RESEARCH LEADS:
{web_context}"""
    import time
    if time.time() >= float(IAM_TOKEN_CACHE["expires_at"]):
        token_request = Request(
            "https://iam.cloud.ibm.com/identity/token",
            data=urlencode({"grant_type": "urn:ibm:params:oauth:grant-type:apikey", "apikey": config["apikey"]}).encode(),
            headers={"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"},
        )
        with trusted_urlopen(token_request, timeout=20) as token_response:
            token_payload = json.loads(token_response.read().decode("utf-8"))
        IAM_TOKEN_CACHE.update({"token": token_payload["access_token"], "expires_at": time.time() + int(token_payload.get("expires_in", 3600)) - 60})
    endpoint = config["url"].rstrip("/") + "/ml/v1/text/chat?version=2026-08-02"
    request_body = {
        "model_id": config["model_id"],
        "project_id": config["project_id"],
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": [{"type": "text", "text": user_content}]},
        ],
        "max_tokens": 280,
        "temperature": 0.2,
    }
    chat_request = Request(endpoint, data=json.dumps(request_body).encode(), headers={"Authorization": "Bearer " + str(IAM_TOKEN_CACHE["token"]), "Content-Type": "application/json", "Accept": "application/json"})
    with trusted_urlopen(chat_request, timeout=None) as chat_response:
        payload = json.loads(chat_response.read().decode("utf-8"))
    answer = str(payload["choices"][0]["message"]["content"]).strip()
    # Avoid presenting a half-finished final sentence if a provider stops at its token limit.
    if answer and answer[-1] not in ".!?`":
        complete = re.findall(r".*?[.!?](?:\s|$)", answer, flags=re.S)
        if complete:
            answer = "".join(complete).strip()
    return answer or "I could not form a complete response. Please try a more specific question."


def generate_watsonx_answer_with_timeout(question: str, curriculum: list[dict[str, str]], sources: list[dict[str, str]]) -> str:
    """Compatibility wrapper: inference is intentionally allowed to finish without a time cutoff."""
    return generate_watsonx_answer(question, curriculum, sources)


class TradeFlowHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def send_json(self, payload: dict, status: int = HTTPStatus.OK) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_json(self) -> dict:
        size = int(self.headers.get("Content-Length", "0"))
        return json.loads(self.rfile.read(size) or b"{}")

    def do_GET(self) -> None:
        request = urlparse(self.path)
        if request.path == "/":
            self.path = "/tradeflow-atlas.html"
            return super().do_GET()
        if request.path == "/api/state":
            zone = parse_qs(request.query).get("zone", ["Home"])[0]
            with connection() as db:
                capsule = db.execute("SELECT content, unlock_date FROM capsules WHERE zone = ?", (zone,)).fetchone()
                progress = db.execute("SELECT checked_keys FROM progress WHERE zone = ?", (zone,)).fetchone()
            return self.send_json({
                "note": "",
                "capsule": dict(capsule) if capsule else None,
                "checkedKeys": json.loads(progress["checked_keys"]) if progress else [],
            })
        if request.path == "/api/settings/theme":
            with connection() as db:
                row = db.execute("SELECT value FROM settings WHERE key = 'theme'").fetchone()
            return self.send_json({"theme": row["value"] if row else "light"})
        if request.path == "/api/assistant/health":
            config, missing = watsonx_config()
            try:
                import ibm_watsonx_ai  # noqa: F401
                package_ready = True
            except ImportError:
                package_ready = False
            return self.send_json({
                "ready": bool(config and package_ready),
                "packageReady": package_ready,
                "missingEnvironment": missing,
                "privacy": "Saved notes and time capsules are never sent to the AI or web search.",
            })
        if request.path == "/api/industry-signals":
            force = parse_qs(request.query).get("refresh", [""])[0] == "1"
            return self.send_json(industry_signals(force=force))
        if request.path == "/api/export.xlsx":
            with connection() as db:
                payload = {
                    "notes": [{"zone": x["zone"], "content": plain_note(x["content"]), "updated_at": x["created_at"]} for x in db.execute("SELECT zone, content, created_at FROM note_entries ORDER BY id DESC")],
                    "capsules": [dict(x) for x in db.execute("SELECT zone, content, unlock_date, updated_at FROM capsules ORDER BY unlock_date")],
                    "progress": [{**dict(x), "checked_keys": progress_text(x["checked_keys"])} for x in db.execute("SELECT zone, checked_keys, updated_at FROM progress ORDER BY updated_at DESC")],
                    "settings": [dict(x) for x in db.execute("SELECT key, value, updated_at FROM settings ORDER BY key")],
                }
            with tempfile.TemporaryDirectory() as temp_dir:
                data_file = Path(temp_dir) / "tradeflow-data.json"
                output_file = Path(temp_dir) / "tradeflow-atlas-export.xlsx"
                data_file.write_text(json.dumps(payload), encoding="utf-8")
                subprocess.run([NODE, str(EXPORTER), str(data_file), str(output_file)], check=True, capture_output=True, text=True)
                export = output_file.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", "attachment; filename=tradeflow-atlas-export.xlsx")
            self.send_header("Content-Length", str(len(export)))
            self.end_headers()
            self.wfile.write(export)
            return
        return super().do_GET()

    def do_POST(self) -> None:
        request = urlparse(self.path)
        if request.path != "/api/assistant":
            return self.send_json({"error": "Unknown endpoint"}, HTTPStatus.NOT_FOUND)
        try:
            data = self.read_json()
        except json.JSONDecodeError:
            return self.send_json({"error": "Invalid JSON request"}, HTTPStatus.BAD_REQUEST)
        question = str(data.get("question", "")).strip()
        if not question:
            return self.send_json({"error": "Ask a question first."}, HTTPStatus.BAD_REQUEST)
        curriculum = relevant_curriculum(question)
        sources = web_search(question) if data.get("useWeb", True) else []
        try:
            answer = generate_watsonx_answer_with_timeout(question, curriculum, sources)
        except RuntimeError as error:
            return self.send_json({
                "answer": local_guidance(question, curriculum),
                "notice": str(error),
                "curriculum": curriculum,
                "sources": sources,
                "mode": "local-guidance",
            })
        except Exception:
            return self.send_json({
                "error": "watsonx.ai could not complete that request. Check your regional URL, project access, and model ID.",
                "curriculum": curriculum,
                "sources": sources,
            }, HTTPStatus.BAD_GATEWAY)
        return self.send_json({"answer": answer, "curriculum": curriculum, "sources": sources})

    def do_PUT(self) -> None:
        request = urlparse(self.path)
        data = self.read_json()
        with connection() as db:
            if request.path == "/api/notes":
                db.execute("INSERT INTO note_entries(zone, content) VALUES(?, ?)", (data["zone"], data.get("content", "")))
            elif request.path == "/api/capsules":
                db.execute("INSERT INTO capsules(zone, content, unlock_date) VALUES(?, ?, ?) ON CONFLICT(zone) DO UPDATE SET content=excluded.content, unlock_date=excluded.unlock_date, updated_at=CURRENT_TIMESTAMP", (data["zone"], data["content"], data["unlockDate"]))
            elif request.path == "/api/progress":
                db.execute("INSERT INTO progress(zone, checked_keys) VALUES(?, ?) ON CONFLICT(zone) DO UPDATE SET checked_keys=excluded.checked_keys, updated_at=CURRENT_TIMESTAMP", (data["zone"], json.dumps(data.get("checkedKeys", []))))
            elif request.path == "/api/settings/theme":
                db.execute("INSERT INTO settings(key, value) VALUES('theme', ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value, updated_at=CURRENT_TIMESTAMP", (data["theme"],))
            else:
                return self.send_json({"error": "Unknown endpoint"}, HTTPStatus.NOT_FOUND)
        self.send_json({"ok": True})


if __name__ == "__main__":
    host, port = "127.0.0.1", int(os.environ.get("TRADEFLOW_PORT", "8799"))
    print(f"TradeFlow Atlas is running at http://{host}:{port}")
    ThreadingHTTPServer((host, port), TradeFlowHandler).serve_forever()
