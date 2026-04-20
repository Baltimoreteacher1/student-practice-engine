#!/usr/bin/env python3
"""Local upload app for the student notebook engine."""

from __future__ import annotations

import json
import os
import re
import shutil
import sys
from datetime import datetime
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import quote, unquote

from notebook_engine import (
    PREMIUM_NOTEBOOK_GUIDANCE,
    choose_api_key,
    enforce_runtime_quality_guidance,
    generate_notebook_artifacts,
    normalize_whitespace,
    normalize_saved_guidance,
    slugify,
)
from notebook_launchers import ensure_inbox_launcher, ensure_lesson_plan_assets


ROOT = Path(__file__).resolve().parent
APP_HOME = ROOT / "notebook_engine_app_data"
CONFIG_PATH = APP_HOME / "config.json"
RUNS_DIR = APP_HOME / "runs"
MAX_UPLOAD_BYTES = 80 * 1024 * 1024
DEFAULT_PORT = 8765
APP_DEFAULT_MODEL = "gpt-5.4"
DEFAULT_APP_GUIDANCE = PREMIUM_NOTEBOOK_GUIDANCE


def ensure_app_dirs() -> None:
    APP_HOME.mkdir(parents=True, exist_ok=True)
    RUNS_DIR.mkdir(parents=True, exist_ok=True)


def default_config() -> dict[str, Any]:
    return {
        "api_key": "",
        "model": APP_DEFAULT_MODEL,
        "default_guidance": DEFAULT_APP_GUIDANCE,
    }


def recover_config_payload(raw_text: str) -> dict[str, Any]:
    recovered: dict[str, Any] = {}
    patterns = {
        "api_key": r'"api_key"\s*:\s*"(.*?)"\s*,\s*"model"',
        "model": r'"model"\s*:\s*"(.*?)"',
        "default_guidance": r'"default_guidance"\s*:\s*"(.*?)"\s*}',
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, raw_text, re.S)
        if not match:
            continue
        recovered[key] = match.group(1).replace("\n", " ").strip()
    return recovered


def normalize_config_payload(data: dict[str, Any]) -> dict[str, Any]:
    config = default_config()
    config["api_key"] = choose_api_key(str(data.get("api_key", "")))
    config["model"] = normalize_whitespace(str(data.get("model", ""))) or APP_DEFAULT_MODEL
    config["default_guidance"] = normalize_saved_guidance(str(data.get("default_guidance", "")) or DEFAULT_APP_GUIDANCE)
    return config


def load_config() -> dict[str, Any]:
    ensure_app_dirs()
    if not CONFIG_PATH.exists():
        return default_config()
    raw_text = CONFIG_PATH.read_text(encoding="utf-8")
    try:
        data = json.loads(raw_text)
    except Exception:
        data = recover_config_payload(raw_text)
        if not data:
            return default_config()
    config = normalize_config_payload(data)
    save_config(config)
    return config


def save_config(config: dict[str, Any]) -> None:
    ensure_app_dirs()
    normalized = normalize_config_payload(config)
    CONFIG_PATH.write_text(json.dumps(normalized, indent=2), encoding="utf-8")


def public_settings(config: dict[str, Any]) -> dict[str, Any]:
    return {
        "has_api_key": bool(choose_api_key(config.get("api_key", ""))),
        "model": config.get("model") or APP_DEFAULT_MODEL,
        "default_guidance": config.get("default_guidance") or DEFAULT_APP_GUIDANCE,
    }


def json_response(handler: BaseHTTPRequestHandler, status: int, payload: dict[str, Any]) -> None:
    body = json.dumps(payload).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def html_response(handler: BaseHTTPRequestHandler, html: str, status: int = 200) -> None:
    body = html.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def text_response(handler: BaseHTTPRequestHandler, text: str, status: int = 200) -> None:
    body = text.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "text/plain; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def read_json_body(handler: BaseHTTPRequestHandler) -> dict[str, Any]:
    length = int(handler.headers.get("Content-Length", "0") or "0")
    raw = handler.rfile.read(length)
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def parse_multipart(handler: BaseHTTPRequestHandler) -> tuple[dict[str, str], dict[str, dict[str, Any]]]:
    content_type = handler.headers.get("Content-Type", "")
    length = int(handler.headers.get("Content-Length", "0") or "0")
    if length > MAX_UPLOAD_BYTES:
        raise ValueError("Upload is too large.")
    body = handler.rfile.read(length)
    message = BytesParser(policy=default).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )
    if not message.is_multipart():
        raise ValueError("Expected multipart form upload.")

    fields: dict[str, str] = {}
    files: dict[str, dict[str, Any]] = {}
    for part in message.iter_parts():
        if part.get_content_disposition() != "form-data":
            continue
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""
        if filename:
            files[name] = {
                "filename": filename,
                "content": payload,
                "content_type": part.get_content_type(),
            }
            continue
        charset = part.get_content_charset() or "utf-8"
        fields[name] = payload.decode(charset, errors="replace")
    return fields, files


def clean_filename(filename: str) -> str:
    name = Path(filename).name
    return re.sub(r"[^A-Za-z0-9._ -]", "_", name).strip() or "uploaded_deck.pptx"


def relative_to_app(path: Path) -> str:
    return str(path.resolve().relative_to(APP_HOME.resolve()))


def file_url(path: Path) -> str:
    relative = relative_to_app(path)
    return "/files/" + quote(relative)


def build_result_links(job_dir: Path, result: dict[str, Any]) -> dict[str, str]:
    outputs = result["outputs"]
    links = {
        "session1": file_url(outputs["session1"]),
        "plan": file_url(result["plan_path"]),
        "deck": file_url(result["deck_path"]),
        "job_dir": file_url(job_dir),
    }
    if outputs.get("session2"):
        links["session2"] = file_url(outputs["session2"])
    return links


def build_job_id(source_filename: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    return f"{timestamp}-{slugify(Path(source_filename).stem)[:48]}"


def list_recent_jobs(limit: int = 8) -> list[dict[str, Any]]:
    ensure_app_dirs()
    jobs: list[dict[str, Any]] = []
    for job_dir in sorted(RUNS_DIR.iterdir(), reverse=True):
        if not job_dir.is_dir():
            continue
        manifest_path = job_dir / "job.json"
        if not manifest_path.exists():
            continue
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        manifest["links"] = {
            key: file_url(job_dir / relative)
            for key, relative in manifest.get("relative_files", {}).items()
            if (job_dir / relative).exists()
        }
        jobs.append(manifest)
        if len(jobs) >= limit:
            break
    return jobs


def write_manifest(job_dir: Path, manifest: dict[str, Any]) -> None:
    (job_dir / "job.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")


def generate_job(
    *,
    uploaded_filename: str,
    uploaded_bytes: bytes,
    model: str,
    offline: bool,
    guidance: str,
    api_key: str,
) -> dict[str, Any]:
    ensure_app_dirs()
    job_id = build_job_id(uploaded_filename)
    job_dir = RUNS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    safe_name = clean_filename(uploaded_filename)
    source_path = job_dir / safe_name
    source_path.write_bytes(uploaded_bytes)

    result = generate_notebook_artifacts(
        source_path,
        output_dir=job_dir,
        model=model or APP_DEFAULT_MODEL,
        offline=offline,
        custom_guidance=guidance,
        api_key=api_key,
    )

    relative_files = {
        "source": source_path.name,
        "deck": result["deck_path"].relative_to(job_dir).as_posix(),
        "plan": result["plan_path"].relative_to(job_dir).as_posix(),
        "session1": result["outputs"]["session1"].relative_to(job_dir).as_posix(),
        "quality_report": result["quality_report_path"].relative_to(job_dir).as_posix(),
    }
    if result["outputs"].get("session2"):
        relative_files["session2"] = result["outputs"]["session2"].relative_to(job_dir).as_posix()
    manifest = {
        "job_id": job_id,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "source_filename": safe_name,
        "model": model or APP_DEFAULT_MODEL,
        "offline": offline,
        "guidance": result.get("effective_guidance", guidance),
        "status": "completed",
        "quality_passed": bool(result.get("quality_report", {}).get("passed", False)),
        "relative_files": relative_files,
    }
    write_manifest(job_dir, manifest)
    manifest["links"] = {
        key: file_url(job_dir / relative)
        for key, relative in relative_files.items()
    }
    return manifest


def resolve_frontend_page_path() -> Path | None:
    for name in ("index.html", "Index.html"):
        candidate = ROOT / name
        if candidate.exists():
            return candidate
    return None


def render_page() -> str:
    page_path = resolve_frontend_page_path()
    if page_path is not None:
        return page_path.read_text(encoding="utf-8")
    return """<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Notebook Engine</title></head><body><p>Missing index.html. Put the frontend file next to notebook_engine_app.py.</p></body></html>"""


class NotebookAppHandler(BaseHTTPRequestHandler):
    server_version = "NotebookEngineApp/1.0"

    def do_GET(self) -> None:  # noqa: N802
        if self.path == "/":
            html_response(self, render_page())
            return
        if self.path == "/api/settings":
            json_response(self, 200, public_settings(load_config()))
            return
        if self.path == "/api/recent-jobs":
            json_response(self, 200, {"jobs": list_recent_jobs()})
            return
        if self.path.startswith("/files/"):
            self.serve_file()
            return
        text_response(self, "Not found", status=404)

    def do_POST(self) -> None:  # noqa: N802
        if self.path == "/api/settings":
            self.save_settings()
            return
        if self.path == "/api/generate":
            self.generate()
            return
        text_response(self, "Not found", status=404)

    def log_message(self, fmt: str, *args: Any) -> None:
        sys.stdout.write("%s - - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), fmt % args))

    def save_settings(self) -> None:
        try:
            payload = read_json_body(self)
            config = load_config()
            incoming_model = normalize_whitespace(payload.get("model", "")) or APP_DEFAULT_MODEL
            config["model"] = incoming_model
            config["default_guidance"] = payload.get("default_guidance", "").strip()
            if payload.get("api_key", "").strip():
                config["api_key"] = payload["api_key"].strip()
            save_config(config)
            json_response(self, 200, {"ok": True, "settings": public_settings(config)})
        except Exception as exc:
            json_response(self, 400, {"error": str(exc)})

    def generate(self) -> None:
        try:
            fields, files = parse_multipart(self)
            upload = files.get("deck")
            if not upload:
                raise ValueError("A PowerPoint deck is required.")
            filename = upload["filename"]
            if not filename.lower().endswith(".pptx"):
                raise ValueError("Please upload a .pptx file.")

            config = load_config()
            offline = fields.get("offline", "").strip() in {"1", "true", "on", "yes"}
            model = normalize_whitespace(fields.get("model", "")) or config.get("model") or APP_DEFAULT_MODEL
            default_guidance = config.get("default_guidance", "")
            run_guidance = fields.get("guidance", "")
            guidance = enforce_runtime_quality_guidance(
                normalize_whitespace(" ".join(part for part in [default_guidance, run_guidance] if part))
            )
            api_key = choose_api_key(config.get("api_key", ""))
            if not offline and not api_key:
                raise ValueError("Save an OpenAI API key first, or use offline mode for a smoke test.")

            job = generate_job(
                uploaded_filename=filename,
                uploaded_bytes=upload["content"],
                model=model,
                offline=offline,
                guidance=guidance,
                api_key=api_key,
            )
            json_response(self, 200, {"ok": True, "job": job})
        except ValueError as exc:
            json_response(self, 400, {"error": str(exc)})
        except Exception as exc:
            json_response(self, 500, {"error": str(exc)})

    def serve_file(self) -> None:
        relative = unquote(self.path[len("/files/"):])
        if not relative:
            text_response(self, "Missing file path", status=404)
            return
        target = (APP_HOME / relative).resolve()
        app_root = APP_HOME.resolve()
        if app_root not in target.parents and target != app_root:
            text_response(self, "Forbidden", status=403)
            return
        if target.is_dir():
            listing = "\n".join(item.name for item in sorted(target.iterdir()))
            text_response(self, listing or "(empty)")
            return
        if not target.exists():
            text_response(self, "Not found", status=404)
            return
        content = target.read_bytes()
        suffix = target.suffix.lower()
        content_type = {
            ".json": "application/json; charset=utf-8",
            ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".md": "text/markdown; charset=utf-8",
            ".txt": "text/plain; charset=utf-8",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
        }.get(suffix, "application/octet-stream")
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        if suffix == ".pptx":
            self.send_header("Content-Disposition", f'attachment; filename="{target.name}"')
        self.end_headers()
        self.wfile.write(content)


def build_bundle(destination: Path) -> Path:
    destination = destination.expanduser().resolve()
    if destination.exists():
        shutil.rmtree(destination)
    destination.mkdir(parents=True, exist_ok=True)
    for name in (
        "notebook_engine.py",
        "notebook_engine_app.py",
        "notebook_folder_runner.py",
        "notebook_launchers.py",
        "NOTEBOOK_ENGINE.md",
        "NOTEBOOK_ENGINE_APP.md",
        "launch_notebook_engine.command",
        "process_notebook_inbox.command",
        "activity_library.txt",
        "activity_database.json",
    ):
        shutil.copy2(ROOT / name, destination / name)
    page_path = resolve_frontend_page_path()
    if page_path is None:
        raise FileNotFoundError("Notebook app frontend file was not found next to notebook_engine_app.py.")
    shutil.copy2(page_path, destination / "index.html")
    return destination


def main() -> int:
    if len(sys.argv) > 1 and sys.argv[1] == "--bundle":
        target = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "notebook-engine-app"
        built = build_bundle(target)
        print(built)
        return 0

    host = "127.0.0.1"
    port = DEFAULT_PORT
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass

    ensure_app_dirs()
    ensure_inbox_launcher(workspace_root=ROOT)
    if (ROOT / "codex-lesson-plan-generator" / "run.py").exists():
        ensure_lesson_plan_assets(workspace_root=ROOT)
    server = ThreadingHTTPServer((host, port), NotebookAppHandler)
    print(f"Notebook Engine app running at http://{host}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
