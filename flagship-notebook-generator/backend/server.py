#!/usr/bin/env python3
"""Stateless HTTP API for the flagship notebook generator."""

from __future__ import annotations

import json
import os
import re
import sys
import tempfile
import zipfile
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from notebook_engine import DEFAULT_MODEL, generate_notebook_artifacts, normalize_whitespace, slugify


ROOT = Path(__file__).resolve().parent
HOST = os.getenv("HOST", "0.0.0.0")
PORT = int(os.getenv("PORT", "8787"))
MAX_UPLOAD_BYTES = 80 * 1024 * 1024
ALLOW_ORIGIN = os.getenv("ALLOW_ORIGIN", "*")


def allow_origin(request_origin: str | None = None) -> str:
    if ALLOW_ORIGIN == "*":
        return "*"
    if request_origin and request_origin in {origin.strip() for origin in ALLOW_ORIGIN.split(",") if origin.strip()}:
        return request_origin
    return ALLOW_ORIGIN.split(",")[0].strip() if ALLOW_ORIGIN else "*"


def send_cors_headers(handler: BaseHTTPRequestHandler) -> None:
    request_origin = handler.headers.get("Origin")
    origin = allow_origin(request_origin)
    handler.send_header("Access-Control-Allow-Origin", origin)
    handler.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    handler.send_header("Access-Control-Allow-Headers", "Content-Type, X-OpenAI-API-Key")
    handler.send_header("Access-Control-Max-Age", "86400")
    handler.send_header("Vary", "Origin")


def json_response(handler: BaseHTTPRequestHandler, status: int, payload: dict[str, Any]) -> None:
    body = json.dumps(payload).encode("utf-8")
    handler.send_response(status)
    send_cors_headers(handler)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def text_response(handler: BaseHTTPRequestHandler, text: str, status: int = 200) -> None:
    body = text.encode("utf-8")
    handler.send_response(status)
    send_cors_headers(handler)
    handler.send_header("Content-Type", "text/plain; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


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


def truthy(value: str) -> bool:
    return value.strip().lower() in {"1", "true", "yes", "on"}


def build_bundle_name(source_name: str) -> str:
    return f"{slugify(Path(source_name).stem)[:48]}-notebook-bundle.zip"


def zip_artifacts(zip_path: Path, *, source_path: Path, result: dict[str, Any]) -> None:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as bundle:
        bundle.write(source_path, arcname=f"input/{source_path.name}")
        bundle.write(result["plan_path"], arcname="diagnostics/notebook_plan.json")
        bundle.write(result["deck_path"], arcname="diagnostics/source_deck.json")
        bundle.write(result["outputs"]["session1"], arcname="outputs/Session 1 - Student Notebook.pptx")
        bundle.write(result["outputs"]["session2"], arcname="outputs/Session 2 - Student Notebook.pptx")
        assets_dir = result["output_dir"] / "assets"
        if assets_dir.exists():
            for asset in sorted(assets_dir.rglob("*")):
                if asset.is_file():
                    bundle.write(asset, arcname=f"assets/{asset.relative_to(assets_dir).as_posix()}")


def openai_api_key(handler: BaseHTTPRequestHandler) -> str:
    return handler.headers.get("X-OpenAI-API-Key", "").strip() or os.getenv("OPENAI_API_KEY", "").strip()


class NotebookApiHandler(BaseHTTPRequestHandler):
    server_version = "FlagshipNotebookGenerator/1.0"

    def do_OPTIONS(self) -> None:  # noqa: N802
        self.send_response(204)
        send_cors_headers(self)
        self.end_headers()

    def do_GET(self) -> None:  # noqa: N802
        if self.path == "/api/health":
            json_response(
                self,
                200,
                {
                    "ok": True,
                    "service": "flagship-notebook-generator",
                    "default_model": DEFAULT_MODEL,
                },
            )
            return
        text_response(self, "Not found", status=404)

    def do_POST(self) -> None:  # noqa: N802
        if self.path == "/api/generate":
            self.generate_bundle()
            return
        text_response(self, "Not found", status=404)

    def log_message(self, fmt: str, *args: Any) -> None:
        sys.stdout.write("%s - - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), fmt % args))

    def generate_bundle(self) -> None:
        try:
            fields, files = parse_multipart(self)
            upload = files.get("deck")
            if not upload:
                raise ValueError("A PowerPoint deck is required.")

            filename = clean_filename(upload["filename"])
            if not filename.lower().endswith(".pptx"):
                raise ValueError("Please upload a .pptx file.")

            offline = truthy(fields.get("offline", ""))
            api_key = openai_api_key(self)
            if not offline and not api_key:
                raise ValueError("Provide an OpenAI API key or use offline mode.")

            model = normalize_whitespace(fields.get("model", "")) or DEFAULT_MODEL
            guidance = normalize_whitespace(fields.get("guidance", ""))

            with tempfile.TemporaryDirectory(prefix="flagship-notebook-") as temp_dir_name:
                temp_dir = Path(temp_dir_name)
                source_path = temp_dir / filename
                source_path.write_bytes(upload["content"])

                result = generate_notebook_artifacts(
                    source_path,
                    output_dir=temp_dir / "build",
                    model=model,
                    offline=offline,
                    custom_guidance=guidance,
                    api_key=api_key,
                )

                zip_path = temp_dir / build_bundle_name(filename)
                zip_artifacts(zip_path, source_path=source_path, result=result)
                payload = zip_path.read_bytes()

            self.send_response(200)
            send_cors_headers(self)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Content-Disposition", f'attachment; filename="{build_bundle_name(filename)}"')
            self.end_headers()
            self.wfile.write(payload)
        except ValueError as exc:
            json_response(self, HTTPStatus.BAD_REQUEST, {"error": str(exc)})
        except Exception as exc:
            json_response(self, HTTPStatus.INTERNAL_SERVER_ERROR, {"error": str(exc)})


def run() -> None:
    server = ThreadingHTTPServer((HOST, PORT), NotebookApiHandler)
    print(f"Flagship notebook backend running on http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    run()
