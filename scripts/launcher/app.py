"""
Local script launcher.

Run with:  powershell -ExecutionPolicy Bypass -File run.ps1
Then open http://127.0.0.1:8765 in your browser.

Add new scripts in registry.py.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from dataclasses import asdict
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from starlette.datastructures import UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel

import registry

HERE = Path(__file__).resolve().parent
templates = Jinja2Templates(directory=str(HERE / "templates"))

app = FastAPI(title="Electronics Engineering Assistant", docs_url="/api/docs", redoc_url=None)
app.mount("/static", StaticFiles(directory=str(HERE / "static")), name="static")


def _serializable_functions() -> list[dict]:
    out: list[dict] = []
    for fn in registry.FUNCTIONS:
        if getattr(fn, "hidden", False):
            continue
        out.append(
            {
                "id": fn.id,
                "title": fn.title,
                "description": fn.description,
                "background": fn.background,
                "group": fn.group,
                "icon": fn.icon,
                "has_file_field": fn.has_file_field,
                "has_loadable_select": fn.has_loadable_select,
                "has_discovery": fn.discovery is not None,
                "fields": [asdict(f) for f in fn.fields],
                "proceed_after_fail": (
                    asdict(fn.proceed_after_fail) if fn.proceed_after_fail else None
                ),
                "duplicate_upload_prompt": fn.duplicate_upload_prompt,
                "zip_overwrite_prompt": fn.zip_overwrite_prompt,
                "confirm_upload_prompt": fn.confirm_upload_prompt,
                "missing_schematic_prompt": fn.missing_schematic_prompt,
                "report_view_actions": fn.report_view_actions,
                "discovery_prefill": fn.discovery_prefill,
                "checkboxes_on_subrow": fn.checkboxes_on_subrow,
                "stack_text_select_fields": registry.stack_text_select_fields(fn),
                "run_actions": [
                    {"id": a.id, "label": a.label} for a in fn.run_actions
                ],
            }
        )
    return out


_GROUP_ORDER = ("Altium", "Arena PLM", "Rarely Used", "Altium MCP")
_GROUP_DISPLAY = {
    "Altium": "Altium Scripts",
    "Arena PLM": "Arena Scripts",
    "Rarely Used": "Rarely Used Scripts",
    "Altium MCP": "Altium MCP Scripts",
}


def _grouped_functions() -> list[tuple[str, list[dict]]]:
    grouped: dict[str, list[dict]] = {}
    for fn in _serializable_functions():
        grouped.setdefault(fn["group"], []).append(fn)
    ordered: list[tuple[str, list[dict]]] = []
    for key in _GROUP_ORDER:
        if key in grouped:
            fns = registry.sort_launcher_group(key, grouped.pop(key))
            ordered.append((_GROUP_DISPLAY.get(key, key), fns))
    for key, fns in grouped.items():
        ordered.append((_GROUP_DISPLAY.get(key, key), registry.sort_launcher_group(key, fns)))
    return ordered


def _arena_status() -> dict:
    # Workspace is set by the topbar dropdown, so we only check credentials here.
    return {
        "ARENA_EMAIL": bool(os.environ.get("ARENA_EMAIL")),
        "ARENA_PASSWORD": bool(os.environ.get("ARENA_PASSWORD")),
    }


def _workspaces() -> list[dict]:
    return [
        {"id": w.id, "label": w.label, "workspace_id": w.workspace_id}
        for w in registry.WORKSPACES
    ]


def _resolve_workspace_id(workspace_key: str | None) -> str | None:
    """Map a workspace selector id to its numeric workspace_id."""
    if not workspace_key:
        return None
    for w in registry.WORKSPACES:
        if w.id == workspace_key:
            return w.workspace_id
    return None


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "groups": _grouped_functions(),
            "arena_status": _arena_status(),
            "workspaces": _workspaces(),
            "default_workspace": registry.DEFAULT_WORKSPACE_ID,
        },
    )


class RunRequest(BaseModel):
    inputs: dict[str, str] = {}
    workspace: str | None = None


def _missing_required(spec: registry.FunctionSpec, values: dict[str, str]) -> list[str]:
    missing: list[str] = []
    for field in spec.fields:
        if field.required and not (values.get(field.name) or "").strip():
            missing.append(field.label)
    return missing


def _execute(
    spec: registry.FunctionSpec,
    values: dict[str, str],
    workspace_key: str | None,
    cleanup_paths: list[Path] | None = None,
) -> JSONResponse:
    """Validate, build the command, and either run or background it."""
    missing = _missing_required(spec, values)
    if missing:
        _cleanup(cleanup_paths)
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error": "Missing required input(s): " + ", ".join(missing),
            },
        )

    try:
        argv = spec.build_command(values)
    except KeyError as e:
        _cleanup(cleanup_paths)
        return JSONResponse(
            status_code=400,
            content={"ok": False, "error": f"Missing input: {e.args[0]}"},
        )

    cwd = str(spec.cwd) if spec.cwd else None
    display_cmd = " ".join(_quote_for_display(a) for a in argv)

    env = dict(os.environ)
    workspace_id = _resolve_workspace_id(workspace_key)
    if workspace_id:
        env["ARENA_WORKSPACE_ID"] = workspace_id

    if spec.background:
        try:
            creationflags = 0
            if hasattr(subprocess, "DETACHED_PROCESS"):
                creationflags |= subprocess.DETACHED_PROCESS
            if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
                creationflags |= subprocess.CREATE_NEW_PROCESS_GROUP
            subprocess.Popen(
                argv,
                cwd=cwd,
                env=env,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                stdin=subprocess.DEVNULL,
                close_fds=True,
                creationflags=creationflags,
            )
        except Exception as e:
            _cleanup(cleanup_paths)
            return JSONResponse(
                status_code=500,
                content={"ok": False, "error": str(e), "command": display_cmd},
            )
        # Background processes own their tempfile lifetime; we cannot safely delete
        # the file out from under them, so leave it for the OS temp sweep.
        return JSONResponse(
            content={
                "ok": True,
                "background": True,
                "command": display_cmd,
                "stdout": "Launched in background.",
                "stderr": "",
                "exit_code": None,
            }
        )

    try:
        result = subprocess.run(
            argv,
            cwd=cwd,
            env=env,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=300,
        )
    except subprocess.TimeoutExpired:
        _cleanup(cleanup_paths)
        return JSONResponse(
            status_code=504,
            content={
                "ok": False,
                "error": "Timed out after 300 seconds.",
                "command": display_cmd,
            },
        )
    except Exception as e:
        _cleanup(cleanup_paths)
        return JSONResponse(
            status_code=500,
            content={"ok": False, "error": str(e), "command": display_cmd},
        )

    _cleanup(cleanup_paths)
    return JSONResponse(
        content={
            "ok": result.returncode == 0,
            "background": False,
            "command": display_cmd,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "exit_code": result.returncode,
        }
    )


def _cleanup(paths: list[Path] | None) -> None:
    if not paths:
        return
    for p in paths:
        try:
            p.unlink(missing_ok=True)
            parent = p.parent
            if parent.exists() and parent.name.startswith("launcher_upload_"):
                shutil.rmtree(parent, ignore_errors=True)
        except Exception:
            pass


class DiscoverRequest(BaseModel):
    inputs: dict[str, str] = {}
    workspace: str | None = None


@app.post("/api/discover/{function_id}")
def discover_function(function_id: str, body: DiscoverRequest) -> JSONResponse:
    """Run a card's discovery script and return its parsed JSON output."""
    spec = registry.find(function_id)
    if spec is None:
        raise HTTPException(status_code=404, detail=f"Unknown function: {function_id}")
    if spec.discovery is None:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "error": "This function has no discovery."},
        )
    values = {k: (v or "").strip() for k, v in body.inputs.items()}
    try:
        argv = spec.discovery.build_command(values)
    except KeyError as e:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "error": f"Missing input: {e.args[0]}"},
        )

    env = dict(os.environ)
    workspace_id = _resolve_workspace_id(body.workspace)
    if workspace_id:
        env["ARENA_WORKSPACE_ID"] = workspace_id

    try:
        result = subprocess.run(
            argv, env=env, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=60,
        )
    except subprocess.TimeoutExpired:
        return JSONResponse(
            status_code=504,
            content={"ok": False, "error": "Discovery timed out after 60s."},
        )
    if result.returncode != 0:
        return JSONResponse(
            status_code=500,
            content={
                "ok": False,
                "error": (result.stderr or result.stdout or "Discovery script failed.").strip(),
                "exit_code": result.returncode,
            },
        )

    import json as _json
    try:
        parsed = _json.loads(result.stdout)
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "ok": False,
                "error": f"Discovery output was not valid JSON: {e}",
                "stdout": result.stdout[:2000],
            },
        )
    return JSONResponse(content={"ok": True, "data": parsed})


@app.post("/api/run/{function_id}")
def run_function(function_id: str, body: RunRequest) -> JSONResponse:
    spec = registry.find(function_id)
    if spec is None:
        raise HTTPException(status_code=404, detail=f"Unknown function: {function_id}")
    if spec.has_file_field:
        return JSONResponse(
            status_code=400,
            content={
                "ok": False,
                "error": (
                    "This function expects file uploads. POST as multipart/form-data "
                    f"to /api/upload/{function_id} instead."
                ),
            },
        )
    values = {k: (v or "").strip() for k, v in body.inputs.items()}
    return _execute(spec, values, body.workspace)


@app.post("/api/upload/{function_id}")
async def upload_function(function_id: str, request: Request) -> JSONResponse:
    """Multipart entry point for cards that include file fields.

    Each non-file form field is read as a string. Each file field is streamed
    to a per-request temp directory, and its absolute path is substituted in as
    the field's value before build_command runs. Temp files are cleaned up
    after the (foreground) script completes.
    """
    spec = registry.find(function_id)
    if spec is None:
        raise HTTPException(status_code=404, detail=f"Unknown function: {function_id}")

    form = await request.form()
    workspace_key = form.get("__workspace") or None
    file_field_names = {f.name for f in spec.fields if f.kind == "file"}

    values: dict[str, str] = {}
    cleanup_paths: list[Path] = []
    upload_dir: Path | None = None

    for key, val in form.multi_items():
        if key == "__workspace":
            continue
        if key in file_field_names:
            # Duck-type for UploadFile: anything with a .file attr and .filename.
            filename = getattr(val, "filename", None)
            if not filename:
                continue
            if upload_dir is None:
                upload_dir = Path(tempfile.mkdtemp(prefix="launcher_upload_"))
            safe_name = Path(filename).name or "upload.bin"
            dest = upload_dir / safe_name
            file_obj = getattr(val, "file", None)
            if file_obj is None:
                continue
            with dest.open("wb") as fh:
                shutil.copyfileobj(file_obj, fh)
            cleanup_paths.append(dest)
            values[key] = str(dest)
        else:
            values[key] = (val if isinstance(val, str) else "").strip()

    return _execute(spec, values, workspace_key, cleanup_paths=cleanup_paths)


def _quote_for_display(s: str) -> str:
    if any(c.isspace() for c in s) or any(c in s for c in "\"'`$&|<>;"):
        escaped = s.replace('"', '\\"')
        return f'"{escaped}"'
    return s


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8765, log_level="info")
