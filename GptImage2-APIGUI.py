from __future__ import annotations

import base64
import json
import logging
import os
import sys
import threading
import urllib.request
from email.message import Message
from urllib.parse import urlparse, urlunparse
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from tkinter import (
    END,
    BOTH,
    LEFT,
    X,
    RIGHT,
    Button,
    Canvas,
    Checkbutton,
    Entry,
    Frame,
    IntVar,
    Label,
    LabelFrame,
    Listbox,
    StringVar,
    Text,
    Tk,
    Toplevel,
)
from tkinter import filedialog, messagebox, simpledialog, ttk
from uuid import uuid4

from openai import OpenAI


DEFAULT_CONFIG = {
    "api_key": "",
    "base_url": "",
    "model": "",
    # Default output dir (optional). Empty => Pic_generate/image
    "out_dir": "",
    "ui_language": "zh",
    "connection_presets_json": "{}",
    "active_preset": "",
}

TEXTS: dict[str, dict[str, str]] = {
    "zh": {
        "window_title": "GPT 图片 GUI（gpt-image-2）",
        "api_key": "API 密钥 (OPENAI_API_KEY):",
        "base_url": "Base URL (OPENAI_BASE_URL):",
        "model": "模型：",
        "quality": "质量：",
        "size": "尺寸：",
        "n": "数量：",
        "language": "语言：",
        "settings": "设置",
        "prompt_images": "提示词与输入图",
        "prompt": "提示词：",
        "selected_images": "原图（用于图像编辑）：",
        "preview": "预览",
        "install_pillow": "（安装 pillow 可显示缩略图）",
        "add_images": "添加图片...",
        "remove_selected": "移除选中",
        "clear": "清空",
        "preview_tip": "提示：双击图片可预览",
        "output_run": "输出与生成",
        "output": "输出目录：",
        "browse": "浏览...",
        "open_folder": "打开文件夹",
        "filename_prefix": "文件名前缀：",
        "open_after_save": "保存后打开文件夹",
        "record_session": "记录会话",
        "generate": "生成",
        "interrupt": "中断当前会话",
        "send_background": "转后台继续",
        "save_preset": "保存预设",
        "rename_preset": "重命名",
        "delete_preset": "删除预设",
        "preset": "预设：",
        "preset_name_prompt": "预设名称：",
        "preset_name_empty": "预设名称不能为空。",
        "preset_rename_prompt": "新预设名称：",
        "preset_name_exists": "预设已存在：{name}",
        "preset_saved": "已保存预设：{name}",
        "preset_deleted": "已删除预设：{name}",
        "preset_not_found": "找不到预设：{name}",
        "refresh_models": "获取模型",
        "fetch_models_failed": "获取模型失败：{msg}",
        "fetching_models": "正在从服务端获取模型列表...",
        "session_cancelled": "会话已中断（后台请求可能仍在收尾）",
        "session_moved_bg": "会话已转后台，已创建可继续编辑的新会话。",
        "sessions": "会话管理",
        "new": "新建",
        "duplicate": "复制",
        "rename": "重命名",
        "delete": "删除",
        "latest_result": "最新结果：",
        "no_result_preview": "暂无结果预览。",
        "open_latest_result": "打开最新结果",
        "open_result_folder": "打开结果目录",
        "tools": "工具",
        "load_session": "加载会话...",
        "load_all_sessions": "加载全部历史",
        "open_logs": "打开日志",
        "help": "帮助",
        "ready": "就绪。",
        "queued": "已排队...",
        "generating_bg": "正在生成...（后台运行中）",
        "waiting_provider": "等待服务端响应...",
        "processing_image": "正在处理图片 {current}/{total}",
        "downloading_image": "正在下载图片 {current}/{total}",
        "done_saved": "完成。已保存：{files}{hint}",
        "error_with_log": "错误：{msg}（日志：{log}）",
        "loading_thumbnails": "正在加载缩略图...",
        "session_loaded": "已加载会话。",
        "session_imported_results": "已导入会话，包含 {count} 个生成结果。",
        "history_loaded": "已加载 {count} 个历史会话。",
        "rename_session": "重命名会话",
        "session_title": "会话标题：",
        "session_title_empty": "会话标题不能为空。",
        "delete_running_session": "正在运行的会话无法删除。",
        "delete_last_session": "至少需要保留一个会话。",
        "no_generated_result": "当前会话还没有生成结果。",
        "file_not_found": "文件不存在：\n{path}",
        "select_images": "选择图片",
        "choose_output_folder": "选择输出目录",
        "load_session_file": "加载会话文件",
        "invalid_session_format": "无效的会话格式。",
        "session_missing_params": "会话缺少 params 对象。",
        "load_session_failed": "加载会话失败",
        "preview_failed": "预览失败",
        "invalid_input": "输入无效",
        "generate_failed": "生成失败",
        "already_running": "该会话已经在运行中。",
        "prompt_empty": "提示词为空。",
        "missing_api_key": "缺少 API key（请设置 OPENAI_API_KEY 或在界面中输入）。",
        "help_title": "帮助",
        "help_body": "使用方式：\n- 填写提示词。\n- （可选）添加原图：若选择了图片，则使用 images.edit（图生图）。\n- 未选择图片时，则使用 images.generate（文生图）。\n- 可通过右侧会话区新建/切换多个会话并行生成。\n",
        "open_in_viewer": "在系统查看器中打开",
        "generated_result_title": "生成结果 - {name}",
        "preview_title": "预览 - {name}",
        "session_summary": "标题：{title}\n状态：{state}\n详情：{detail}\n图片：{inputs} 张输入 / {outputs} 张输出",
        "last_error": "最近错误：{error}",
        "status_idle": "空闲",
        "status_running": "运行中",
        "status_processing": "处理中",
        "status_downloading": "下载中",
        "status_done": "已完成",
        "status_error": "错误",
        "status_unknown": "未知",
        "status_cancelled": "已中断",
    },
    "en": {
        "window_title": "GPT Image GUI (gpt-image-2)",
        "api_key": "API Key (OPENAI_API_KEY):",
        "base_url": "Base URL (OPENAI_BASE_URL):",
        "model": "Model:",
        "quality": "Quality:",
        "size": "Size:",
        "n": "N:",
        "language": "Language:",
        "settings": "Settings",
        "prompt_images": "Prompt & Images",
        "prompt": "Prompt:",
        "selected_images": "Source images (for image edit):",
        "preview": "Preview",
        "install_pillow": "(Install pillow for thumbnails)",
        "add_images": "Add images...",
        "remove_selected": "Remove selected",
        "clear": "Clear",
        "preview_tip": "Tip: double-click an image to preview",
        "output_run": "Output & Run",
        "output": "Output:",
        "browse": "Browse...",
        "open_folder": "Open folder",
        "filename_prefix": "Filename prefix:",
        "open_after_save": "Open folder after save",
        "record_session": "Record session",
        "generate": "Generate",
        "interrupt": "Interrupt Session",
        "send_background": "Send to Background",
        "save_preset": "Save preset",
        "rename_preset": "Rename preset",
        "delete_preset": "Delete preset",
        "preset": "Preset:",
        "preset_name_prompt": "Preset name:",
        "preset_name_empty": "Preset name cannot be empty.",
        "preset_rename_prompt": "New preset name:",
        "preset_name_exists": "Preset already exists: {name}",
        "preset_saved": "Preset saved: {name}",
        "preset_deleted": "Preset deleted: {name}",
        "preset_not_found": "Preset not found: {name}",
        "refresh_models": "Fetch models",
        "fetch_models_failed": "Fetch models failed: {msg}",
        "fetching_models": "Fetching model list from provider...",
        "session_cancelled": "Session interrupted (provider request may still be finalizing in background).",
        "session_moved_bg": "Session moved to background; a new editable session is ready.",
        "sessions": "Sessions",
        "new": "New",
        "duplicate": "Duplicate",
        "rename": "Rename",
        "delete": "Delete",
        "latest_result": "Latest result:",
        "no_result_preview": "No result preview yet.",
        "open_latest_result": "Open latest result",
        "open_result_folder": "Open result folder",
        "tools": "Tools",
        "load_session": "Load session...",
        "load_all_sessions": "Load all history",
        "open_logs": "Open logs",
        "help": "Help",
        "ready": "Ready.",
        "queued": "Queued...",
        "generating_bg": "Generating... (running in background)",
        "waiting_provider": "Waiting for provider response...",
        "processing_image": "Processing image {current}/{total}",
        "downloading_image": "Downloading image {current}/{total}",
        "done_saved": "Done. Saved: {files}{hint}",
        "error_with_log": "Error: {msg} (log: {log})",
        "loading_thumbnails": "Loading thumbnails...",
        "session_loaded": "Session loaded.",
        "session_imported_results": "Imported session with {count} generated file(s).",
        "history_loaded": "Loaded {count} historical session(s).",
        "rename_session": "Rename session",
        "session_title": "Session title:",
        "session_title_empty": "Session title cannot be empty.",
        "delete_running_session": "Cannot delete a session while it is running.",
        "delete_last_session": "At least one session must remain.",
        "no_generated_result": "This session has no generated result yet.",
        "file_not_found": "File not found:\n{path}",
        "select_images": "Select images",
        "choose_output_folder": "Choose output folder",
        "load_session_file": "Load session file",
        "invalid_session_format": "Invalid session format.",
        "session_missing_params": "Session missing params object.",
        "load_session_failed": "Load session failed",
        "preview_failed": "Preview failed",
        "invalid_input": "Invalid input",
        "generate_failed": "Generate failed",
        "already_running": "This session is already running.",
        "prompt_empty": "Prompt is empty.",
        "missing_api_key": "Missing API key (set OPENAI_API_KEY or paste it here).",
        "help_title": "Help",
        "help_body": "Usage:\n- Fill prompt.\n- (Optional) Add source images: if images are selected, the app uses images.edit.\n- If no images are selected, the app uses images.generate.\n- Use the session area on the right to create/switch sessions and run tasks in parallel.\n",
        "open_in_viewer": "Open in system viewer",
        "generated_result_title": "Generated Result - {name}",
        "preview_title": "Preview - {name}",
        "session_summary": "Title: {title}\nState: {state}\nDetail: {detail}\nImages: {inputs} input / {outputs} output",
        "last_error": "Last error: {error}",
        "status_idle": "idle",
        "status_running": "running",
        "status_processing": "processing",
        "status_downloading": "downloading",
        "status_done": "done",
        "status_error": "error",
        "status_unknown": "unknown",
        "status_cancelled": "cancelled",
    },
}


def _app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _config_file_path() -> Path:
    return _app_base_dir() / "config.json"


def _sessions_dir() -> Path:
    return _app_base_dir() / "sessions"


def _logs_dir() -> Path:
    return _app_base_dir() / "logs"


def _log_file_path() -> Path:
    return _logs_dir() / "app.log"


def _setup_logger() -> logging.Logger:
    _logs_dir().mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("chatanywhere_gui")
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    handler = logging.FileHandler(_log_file_path(), encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(handler)
    logger.propagate = False
    return logger


def _load_config() -> dict[str, str]:
    cfg_path = _config_file_path()
    if not cfg_path.exists():
        cfg_path.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2), encoding="utf-8")
        return dict(DEFAULT_CONFIG)

    try:
        raw = json.loads(cfg_path.read_text(encoding="utf-8"))
    except Exception:
        return dict(DEFAULT_CONFIG)

    if not isinstance(raw, dict):
        return dict(DEFAULT_CONFIG)

    merged = dict(DEFAULT_CONFIG)
    for k in DEFAULT_CONFIG:
        v = raw.get(k, "")
        merged[k] = str(v).strip() if v is not None else ""
    return merged


CONFIG = _load_config()
LOGGER = _setup_logger()


def _save_config_values(updates: dict[str, str]) -> None:
    CONFIG.update({k: str(v) for k, v in updates.items()})
    cfg_path = _config_file_path()
    payload = {k: CONFIG.get(k, DEFAULT_CONFIG.get(k, "")) for k in DEFAULT_CONFIG}
    cfg_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _first_non_empty(*vals: str) -> str:
    for v in vals:
        v = (v or "").strip()
        if v:
            return v
    return ""


def _open_folder(path: Path) -> None:
    path = path.resolve()
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        raise RuntimeError("Open folder is only implemented for Windows in this script.")


def _open_file(path: Path) -> None:
    path = path.resolve()
    if os.name == "nt":
        os.startfile(str(path))  # type: ignore[attr-defined]
    else:
        raise RuntimeError("Open file is only implemented for Windows in this script.")


def _safe_mkdir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def _next_available_path(path: Path) -> Path:
    if not path.exists():
        return path
    base = path.with_suffix("")
    ext = path.suffix or ".png"
    counter = 1
    while True:
        candidate = Path(f"{base}_{counter}{ext}")
        if not candidate.exists():
            return candidate
        counter += 1


def _default_filename(prefix: str, ext: str = ".png") -> str:
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    prefix = prefix.strip() or "image"
    return f"{prefix}_{ts}{ext}"


def _normalize_base_url(base_url: str) -> str:
    """
    Normalize a provider base_url to an OpenAI-compatible API root.

    Many providers require the base url to end with `/v1`. If the user pastes
    a bare host (e.g. https://api.example.com) it often returns an HTML page
    instead of JSON. This function appends `/v1` when it looks missing.
    """
    base_url = (base_url or "").strip()
    if not base_url:
        return base_url

    # Trim trailing slashes but preserve scheme/host.
    base_url = base_url.rstrip("/")

    parsed = urlparse(base_url)
    # If user entered only host without scheme, keep as-is and let SDK error.
    if not parsed.scheme or not parsed.netloc:
        return base_url

    path = (parsed.path or "").rstrip("/")
    if path.endswith("/v1") or path == "/v1":
        new_path = path
    else:
        new_path = (path + "/v1") if path else "/v1"

    return urlunparse((parsed.scheme, parsed.netloc, new_path, "", "", ""))


def _download_url_bytes(url: str, timeout_s: int = 60) -> tuple[bytes, str | None]:
    url = (url or "").strip()
    if not url:
        raise RuntimeError("Empty url in provider response item.")
    req = urllib.request.Request(url, headers={"User-Agent": "GptImage2-APIGUI/1.0"})
    with urllib.request.urlopen(req, timeout=timeout_s) as resp:  # nosec - URL comes from provider response
        body = resp.read()
        headers: Message | None = getattr(resp, "headers", None)
        content_type = headers.get_content_type() if headers else None
        return body, content_type


def _b64decode_relaxed(data: str) -> bytes:
    """
    Decode base64 that may be missing padding or contain whitespace.
    Some providers return non-padded base64 for image bytes.
    """
    s = (data or "").strip()
    if not s:
        raise ValueError("Empty base64 data")
    s = "".join(s.split())
    # Fix missing padding
    pad = (-len(s)) % 4
    if pad:
        s = s + ("=" * pad)
    return base64.b64decode(s)


def _split_data_url(data: str) -> tuple[str | None, str]:
    """
    Split a data URL like:
      data:image/png;base64,AAAA...
    Returns (mime, payload_string). If not a data URL, mime is None and payload is original.
    """
    s = (data or "").strip()
    if not s.lower().startswith("data:"):
        return None, s

    # Minimal data URL parsing: data:[<mime>][;base64],<payload>
    comma = s.find(",")
    if comma < 0:
        # Malformed, treat as raw string
        return None, s
    meta = s[5:comma]  # after "data:"
    payload = s[comma + 1 :]
    mime = None
    if meta:
        mime = meta.split(";")[0].strip() or None
    return mime, payload


def _mime_to_ext(mime: str | None) -> str | None:
    if not mime:
        return None
    m = mime.lower().strip()
    if m == "image/png":
        return ".png"
    if m in ("image/jpeg", "image/jpg"):
        return ".jpg"
    if m == "image/gif":
        return ".gif"
    if m == "image/webp":
        return ".webp"
    return None


def _is_plausible_base64(data: str) -> bool:
    """
    Fast, conservative base64 sanity check to avoid expensive decoding on
    obviously-invalid payloads.
    """
    mime, s = _split_data_url(data)
    s = (s or "").strip()
    if not s:
        return False
    # Remove whitespace cheaply
    s = "".join(s.split())
    # If len % 4 == 1, base64 is impossible (even with padding).
    if (len(s) % 4) == 1:
        return False
    # Quick character check (allow padding '=' only at the end).
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=")
    # Only sample up to a point to keep this O(1) for huge strings.
    sample = s[:2048]
    for ch in sample:
        if ch not in allowed:
            return False
    return True


def _guess_image_ext(data: bytes) -> str | None:
    """
    Guess image extension by magic bytes.
    Returns extension including dot (e.g. ".png") or None if unknown.
    """
    if not data or len(data) < 12:
        return None
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if data.startswith(b"\xff\xd8\xff"):
        return ".jpg"
    if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
        return ".gif"
    # WEBP: RIFF....WEBP
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return ".webp"
    return None


def _looks_like_html(data: bytes) -> bool:
    if not data:
        return False
    head = data[:200].lstrip().lower()
    return head.startswith(b"<!doctype html") or head.startswith(b"<html") or b"<head" in head


@dataclass
class GenerateParams:
    api_key: str
    base_url: str
    model: str
    prompt: str
    quality: str
    size: str
    n: int
    images: list[Path]
    out_dir: Path
    filename_prefix: str


@dataclass
class SessionState:
    session_id: str
    title: str
    api_key: str
    base_url: str
    model: str
    prompt: str
    quality: str
    size: str
    n: int
    images: list[Path]
    out_dir: Path
    filename_prefix: str
    open_out_after: bool = False
    record_session: bool = True
    status_state: str = "idle"
    status_detail: str = ""
    generated_files: list[str] | None = None
    last_error: str = ""
    last_session_file: str = ""
    imported_from_path: str = ""
    running: bool = False
    current_run_id: int = 0
    run_count: int = 0
    selected_input_index: int | None = None
    cancelled_run_ids: set[int] = field(default_factory=set)

    def __post_init__(self) -> None:
        if self.generated_files is None:
            self.generated_files = []


class App:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.geometry("1360x920")
        self.ui_language_var = StringVar(value=(CONFIG.get("ui_language", "") or "zh"))
        self.root.title(self._t("window_title"))

        self.api_key_var = StringVar(value=_first_non_empty(CONFIG.get("api_key", ""), os.getenv("OPENAI_API_KEY", "")))
        self.base_url_var = StringVar(
            value=_first_non_empty(CONFIG.get("base_url", ""), os.getenv("OPENAI_BASE_URL", ""), "https://api.openai.com/v1")
        )
        self.model_var = StringVar(
            value=_first_non_empty(CONFIG.get("model", ""), os.getenv("OPENAI_IMAGE_MODEL", ""), "gpt-image-2")
        )
        self._model_choices: list[str] = []
        self._model_choices_full: list[str] = []
        self.quality_var = StringVar(value="high")
        self.size_var = StringVar(value="1024x1024")
        self.n_var = StringVar(value="1")
        default_out = _first_non_empty(CONFIG.get("out_dir", ""), str((Path("Pic_generate") / "image").resolve()))
        self.out_dir_var = StringVar(value=default_out)
        self.prefix_var = StringVar(value="gpt")
        self.status_var = StringVar(value=self._t("ready"))
        self.open_out_after_var = IntVar(value=0)
        self.record_session_var = IntVar(value=1)
        self.preset_var = StringVar(value="")
        self.connection_presets: dict[str, dict[str, str]] = self._load_connection_presets()
        self._active_preset_name = (CONFIG.get("active_preset", "") or "").strip()

        self._image_preview_photo = None
        self._thumb_photos: list[object] = []
        self._thumb_labels: list[Label] = []
        self._pillow_available = self._check_pillow()
        self._result_thumb_photo = None
        self._loading_session = False
        self._thumb_request_id = 0
        self._thumb_session_id: str | None = None
        self.sessions: list[SessionState] = []
        self.active_session_id: str | None = None
        self._model_popup: Toplevel | None = None
        self._model_popup_list: Listbox | None = None
        self._model_popup_after_id: str | None = None

        self._build_ui()
        self._create_initial_session()
        self._auto_load_recent_sessions()
        if self._active_preset_name and self._active_preset_name in self.connection_presets:
            self._apply_preset(self._active_preset_name)

    def _load_connection_presets(self) -> dict[str, dict[str, str]]:
        raw = (CONFIG.get("connection_presets_json", "") or "{}").strip()
        try:
            parsed = json.loads(raw) if raw else {}
        except Exception:
            parsed = {}
        presets: dict[str, dict[str, str]] = {}
        if isinstance(parsed, dict):
            for name, value in parsed.items():
                name_str = str(name).strip()
                if not name_str or not isinstance(value, dict):
                    continue
                presets[name_str] = {
                    "api_key": str(value.get("api_key", "")).strip(),
                    "base_url": _normalize_base_url(str(value.get("base_url", "")).strip()),
                    "model": str(value.get("model", "")).strip() or "gpt-image-2",
                }
                models = value.get("models", [])
                if isinstance(models, list):
                    cleaned: list[str] = []
                    for m in models:
                        ms = str(m).strip()
                        if ms:
                            cleaned.append(ms)
                    if cleaned:
                        presets[name_str]["models"] = cleaned
        return presets

    def _persist_connection_presets(self) -> None:
        _save_config_values(
            {
                "connection_presets_json": json.dumps(self.connection_presets, ensure_ascii=False),
                "active_preset": self._active_preset_name,
            }
        )

    def _persist_form_defaults(self) -> None:
        _save_config_values(
            {
                "api_key": self.api_key_var.get().strip(),
                "base_url": _normalize_base_url(self.base_url_var.get().strip()),
                "model": self.model_var.get().strip() or "gpt-image-2",
                "out_dir": self.out_dir_var.get().strip(),
            }
        )

    def _lang(self) -> str:
        lang = self.ui_language_var.get().strip().lower()
        return lang if lang in TEXTS else "zh"

    def _t(self, key: str, **kwargs: object) -> str:
        template = TEXTS.get(self._lang(), TEXTS["zh"]).get(key, key)
        return template.format(**kwargs)

    def _status_state_label(self, state: str) -> str:
        return self._t(
            {
                "idle": "status_idle",
                "running": "status_running",
                "processing": "status_processing",
                "downloading": "status_downloading",
                "done": "status_done",
                "error": "status_error",
                "cancelled": "status_cancelled",
            }.get(state, "status_unknown")
        )

    def _apply_language_change(self, _event=None) -> None:
        _save_config_values({"ui_language": self._lang()})
        active = self._get_active_session()
        if active:
            self._save_form_into_active_session()
        self._rebuild_ui()

    def _rebuild_ui(self) -> None:
        for child in list(self.root.winfo_children()):
            child.destroy()
        self._build_ui()
        self.root.title(self._t("window_title"))
        active = self._get_active_session()
        if active:
            self._load_session_into_form(active)
            self._refresh_sessions_list()
        else:
            self.status_var.set(self._t("ready"))

    def _check_pillow(self) -> bool:
        try:
            import PIL  # noqa: F401

            return True
        except Exception:
            return False

    def _build_ui(self) -> None:
        top = LabelFrame(self.root, text=self._t("settings"))
        top.pack(fill=X, padx=8, pady=8)

        def row(parent: Frame) -> Frame:
            f = Frame(parent)
            f.pack(fill=BOTH, pady=3)
            return f

        r = row(top)
        Label(r, text=self._t("api_key"), width=24, anchor="w").pack(side=LEFT)
        Entry(r, textvariable=self.api_key_var, show="*", width=80).pack(side=LEFT, fill=BOTH, expand=True)
        Label(r, text=self._t("language"), width=8, anchor="e").pack(side=LEFT, padx=(12, 0))
        lang_box = ttk.Combobox(r, textvariable=self.ui_language_var, values=["zh", "en"], width=8, state="readonly")
        lang_box.pack(side=LEFT)
        lang_box.bind("<<ComboboxSelected>>", self._apply_language_change)

        r = row(top)
        Label(r, text=self._t("base_url"), width=24, anchor="w").pack(side=LEFT)
        Entry(r, textvariable=self.base_url_var, width=80).pack(side=LEFT, fill=BOTH, expand=True)

        r = row(top)
        Label(r, text=self._t("model"), width=24, anchor="w").pack(side=LEFT)
        self.model_box = ttk.Combobox(r, textvariable=self.model_var, values=self._model_choices, width=30, state="normal")
        self.model_box.pack(side=LEFT)
        self.model_box.bind("<KeyRelease>", self.on_model_search_key)
        self.model_box.bind("<FocusOut>", self.on_model_focus_out)
        self.model_box.bind("<Escape>", self.on_model_escape)
        self.model_box.bind("<Down>", self.on_model_down)
        self.model_box.bind("<Up>", self.on_model_up)
        self.model_box.bind("<Return>", self.on_model_enter)
        Button(r, text=self._t("refresh_models"), command=self.on_refresh_models).pack(side=LEFT, padx=(6, 0))
        Label(r, text=self._t("quality"), width=10, anchor="w").pack(side=LEFT, padx=(12, 0))
        ttk.Combobox(r, textvariable=self.quality_var, values=["low", "medium", "high", "auto"], width=10).pack(
            side=LEFT
        )
        Label(r, text=self._t("size"), width=6, anchor="w").pack(side=LEFT, padx=(12, 0))
        ttk.Combobox(
            r,
            textvariable=self.size_var,
            values=["512x512", "1024x1024", "1024x1536", "1536x1024"],
            width=12,
        ).pack(side=LEFT)
        Label(r, text=self._t("n"), width=5, anchor="w").pack(side=LEFT, padx=(12, 0))
        Entry(r, textvariable=self.n_var, width=5).pack(side=LEFT)

        r = row(top)
        Label(r, text=self._t("preset"), width=24, anchor="w").pack(side=LEFT)
        self.preset_box = ttk.Combobox(r, textvariable=self.preset_var, values=[], width=30, state="readonly")
        self.preset_box.pack(side=LEFT)
        self.preset_box.bind("<<ComboboxSelected>>", self.on_preset_selected)
        Button(r, text=self._t("save_preset"), command=self.on_save_preset).pack(side=LEFT, padx=(6, 0))
        Button(r, text=self._t("rename_preset"), command=self.on_rename_preset).pack(side=LEFT, padx=(6, 0))
        Button(r, text=self._t("delete_preset"), command=self.on_delete_preset).pack(side=LEFT, padx=(6, 0))
        self._refresh_preset_selector()

        mid = Frame(self.root)
        mid.pack(fill=BOTH, expand=True, padx=8, pady=(0, 8))

        left = Frame(mid)
        left.pack(side=LEFT, fill=BOTH, expand=True)
        right = Frame(mid, width=340)
        right.pack(side=RIGHT, fill=BOTH, expand=False, padx=(8, 0))

        prompt_group = LabelFrame(left, text=self._t("prompt_images"))
        prompt_group.pack(fill=BOTH, expand=True)

        Label(prompt_group, text=self._t("prompt")).pack(anchor="w")
        self.prompt_text = Text(prompt_group, height=8)
        self.prompt_text.pack(fill=BOTH, expand=False)

        Label(prompt_group, text=self._t("selected_images")).pack(anchor="w", pady=(10, 0))
        images_area = Frame(prompt_group)
        images_area.pack(fill=BOTH, expand=True)

        self.images_list = Listbox(images_area, height=10)
        self.images_list.pack(side=LEFT, fill=BOTH, expand=True)

        thumbs = Frame(images_area, width=120)
        thumbs.pack(side=RIGHT, fill=BOTH, padx=(8, 0))
        Label(thumbs, text=self._t("preview")).pack(anchor="w")
        if not self._pillow_available:
            Label(thumbs, text=self._t("install_pillow"), fg="gray").pack(anchor="w")
        self.thumbs_canvas = Canvas(thumbs, width=120, highlightthickness=0)
        self.thumbs_canvas.pack(side=LEFT, fill=BOTH, expand=True)
        self.thumbs_scrollbar = ttk.Scrollbar(thumbs, orient="vertical", command=self.thumbs_canvas.yview)
        self.thumbs_scrollbar.pack(side=RIGHT, fill="y")
        self.thumbs_canvas.configure(yscrollcommand=self.thumbs_scrollbar.set)
        self.thumbs_frame = Frame(self.thumbs_canvas)
        self._thumbs_window = self.thumbs_canvas.create_window((0, 0), window=self.thumbs_frame, anchor="nw")

        def _on_thumbs_frame_configure(_event=None) -> None:
            self.thumbs_canvas.configure(scrollregion=self.thumbs_canvas.bbox("all"))

        def _on_thumbs_canvas_configure(event) -> None:
            self.thumbs_canvas.itemconfigure(self._thumbs_window, width=event.width)

        self.thumbs_frame.bind("<Configure>", _on_thumbs_frame_configure)
        self.thumbs_canvas.bind("<Configure>", _on_thumbs_canvas_configure)
        self.thumbs_canvas.bind("<MouseWheel>", lambda event: self.thumbs_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

        btns = Frame(prompt_group)
        btns.pack(fill=BOTH, pady=(4, 0))
        Button(btns, text=self._t("add_images"), command=self.on_add_images).pack(side=LEFT)
        Button(btns, text=self._t("remove_selected"), command=self.on_remove_image).pack(side=LEFT, padx=(6, 0))
        Button(btns, text=self._t("clear"), command=self.on_clear_images).pack(side=LEFT, padx=(6, 0))
        Label(prompt_group, text=self._t("preview_tip")).pack(anchor="w", pady=(4, 0))
        self.images_list.bind("<Double-Button-1>", self.on_preview_selected_image)
        self.images_list.bind("<<ListboxSelect>>", self.on_list_selected)
        output_group = LabelFrame(right, text=self._t("output_run"))
        output_group.pack(fill=X)

        r = Frame(output_group)
        r.pack(fill=BOTH, pady=4)
        Label(r, text=self._t("output"), width=12, anchor="w").pack(side=LEFT)
        Entry(r, textvariable=self.out_dir_var, width=50).pack(side=LEFT, fill=BOTH, expand=True)
        Button(r, text=self._t("browse"), command=self.on_pick_out_dir).pack(side=LEFT, padx=(6, 0))

        r = Frame(output_group)
        r.pack(fill=BOTH, pady=4)
        Label(r, text=self._t("filename_prefix"), width=16, anchor="w").pack(side=LEFT)
        Entry(r, textvariable=self.prefix_var, width=30).pack(side=LEFT)
        Checkbutton(r, text=self._t("open_after_save"), variable=self.open_out_after_var).pack(side=LEFT, padx=(10, 0))

        r = Frame(output_group)
        r.pack(fill=BOTH, pady=4)
        Checkbutton(r, text=self._t("record_session"), variable=self.record_session_var).pack(side=LEFT)
        Button(r, text=self._t("open_folder"), command=self.on_open_out_dir).pack(side=LEFT, padx=(8, 0))

        run = Frame(output_group)
        run.pack(fill=BOTH, pady=(8, 0))
        Button(run, text=self._t("generate"), command=self.on_generate_clicked, height=2, width=18).pack(side=LEFT)
        Button(run, text=self._t("interrupt"), command=self.on_interrupt_session, height=2).pack(side=LEFT, padx=(6, 0))
        Button(run, text=self._t("send_background"), command=self.on_send_background, height=2).pack(side=LEFT, padx=(6, 0))

        tools_group = LabelFrame(right, text=self._t("tools"))
        tools_group.pack(fill=X, pady=(8, 0))
        tools = Frame(tools_group)
        tools.pack(fill=BOTH, pady=4)
        Button(tools, text=self._t("load_session"), command=self.on_load_session).pack(side=LEFT)
        Button(tools, text=self._t("load_all_sessions"), command=self.on_load_all_sessions).pack(side=LEFT, padx=(6, 0))
        Button(tools, text=self._t("open_logs"), command=self.on_open_logs).pack(side=LEFT, padx=(6, 0))
        Button(tools, text=self._t("help"), command=self.on_help).pack(side=LEFT, padx=(6, 0))

        sessions_wrap = LabelFrame(right, text=self._t("sessions"))
        sessions_wrap.pack(fill=BOTH, expand=True, pady=(8, 0))

        session_actions = Frame(sessions_wrap)
        session_actions.pack(fill=BOTH, pady=(3, 4))
        Button(session_actions, text=self._t("new"), command=self.on_new_session).pack(side=LEFT)
        Button(session_actions, text=self._t("duplicate"), command=self.on_duplicate_session).pack(side=LEFT, padx=(6, 0))
        Button(session_actions, text=self._t("rename"), command=self.on_rename_session).pack(side=LEFT, padx=(6, 0))
        Button(session_actions, text=self._t("delete"), command=self.on_delete_session).pack(side=LEFT, padx=(6, 0))

        list_wrap = Frame(sessions_wrap)
        list_wrap.pack(fill=BOTH, expand=True)
        self.sessions_list = Listbox(list_wrap, height=10)
        self.sessions_list.pack(side=LEFT, fill=BOTH, expand=True)
        sessions_scroll = ttk.Scrollbar(list_wrap, orient="vertical", command=self.sessions_list.yview)
        sessions_scroll.pack(side=RIGHT, fill="y")
        self.sessions_list.configure(yscrollcommand=sessions_scroll.set)
        self.sessions_list.bind("<<ListboxSelect>>", self.on_session_selected)
        self.sessions_list.bind("<Double-Button-1>", self.on_rename_session)

        info_wrap = Frame(sessions_wrap)
        info_wrap.pack(fill=X, pady=(6, 4))

        info_text_wrap = Frame(info_wrap)
        info_text_wrap.pack(side=LEFT, fill=BOTH, expand=True)
        self.session_summary_var = StringVar(value=self._t("ready"))
        Label(
            info_text_wrap,
            textvariable=self.session_summary_var,
            justify="left",
            anchor="w",
            wraplength=205,
        ).pack(fill=BOTH)
        self.session_result_label = Label(info_text_wrap, text=self._t("no_result_preview"), fg="gray", anchor="w", justify="left")
        self.session_result_label.pack(anchor="w", pady=(4, 0))

        self.session_result_thumb_wrap = Frame(info_wrap, width=112, height=112)
        self.session_result_thumb_wrap.pack(side=RIGHT, padx=(8, 0))
        self.session_result_thumb_wrap.pack_propagate(False)
        self.session_result_thumb = Label(self.session_result_thumb_wrap, anchor="center")
        self.session_result_thumb.pack(fill=BOTH, expand=True)
        self.session_result_thumb.bind("<Double-Button-1>", self.on_open_latest_result)

        session_result_actions = Frame(sessions_wrap)
        session_result_actions.pack(fill=BOTH, pady=(4, 0))
        Button(session_result_actions, text=self._t("open_latest_result"), command=self.on_open_latest_result).pack(side=LEFT)
        Button(session_result_actions, text=self._t("open_result_folder"), command=self.on_open_session_result_folder).pack(
            side=LEFT, padx=(6, 0)
        )

        status = Frame(self.root)
        status.pack(fill=BOTH, padx=8, pady=(0, 8))
        Label(status, textvariable=self.status_var, anchor="w").pack(fill=BOTH)

    def on_help(self) -> None:
        win = Toplevel(self.root)
        win.title(self._t("help_title"))
        msg = Text(win, height=18, width=100)
        msg.pack(fill=BOTH, expand=True)
        msg.insert(END, self._t("help_body"))
        msg.configure(state="disabled")

    def _create_session(self, *, from_session: SessionState | None = None, title: str | None = None) -> SessionState:
        if from_session is None:
            session = SessionState(
                session_id=uuid4().hex,
                title=title or f"Session {len(self.sessions) + 1}",
                api_key=self.api_key_var.get().strip() or _first_non_empty(CONFIG.get("api_key", ""), os.getenv("OPENAI_API_KEY", "")),
                base_url=_normalize_base_url(
                    self.base_url_var.get().strip()
                    or _first_non_empty(CONFIG.get("base_url", ""), os.getenv("OPENAI_BASE_URL", ""), "https://api.openai.com/v1")
                ),
                model=self.model_var.get().strip() or _first_non_empty(CONFIG.get("model", ""), os.getenv("OPENAI_IMAGE_MODEL", ""), "gpt-image-2"),
                prompt=self.prompt_text.get("1.0", END).strip() if hasattr(self, "prompt_text") else "",
                quality=self.quality_var.get().strip().lower() or "high",
                size=self.size_var.get().strip() or "1024x1024",
                n=int(self.n_var.get().strip() or "1"),
                images=[],
                out_dir=Path(self.out_dir_var.get().strip() or (Path("Pic_generate") / "image")).resolve(),
                filename_prefix=self.prefix_var.get().strip() or "gpt",
                open_out_after=self.open_out_after_var.get() == 1,
                record_session=self.record_session_var.get() == 1,
                status_detail=self._t("ready"),
            )
        else:
            session = SessionState(
                session_id=uuid4().hex,
                title=title or f"{from_session.title} Copy",
                api_key=from_session.api_key,
                base_url=from_session.base_url,
                model=from_session.model,
                prompt=from_session.prompt,
                quality=from_session.quality,
                size=from_session.size,
                n=from_session.n,
                images=list(from_session.images),
                out_dir=Path(from_session.out_dir),
                filename_prefix=from_session.filename_prefix,
                open_out_after=from_session.open_out_after,
                record_session=from_session.record_session,
                generated_files=list(from_session.generated_files),
            )
        self.sessions.append(session)
        self._refresh_sessions_list()
        return session

    def _create_initial_session(self) -> None:
        session = self._create_session(title="Session 1")
        self._set_active_session(session.session_id)

    def _session_exists_for_path(self, path: Path) -> bool:
        target = str(path.resolve())
        for session in self.sessions:
            if session.imported_from_path and session.imported_from_path == target:
                return True
        return False

    def _import_session_file(self, path: Path, *, activate: bool) -> SessionState | None:
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(payload, dict):
                raise ValueError(self._t("invalid_session_format"))

            params = payload.get("params", {})
            if not isinstance(params, dict):
                raise ValueError(self._t("session_missing_params"))

            images = params.get("images", [])
            session = SessionState(
                session_id=uuid4().hex,
                title=path.stem,
                api_key=str(params.get("api_key", "")),
                base_url=_normalize_base_url(str(params.get("base_url", ""))),
                model=str(params.get("model", "gpt-image-2")),
                prompt=str(params.get("prompt", "")),
                quality=str(params.get("quality", "high")),
                size=str(params.get("size", "1024x1024")),
                n=int(str(params.get("n", "1")) or "1"),
                images=[],
                out_dir=Path(str(params.get("out_dir", self.out_dir_var.get()))).resolve(),
                filename_prefix=str(params.get("filename_prefix", "gpt")),
                generated_files=[str(x) for x in payload.get("generated_files", []) if str(x).strip()],
                status_state="done" if payload.get("generated_files") else "idle",
                status_detail=self._t("session_loaded"),
                imported_from_path=str(path.resolve()),
            )
            if isinstance(images, list):
                for p in images:
                    p_str = str(p).strip()
                    if p_str:
                        session.images.append(Path(p_str))
            self.sessions.append(session)
            if activate:
                self._set_active_session(session.session_id)
            else:
                self._refresh_sessions_list()
            if session.generated_files:
                session.status_detail = self._t("session_imported_results", count=len(session.generated_files))
            return session
        except Exception:
            LOGGER.exception("Load session failed: %s", path)
            raise

    def _load_saved_sessions(self, *, limit: int | None, activate_latest: bool) -> int:
        sessions_dir = _sessions_dir()
        if not sessions_dir.exists():
            return 0
        paths = sorted(sessions_dir.glob("session_*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
        if limit is not None:
            paths = paths[:limit]

        loaded = 0
        latest_loaded: SessionState | None = None
        for path in paths:
            if self._session_exists_for_path(path):
                continue
            session = self._import_session_file(path, activate=False)
            if session:
                loaded += 1
                latest_loaded = session
        if activate_latest and latest_loaded:
            self._set_active_session(latest_loaded.session_id)
        return loaded

    def _auto_load_recent_sessions(self) -> None:
        loaded = self._load_saved_sessions(limit=5, activate_latest=False)
        if loaded:
            active = self._get_active_session()
            if active:
                self.status_var.set(f"{active.title}: {self._t('history_loaded', count=loaded)}")

    def _get_session(self, session_id: str | None) -> SessionState | None:
        if not session_id:
            return None
        for session in self.sessions:
            if session.session_id == session_id:
                return session
        return None

    def _get_active_session(self) -> SessionState | None:
        return self._get_session(self.active_session_id)

    def _session_display_text(self, session: SessionState) -> str:
        marker = ">" if session.session_id == self.active_session_id else " "
        status_icon = {
            "idle": " ",
            "running": "~",
            "processing": "~",
            "downloading": "~",
            "done": "+",
            "error": "!",
            "cancelled": "x",
        }.get(session.status_state, "*")
        result_count = len(session.generated_files)
        suffix = f" | {result_count} result(s)" if result_count else ""
        detail = session.status_detail.strip()
        detail_suffix = ""
        if detail:
            detail_suffix = f" | {detail[:36]}{'...' if len(detail) > 36 else ''}"
        return f"{marker}{status_icon} {session.title} [{self._status_state_label(session.status_state)}]{suffix}{detail_suffix}"

    def _refresh_sessions_list(self) -> None:
        if not hasattr(self, "sessions_list"):
            return
        self.sessions_list.delete(0, END)
        for session in self.sessions:
            self.sessions_list.insert(END, self._session_display_text(session))
        active = self._get_active_session()
        if active:
            idx = self.sessions.index(active)
            self.sessions_list.selection_clear(0, END)
            self.sessions_list.selection_set(idx)
            self.sessions_list.see(idx)
        self._refresh_session_summary()

    def _refresh_session_summary(self) -> None:
        session = self._get_active_session()
        if not session:
            self.session_summary_var.set(self._t("ready"))
            self.session_result_label.configure(text=self._t("no_result_preview"), fg="gray")
            self.session_result_thumb.configure(image="", text="")
            self._result_thumb_photo = None
            return
        summary = self._t(
            "session_summary",
            title=session.title,
            state=self._status_state_label(session.status_state),
            detail=session.status_detail,
            inputs=len(session.images),
            outputs=len(session.generated_files),
        )
        if session.last_error:
            summary += "\n" + self._t("last_error", error=session.last_error)
        self.session_summary_var.set(summary)
        if session.generated_files:
            self.session_result_label.configure(text=f"{self._t('latest_result')} {Path(session.generated_files[-1]).name}", fg="black")
            self._refresh_result_thumbnail(session)
        else:
            self.session_result_label.configure(text=self._t("no_result_preview"), fg="gray")
            self.session_result_thumb.configure(image="", text="")
            self._result_thumb_photo = None

    def _refresh_result_thumbnail(self, session: SessionState) -> None:
        if not self._pillow_available or not session.generated_files:
            self.session_result_thumb.configure(image="", text="")
            self._result_thumb_photo = None
            return
        latest = Path(session.generated_files[-1])
        if not latest.exists():
            self.session_result_thumb.configure(image="", text=self._t("file_not_found", path=latest))
            self._result_thumb_photo = None
            return
        try:
            from PIL import Image, ImageTk  # type: ignore

            img = Image.open(latest)
            img.thumbnail((100, 100))
            tk_img = ImageTk.PhotoImage(img)
            self._result_thumb_photo = tk_img
            self.session_result_thumb.configure(image=tk_img, text="")
        except Exception:
            self.session_result_thumb.configure(image="", text=self._t("preview_failed"))
            self._result_thumb_photo = None

    def _load_session_into_form(self, session: SessionState) -> None:
        self._loading_session = True
        try:
            self.api_key_var.set(session.api_key)
            self.base_url_var.set(session.base_url)
            self.model_var.set(session.model)
            self.quality_var.set(session.quality)
            self.size_var.set(session.size)
            self.n_var.set(str(session.n))
            self.out_dir_var.set(str(session.out_dir))
            self.prefix_var.set(session.filename_prefix)
            self.open_out_after_var.set(1 if session.open_out_after else 0)
            self.record_session_var.set(1 if session.record_session else 0)
            self.prompt_text.delete("1.0", END)
            self.prompt_text.insert("1.0", session.prompt)
            self.images_list.delete(0, END)
            for path in session.images:
                self.images_list.insert(END, str(path))
            self._schedule_thumbnail_refresh(session.session_id)
            if session.status_detail:
                self.status_var.set(f"{session.title}: {session.status_detail}")
        finally:
            self._loading_session = False

    def _save_form_into_active_session(self) -> SessionState:
        session = self._get_active_session()
        if not session:
            raise RuntimeError("No active session.")
        if self._loading_session:
            return session
        session.api_key = self.api_key_var.get().strip()
        session.base_url = _normalize_base_url(self.base_url_var.get().strip())
        session.model = self.model_var.get().strip() or "gpt-image-2"
        session.quality = self.quality_var.get().strip().lower() or "high"
        session.size = self.size_var.get().strip() or "1024x1024"
        try:
            session.n = int(self.n_var.get().strip() or "1")
        except ValueError:
            pass
        session.out_dir = Path(self.out_dir_var.get().strip() or (Path("Pic_generate") / "image")).resolve()
        session.filename_prefix = self.prefix_var.get().strip() or "gpt"
        session.open_out_after = self.open_out_after_var.get() == 1
        session.record_session = self.record_session_var.get() == 1
        session.prompt = self.prompt_text.get("1.0", END).strip()
        session.selected_input_index = self._current_selected_input_index()
        self._refresh_sessions_list()
        return session

    def _set_active_session(self, session_id: str) -> None:
        current = self._get_active_session()
        if current and current.session_id == session_id:
            self._refresh_sessions_list()
            return
        if current and current.session_id != session_id:
            self._save_form_into_active_session()
        session = self._get_session(session_id)
        if not session:
            return
        self.active_session_id = session_id
        self._load_session_into_form(session)
        self._refresh_sessions_list()

    def _refresh_preset_selector(self) -> None:
        if not hasattr(self, "preset_box"):
            return
        names = sorted(self.connection_presets.keys())
        self.preset_box.configure(values=names)
        if self._active_preset_name in self.connection_presets:
            self.preset_var.set(self._active_preset_name)
        elif names:
            self.preset_var.set(names[0])
        else:
            self.preset_var.set("")

    def _apply_preset(self, name: str) -> bool:
        preset = self.connection_presets.get(name)
        if not preset:
            return False
        self.api_key_var.set(preset.get("api_key", ""))
        self.base_url_var.set(preset.get("base_url", ""))
        self.model_var.set(preset.get("model", "gpt-image-2"))
        models = preset.get("models", [])
        if isinstance(models, list):
            cleaned = [str(x).strip() for x in models if str(x).strip()]
            self._model_choices_full = list(dict.fromkeys(cleaned))
        else:
            self._model_choices_full = []
        self._model_choices = list(self._model_choices_full)
        if hasattr(self, "model_box"):
            self.model_box.configure(values=self._model_choices)
        self._active_preset_name = name
        self.preset_var.set(name)
        self._persist_connection_presets()
        self._save_form_into_active_session()
        return True

    def on_preset_selected(self, _event=None) -> None:
        name = self.preset_var.get().strip()
        if not name:
            return
        if not self._apply_preset(name):
            messagebox.showerror(self._t("preset"), self._t("preset_not_found", name=name))

    def on_save_preset(self) -> None:
        current_name = self.preset_var.get().strip()
        name = simpledialog.askstring(
            self._t("save_preset"),
            self._t("preset_name_prompt"),
            initialvalue=current_name,
            parent=self.root,
        )
        if name is None:
            return
        name = name.strip()
        if not name:
            messagebox.showerror(self._t("save_preset"), self._t("preset_name_empty"))
            return
        self.connection_presets[name] = {
            "api_key": self.api_key_var.get().strip(),
            "base_url": _normalize_base_url(self.base_url_var.get().strip()),
            "model": self.model_var.get().strip() or "gpt-image-2",
        }
        self._active_preset_name = name
        self._persist_connection_presets()
        self._refresh_preset_selector()
        self.status_var.set(self._t("preset_saved", name=name))
        self._save_form_into_active_session()

    def on_delete_preset(self) -> None:
        name = self.preset_var.get().strip()
        if not name:
            return
        if name not in self.connection_presets:
            messagebox.showerror(self._t("delete_preset"), self._t("preset_not_found", name=name))
            return
        del self.connection_presets[name]
        if self._active_preset_name == name:
            self._active_preset_name = ""
        self._persist_connection_presets()
        self._refresh_preset_selector()
        self.status_var.set(self._t("preset_deleted", name=name))

    def on_rename_preset(self) -> None:
        old = self.preset_var.get().strip()
        if not old:
            return
        if old not in self.connection_presets:
            messagebox.showerror(self._t("rename_preset"), self._t("preset_not_found", name=old))
            return
        new = simpledialog.askstring(
            self._t("rename_preset"),
            self._t("preset_rename_prompt"),
            initialvalue=old,
            parent=self.root,
        )
        if new is None:
            return
        new = new.strip()
        if not new:
            messagebox.showerror(self._t("rename_preset"), self._t("preset_name_empty"))
            return
        if new != old and new in self.connection_presets:
            messagebox.showerror(self._t("rename_preset"), self._t("preset_name_exists", name=new))
            return
        if new == old:
            return
        self.connection_presets[new] = dict(self.connection_presets[old])
        del self.connection_presets[old]
        if self._active_preset_name == old:
            self._active_preset_name = new
        self._persist_connection_presets()
        self._refresh_preset_selector()
        self._apply_preset(new)

    def on_refresh_models(self) -> None:
        api_key = self.api_key_var.get().strip()
        base_url = _normalize_base_url(self.base_url_var.get().strip())
        if not api_key:
            messagebox.showerror(self._t("invalid_input"), self._t("missing_api_key"))
            return
        if self._get_active_session():
            self.status_var.set(f"{self._get_active_session().title}: {self._t('fetching_models')}")
        else:
            self.status_var.set(self._t("fetching_models"))
        t = threading.Thread(target=self._fetch_models_worker, args=(api_key, base_url), daemon=True)
        t.start()

    def on_model_search_key(self, _event=None) -> None:
        # Debounced incremental search: update a non-focus-stealing popup list.
        # IMPORTANT: ignore navigation/confirm keys; otherwise arrow/enter will
        # trigger refresh and reset selection.
        if _event is not None:
            ks = getattr(_event, "keysym", "") or ""
            if ks in {"Up", "Down", "Return", "Escape"}:
                return
        if self._model_popup_after_id:
            try:
                self.root.after_cancel(self._model_popup_after_id)
            except Exception:
                pass
            self._model_popup_after_id = None
        self._model_popup_after_id = self.root.after(80, self._model_popup_update)

    def _cancel_model_popup_debounce(self) -> None:
        if self._model_popup_after_id:
            try:
                self.root.after_cancel(self._model_popup_after_id)
            except Exception:
                pass
            self._model_popup_after_id = None

    def _model_popup_update(self) -> None:
        self._model_popup_after_id = None
        full = list(self._model_choices_full)
        if not full:
            self._hide_model_popup()
            return
        needle = (self.model_var.get() or "").strip().lower()
        if not needle:
            filtered = full[:50]
        else:
            filtered = [m for m in full if needle in m.lower()][:50]
        if not filtered:
            self._hide_model_popup()
            return
        self._show_model_popup(filtered)

    def _show_model_popup(self, items: list[str]) -> None:
        if not hasattr(self, "model_box"):
            return
        if self._model_popup is None or not self._model_popup.winfo_exists():
            pop = Toplevel(self.root)
            pop.overrideredirect(True)
            pop.attributes("-topmost", True)
            lb = Listbox(pop, height=min(10, max(3, len(items))), exportselection=False)
            lb.pack(fill=BOTH, expand=True)
            lb.bind("<ButtonRelease-1>", self.on_model_popup_click)
            self._model_popup = pop
            self._model_popup_list = lb
        else:
            try:
                self._model_popup.deiconify()
            except Exception:
                pass
        pop = self._model_popup
        lb = self._model_popup_list
        if pop is None or lb is None:
            return

        # Position right under the combobox entry area.
        x = self.model_box.winfo_rootx()
        y = self.model_box.winfo_rooty() + self.model_box.winfo_height()
        w = self.model_box.winfo_width()
        pop.geometry(f"{w}x{0}+{x}+{y}")

        lb.delete(0, END)
        for it in items:
            lb.insert(END, it)
        lb.configure(height=min(10, max(3, len(items))))
        pop.update_idletasks()
        h = lb.winfo_reqheight()
        pop.geometry(f"{w}x{h}+{x}+{y}")

        # Keep focus on entry so typing isn't interrupted.
        try:
            self.model_box.focus_set()
        except Exception:
            pass

    def _hide_model_popup(self) -> None:
        self._cancel_model_popup_debounce()
        if self._model_popup and self._model_popup.winfo_exists():
            try:
                if self._model_popup_list and self._model_popup_list.winfo_exists():
                    self._model_popup_list.selection_clear(0, END)
                self._model_popup.withdraw()
                self._model_popup.update_idletasks()
            except Exception:
                pass

    def _is_model_popup_visible(self) -> bool:
        return bool(self._model_popup and self._model_popup.winfo_exists() and self._model_popup.state() == "normal")

    def _model_popup_current_list(self) -> list[str]:
        lb = self._model_popup_list
        if lb is None or not lb.winfo_exists():
            return []
        return [lb.get(i) for i in range(lb.size())]

    def _model_popup_select_index(self, idx: int) -> None:
        lb = self._model_popup_list
        if lb is None or not lb.winfo_exists():
            return
        if lb.size() <= 0:
            return
        idx = max(0, min(idx, lb.size() - 1))
        lb.selection_clear(0, END)
        lb.selection_set(idx)
        lb.activate(idx)
        lb.see(idx)

    def _model_popup_get_selected(self) -> str | None:
        lb = self._model_popup_list
        if lb is None or not lb.winfo_exists():
            return None
        sel = list(lb.curselection())
        if not sel:
            return None
        val = lb.get(sel[0])
        return str(val) if isinstance(val, str) and val.strip() else None

    def on_model_popup_click(self, _event=None) -> None:
        val = self._model_popup_get_selected()
        if not val:
            return
        self.model_var.set(val)
        self._hide_model_popup()
        self._save_form_into_active_session()
        try:
            self.model_box.focus_set()
        except Exception:
            pass

    def on_model_focus_out(self, _event=None) -> None:
        # Delay hide so click can register.
        self.root.after(120, self._hide_model_popup)

    def on_model_escape(self, _event=None) -> None:
        self._hide_model_popup()

    def on_model_down(self, _event=None) -> str | None:
        self._cancel_model_popup_debounce()
        if not self._is_model_popup_visible():
            self._model_popup_update()
        if self._is_model_popup_visible():
            lb = self._model_popup_list
            if lb and lb.size() > 0:
                sel = list(lb.curselection())
                next_idx = (sel[0] + 1) if sel else 0
                self._model_popup_select_index(next_idx)
            return "break"
        return None

    def on_model_up(self, _event=None) -> str | None:
        self._cancel_model_popup_debounce()
        if self._is_model_popup_visible():
            lb = self._model_popup_list
            if lb and lb.size() > 0:
                sel = list(lb.curselection())
                next_idx = (sel[0] - 1) if sel else lb.size() - 1
                self._model_popup_select_index(next_idx)
            return "break"
        return None

    def on_model_enter(self, _event=None) -> str | None:
        self._cancel_model_popup_debounce()
        if self._is_model_popup_visible():
            val = self._model_popup_get_selected()
            if val:
                self.model_var.set(val)
            self._hide_model_popup()
            self._save_form_into_active_session()
            return "break"
        return None

    def _fetch_models_worker(self, api_key: str, base_url: str) -> None:
        try:
            client = OpenAI(api_key=api_key, base_url=base_url)
            resp = client.models.list()
            models: list[str] = []
            data = getattr(resp, "data", None)
            if isinstance(data, list):
                for item in data:
                    mid = getattr(item, "id", None)
                    if isinstance(mid, str) and mid.strip():
                        models.append(mid.strip())
            models = sorted(set(models))
            if not models:
                raise RuntimeError("Provider returned an empty model list.")

            def _apply() -> None:
                self._model_choices = models
                self._model_choices_full = list(models)
                if hasattr(self, "model_box"):
                    self.model_box.configure(values=models)
                # Bind/persist into active preset if any
                if self._active_preset_name and self._active_preset_name in self.connection_presets:
                    self.connection_presets[self._active_preset_name]["models"] = list(models)
                    self._persist_connection_presets()
                self._save_form_into_active_session()
                active = self._get_active_session()
                if active:
                    self.status_var.set(f"{active.title}: {self._t('ready')}")
                else:
                    self.status_var.set(self._t("ready"))

            self.root.after(0, _apply)
        except Exception as e:
            msg = str(e)

            def _fail() -> None:
                messagebox.showerror(self._t("refresh_models"), self._t("fetch_models_failed", msg=msg))
                active = self._get_active_session()
                if active:
                    self.status_var.set(f"{active.title}: {active.status_detail or self._t('ready')}")
                else:
                    self.status_var.set(self._t("ready"))

            self.root.after(0, _fail)

    def _current_selected_input_index(self) -> int | None:
        idxs = list(self.images_list.curselection())
        return idxs[0] if idxs else None

    def _clear_thumbnail_widgets(self) -> None:
        self._thumb_photos.clear()
        self._thumb_labels.clear()
        if hasattr(self, "thumbs_frame"):
            for child in list(self.thumbs_frame.winfo_children()):
                child.destroy()

    def _schedule_thumbnail_refresh(self, session_id: str | None = None) -> None:
        session = self._get_session(session_id) if session_id else self._get_active_session()
        self._thumb_request_id += 1
        request_id = self._thumb_request_id
        self._thumb_session_id = session.session_id if session else None
        self._clear_thumbnail_widgets()
        if not session or not hasattr(self, "thumbs_frame"):
            return
        if not self._pillow_available:
            return
        if session.images:
            if session.session_id == self.active_session_id:
                self.status_var.set(f"{session.title}: {self._t('loading_thumbnails')}")
            self.root.after(0, lambda sid=session.session_id, rid=request_id: self._thumbnail_step(sid, rid, 0))
        else:
            session.selected_input_index = None
            self._refresh_sessions_list()

    def _thumbnail_step(self, session_id: str, request_id: int, idx: int) -> None:
        session = self._get_session(session_id)
        if (
            not session
            or request_id != self._thumb_request_id
            or session_id != self._thumb_session_id
            or session_id != self.active_session_id
            or not self._pillow_available
        ):
            return
        images = list(session.images)
        if idx >= len(images):
            selected_idx = session.selected_input_index
            if images:
                if selected_idx is None or selected_idx >= len(images):
                    self._select_image_index(0, scroll_thumb=False)
                else:
                    self._select_image_index(selected_idx, scroll_thumb=False)
            else:
                session.selected_input_index = None
            self._refresh_sessions_list()
            return

        from PIL import Image, ImageTk  # type: ignore

        thumb_size = (96, 96)
        pad_y = 6
        path = images[idx]
        try:
            img = Image.open(path)
            img.thumbnail(thumb_size)
            tk_img = ImageTk.PhotoImage(img)
            self._thumb_photos.append(tk_img)
            lbl = Label(self.thumbs_frame, image=tk_img, cursor="hand2")
            lbl.pack(pady=(pad_y, 0))
            lbl.bind("<Button-1>", lambda _e, i=idx: self._select_image_index(i, scroll_thumb=False))
            lbl.bind("<Double-Button-1>", lambda _e, i=idx: (self._select_image_index(i, scroll_thumb=False), self.on_preview_selected_image()))
            self._thumb_labels.append(lbl)
        except Exception:
            lbl = Label(self.thumbs_frame, text="(failed)", fg="gray")
            lbl.pack(pady=(pad_y, 0))
            self._thumb_labels.append(lbl)
        self.root.after(1, lambda sid=session_id, rid=request_id, next_idx=idx + 1: self._thumbnail_step(sid, rid, next_idx))

    def _active_images(self) -> list[Path]:
        session = self._get_active_session()
        return session.images if session else []

    def _update_session_state(
        self,
        session_id: str,
        run_id: int,
        *,
        state: str | None = None,
        detail: str | None = None,
        running: bool | None = None,
        generated_files: list[str] | None = None,
        error: str | None = None,
        session_file: str | None = None,
    ) -> None:
        session = self._get_session(session_id)
        if not session or session.current_run_id != run_id:
            return
        if state is not None:
            session.status_state = state
        if detail is not None:
            session.status_detail = detail
        if running is not None:
            session.running = running
        if generated_files is not None:
            session.generated_files = list(generated_files)
        if error is not None:
            session.last_error = error
        if session_file is not None:
            session.last_session_file = session_file
        if session.session_id == self.active_session_id:
            self.status_var.set(f"{session.title}: {session.status_detail}")
        self._refresh_sessions_list()

    def on_new_session(self) -> None:
        self._save_form_into_active_session()
        session = self._create_session(title=f"Session {len(self.sessions) + 1}")
        self._set_active_session(session.session_id)

    def on_duplicate_session(self) -> None:
        source = self._save_form_into_active_session()
        session = self._create_session(from_session=source)
        self._set_active_session(session.session_id)

    def on_rename_session(self, _event=None) -> None:
        session = self._get_active_session()
        if not session:
            return
        name = simpledialog.askstring(self._t("rename_session"), self._t("session_title"), initialvalue=session.title, parent=self.root)
        if name is None:
            return
        name = name.strip()
        if not name:
            messagebox.showerror(self._t("rename_session"), self._t("session_title_empty"))
            return
        session.title = name
        if session.session_id == self.active_session_id:
            self.status_var.set(f"{session.title}: {session.status_detail}")
        self._refresh_sessions_list()

    def on_delete_session(self) -> None:
        session = self._get_active_session()
        if not session:
            return
        if session.running:
            messagebox.showerror(self._t("delete"), self._t("delete_running_session"))
            return
        if len(self.sessions) == 1:
            messagebox.showerror(self._t("delete"), self._t("delete_last_session"))
            return
        idx = self.sessions.index(session)
        del self.sessions[idx]
        next_idx = min(idx, len(self.sessions) - 1)
        self.active_session_id = None
        self._set_active_session(self.sessions[next_idx].session_id)

    def on_session_selected(self, _event=None) -> None:
        idxs = list(self.sessions_list.curselection())
        if not idxs:
            return
        idx = idxs[0]
        if idx < 0 or idx >= len(self.sessions):
            return
        target_session_id = self.sessions[idx].session_id
        if target_session_id == self.active_session_id:
            return
        self._set_active_session(target_session_id)

    def on_open_latest_result(self, _event=None) -> None:
        session = self._get_active_session()
        if not session or not session.generated_files:
            messagebox.showerror(self._t("open_latest_result"), self._t("no_generated_result"))
            return
        latest = Path(session.generated_files[-1])
        if not latest.exists():
            messagebox.showerror(self._t("open_latest_result"), self._t("file_not_found", path=latest))
            return
        try:
            self._preview_generated_file(latest)
        except Exception as e:
            messagebox.showerror(self._t("open_latest_result"), str(e))

    def on_open_session_result_folder(self) -> None:
        session = self._get_active_session()
        if not session:
            return
        if session.generated_files:
            latest = Path(session.generated_files[-1]).parent
        else:
            latest = session.out_dir
        try:
            _open_folder(latest)
        except Exception as e:
            messagebox.showerror(self._t("open_result_folder"), str(e))

    def _preview_generated_file(self, path: Path) -> None:
        if not path.exists():
            raise RuntimeError(self._t("file_not_found", path=path))
        try:
            from PIL import Image, ImageTk  # type: ignore

            img = Image.open(path)
            max_w, max_h = 700, 700
            img.thumbnail((max_w, max_h))

            win = Toplevel(self.root)
            win.title(self._t("generated_result_title", name=path.name))

            tk_img = ImageTk.PhotoImage(img)
            self._image_preview_photo = tk_img
            lbl = Label(win, image=tk_img)
            lbl.pack(fill=BOTH, expand=True)
            Button(win, text=self._t("open_in_viewer"), command=lambda: _open_file(path)).pack(pady=8)
        except Exception:
            _open_file(path)

    def on_add_images(self) -> None:
        session = self._get_active_session()
        if not session:
            return
        self._save_form_into_active_session()
        paths = filedialog.askopenfilenames(
            title=self._t("select_images"),
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp"), ("All files", "*.*")],
        )
        if not paths:
            return
        for p in paths:
            path = Path(p)
            if path not in session.images:
                session.images.append(path)
                self.images_list.insert(END, str(path))
        self._schedule_thumbnail_refresh(session.session_id)
        self._refresh_sessions_list()

    def on_remove_image(self) -> None:
        session = self._get_active_session()
        if not session:
            return
        idxs = list(self.images_list.curselection())
        if not idxs:
            return
        for idx in reversed(idxs):
            self.images_list.delete(idx)
            del session.images[idx]
        self._schedule_thumbnail_refresh(session.session_id)
        self._refresh_sessions_list()

    def on_clear_images(self) -> None:
        session = self._get_active_session()
        if not session:
            return
        self.images_list.delete(0, END)
        session.images.clear()
        self._schedule_thumbnail_refresh(session.session_id)
        self._refresh_sessions_list()

    def on_pick_out_dir(self) -> None:
        d = filedialog.askdirectory(title=self._t("choose_output_folder"))
        if d:
            self.out_dir_var.set(d)
            self._save_form_into_active_session()

    def on_open_out_dir(self) -> None:
        try:
            session = self._save_form_into_active_session()
            _open_folder(Path(session.out_dir))
        except Exception as e:
            messagebox.showerror(self._t("open_folder"), str(e))

    def on_open_logs(self) -> None:
        try:
            _open_folder(_logs_dir())
        except Exception as e:
            messagebox.showerror(self._t("open_logs"), str(e))

    def on_load_session(self) -> None:
        path = filedialog.askopenfilename(
            title=self._t("load_session_file"),
            initialdir=str(_sessions_dir()),
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            session = self._import_session_file(Path(path), activate=True)
            if not session:
                return
            if session.generated_files:
                session.status_detail = self._t("session_imported_results", count=len(session.generated_files))
                self.status_var.set(f"{session.title}: {session.status_detail}")
        except Exception as e:
            messagebox.showerror(self._t("load_session_failed"), f"{e}\n\nSee log: {_log_file_path()}")

    def on_load_all_sessions(self) -> None:
        try:
            loaded = self._load_saved_sessions(limit=None, activate_latest=False)
            active = self._get_active_session()
            if active:
                active.status_detail = self._t("history_loaded", count=loaded)
                self.status_var.set(f"{active.title}: {active.status_detail}")
            self._refresh_sessions_list()
        except Exception as e:
            messagebox.showerror(self._t("load_session_failed"), f"{e}\n\nSee log: {_log_file_path()}")

    def on_list_selected(self, _event=None) -> None:
        idxs = list(self.images_list.curselection())
        if not idxs:
            return
        self._select_image_index(idxs[0], scroll_thumb=True)

    def on_preview_selected_image(self, _event=None) -> None:
        idxs = list(self.images_list.curselection())
        if not idxs:
            return
        images = self._active_images()
        path = images[idxs[0]]
        if not path.exists():
            messagebox.showerror(self._t("preview_failed"), self._t("file_not_found", path=path))
            return

        # Best effort: if Pillow exists, show in-app preview; otherwise open with system viewer.
        try:
            from PIL import Image, ImageTk  # type: ignore

            img = Image.open(path)
            max_w, max_h = 700, 700
            img.thumbnail((max_w, max_h))

            win = Toplevel(self.root)
            win.title(self._t("preview_title", name=path.name))

            tk_img = ImageTk.PhotoImage(img)
            self._image_preview_photo = tk_img  # prevent GC
            lbl = Label(win, image=tk_img)
            lbl.pack(fill=BOTH, expand=True)
            Button(win, text=self._t("open_in_viewer"), command=lambda: _open_file(path)).pack(pady=8)
        except Exception:
            try:
                _open_file(path)
            except Exception as e:
                messagebox.showerror(self._t("preview_failed"), str(e))

    def _select_image_index(self, idx: int, *, scroll_thumb: bool) -> None:
        session = self._get_active_session()
        if session:
            session.selected_input_index = idx
        self.images_list.selection_clear(0, END)
        self.images_list.selection_set(idx)
        self.images_list.see(idx)

        # Highlight thumbnail
        for i, lbl in enumerate(self._thumb_labels):
            if i == idx:
                lbl.configure(relief="solid", bd=2)
            else:
                lbl.configure(relief="flat", bd=0)

        if scroll_thumb and 0 <= idx < len(self._thumb_labels):
            self.root.after(0, lambda: self.thumbs_canvas.yview_moveto(max(0.0, (idx * 110) / max(1, self.thumbs_frame.winfo_height()))))

    def _collect_params(self) -> GenerateParams:
        session = self._save_form_into_active_session()
        self._persist_form_defaults()
        prompt = self.prompt_text.get("1.0", END).strip()
        if not prompt:
            raise ValueError(self._t("prompt_empty"))

        api_key = self.api_key_var.get().strip()
        if not api_key:
            raise ValueError(self._t("missing_api_key"))

        base_url = _normalize_base_url(self.base_url_var.get().strip())
        model = self.model_var.get().strip() or "gpt-image-2"
        quality = self.quality_var.get().strip().lower() or "high"
        size = self.size_var.get().strip() or "1024x1024"
        n = int(self.n_var.get().strip() or "1")
        out_dir = Path(self.out_dir_var.get().strip() or (Path("Pic_generate") / "image")).resolve()
        prefix = self.prefix_var.get().strip() or "gpt"

        return GenerateParams(
            api_key=api_key,
            base_url=base_url,
            model=model,
            prompt=prompt,
            quality=quality,
            size=size,
            n=n,
            images=list(session.images),
            out_dir=out_dir,
            filename_prefix=prefix,
        )

    def on_generate_clicked(self) -> None:
        self._generate_active_session()

    def on_generate_selected_session(self) -> None:
        idxs = list(self.sessions_list.curselection())
        if not idxs:
            return
        self._set_active_session(self.sessions[idxs[0]].session_id)
        self._generate_active_session()

    def _generate_active_session(self) -> None:
        session = self._get_active_session()
        if not session:
            return
        if session.running:
            messagebox.showerror(self._t("generate"), self._t("already_running"))
            return
        try:
            params = self._collect_params()
        except Exception as e:
            messagebox.showerror(self._t("invalid_input"), str(e))
            return

        should_open_out = session.open_out_after
        should_record_session = session.record_session
        session.run_count += 1
        session.current_run_id = session.run_count
        session.cancelled_run_ids.discard(session.current_run_id)
        session.running = True
        session.last_error = ""
        session.status_state = "running"
        session.status_detail = self._t("queued")
        run_id = session.current_run_id
        self.status_var.set(f"{session.title}: {self._t('generating_bg')}")
        self._refresh_sessions_list()
        t = threading.Thread(
            target=self._generate_worker,
            args=(session.session_id, run_id, params, should_open_out, should_record_session),
            daemon=True,
        )
        t.start()

    def _is_run_cancelled(self, session_id: str, run_id: int) -> bool:
        session = self._get_session(session_id)
        if not session:
            return True
        return run_id in session.cancelled_run_ids

    def _raise_if_cancelled(self, session_id: str, run_id: int) -> None:
        if self._is_run_cancelled(session_id, run_id):
            raise RuntimeError("__SESSION_CANCELLED__")

    def on_interrupt_session(self) -> None:
        session = self._get_active_session()
        if not session or not session.running:
            return
        run_id = session.current_run_id
        session.cancelled_run_ids.add(run_id)
        session.running = False
        session.status_state = "cancelled"
        session.status_detail = self._t("session_cancelled")
        self.status_var.set(f"{session.title}: {session.status_detail}")
        self._refresh_sessions_list()

    def on_send_background(self) -> None:
        session = self._get_active_session()
        if not session:
            return
        if not session.running:
            return
        background_title = session.title if session.title.endswith(" [BG]") else f"{session.title} [BG]"
        session.title = background_title
        self._refresh_sessions_list()
        follow_up = self._create_session(from_session=session, title=f"{session.title.replace(' [BG]', '')} Continue")
        follow_up.running = False
        follow_up.status_state = "idle"
        follow_up.status_detail = self._t("ready")
        follow_up.cancelled_run_ids.clear()
        self._set_active_session(follow_up.session_id)
        self.status_var.set(f"{follow_up.title}: {self._t('session_moved_bg')}")

    def _save_session(self, params: GenerateParams, saved_files: list[str]) -> Path:
        _safe_mkdir(_sessions_dir())
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        session_path = _sessions_dir() / f"session_{ts}.json"
        payload = {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "params": {
                "api_key": params.api_key,
                "base_url": params.base_url,
                "model": params.model,
                "prompt": params.prompt,
                "quality": params.quality,
                "size": params.size,
                "n": params.n,
                "images": [str(p) for p in params.images],
                "out_dir": str(params.out_dir),
                "filename_prefix": params.filename_prefix,
            },
            "generated_files": saved_files,
        }
        session_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return session_path

    def _extract_response_items(self, resp: object) -> list[object]:
        data = None
        if hasattr(resp, "data"):
            data = getattr(resp, "data")
        elif isinstance(resp, dict):
            data = resp.get("data")
        elif isinstance(resp, str):
            text = resp.strip()
            if not text:
                raise RuntimeError("Provider returned an empty string response.")
            try:
                parsed = json.loads(text)
            except json.JSONDecodeError:
                snippet = text[:300]
                lower = snippet.lstrip().lower()
                if lower.startswith("<!doctype html") or lower.startswith("<html") or "<head" in lower:
                    raise RuntimeError(
                        "Provider returned HTML instead of JSON. This usually means your `base_url` is not an API endpoint "
                        "(often missing `/v1`), or the request was redirected to a website/login page, or the provider rejected "
                        f"the request. Response preview: {snippet}"
                    )
                raise RuntimeError(f"Provider returned non-JSON string response: {snippet}")
            if isinstance(parsed, dict):
                if parsed.get("error"):
                    raise RuntimeError(f"Provider error: {parsed.get('error')}")
                data = parsed.get("data")
            else:
                raise RuntimeError(f"Unexpected JSON response type: {type(parsed).__name__}")
        else:
            raise RuntimeError(f"Unsupported provider response type: {type(resp).__name__}")

        if not isinstance(data, list):
            preview = str(data)[:300]
            raise RuntimeError(f"Provider response missing list field 'data'. actual={type(data).__name__}, value={preview}")
        return data

    def _extract_b64_from_item(self, item: object) -> str:
        if hasattr(item, "b64_json"):
            b64 = getattr(item, "b64_json")
        elif isinstance(item, dict):
            b64 = item.get("b64_json")
        else:
            b64 = None
        if not isinstance(b64, str) or not b64.strip():
            item_preview = str(item)[:300]
            raise RuntimeError(f"No b64_json found in response item. item={item_preview}")
        return b64

    def _extract_url_from_item(self, item: object) -> str | None:
        if hasattr(item, "url"):
            url = getattr(item, "url")
        elif isinstance(item, dict):
            url = item.get("url")
        else:
            url = None
        return url if isinstance(url, str) and url.strip() else None

    def _generate_worker(
        self,
        session_id: str,
        run_id: int,
        params: GenerateParams,
        should_open_out: bool,
        should_record_session: bool,
    ) -> None:
        try:
            self._raise_if_cancelled(session_id, run_id)
            _safe_mkdir(params.out_dir)
            client = OpenAI(api_key=params.api_key, base_url=params.base_url)
            self.root.after(
                0,
                lambda sid=session_id, rid=run_id: self._update_session_state(
                    sid, rid, state="running", detail=self._t("waiting_provider"), running=True
                ),
            )
            self._raise_if_cancelled(session_id, run_id)

            if params.images:
                # image-to-image
                opened_files = [p.open("rb") for p in params.images]
                try:
                    resp = client.images.edit(
                        model=params.model,
                        prompt=params.prompt,
                        n=params.n,
                        quality=params.quality,
                        image=opened_files,
                        response_format="b64_json",
                    )
                finally:
                    for f in opened_files:
                        f.close()
            else:
                # text-to-image
                resp = client.images.generate(
                    model=params.model,
                    prompt=params.prompt,
                    n=params.n,
                    quality=params.quality,
                    size=params.size,
                    response_format="b64_json",
                )

            items = self._extract_response_items(resp)
            saved: list[str] = []
            for i, item in enumerate(items):
                self._raise_if_cancelled(session_id, run_id)
                img_bytes: bytes | None = None
                ext: str | None = None
                url = self._extract_url_from_item(item)
                self.root.after(
                    0,
                    lambda sid=session_id, rid=run_id, i=i, total=len(items): self._update_session_state(
                        sid, rid, state="processing", detail=self._t("processing_image", current=i + 1, total=total), running=True
                    ),
                )

                # 1) Prefer b64_json
                b64_ok = False
                try:
                    b64 = self._extract_b64_from_item(item)
                    mime, b64_payload = _split_data_url(b64)
                    # If provider also provides url, avoid expensive decode on obviously-invalid base64.
                    if _is_plausible_base64(b64_payload):
                        candidate = _b64decode_relaxed(b64_payload)
                        ext = _mime_to_ext(mime) or _guess_image_ext(candidate)
                        if ext:
                            img_bytes = candidate
                            b64_ok = True
                except Exception:
                    b64_ok = False

                # 2) Fall back to url download if base64 missing/invalid/unknown
                if not b64_ok:
                    if not url:
                        raise RuntimeError(f"Provider response item has no usable image payload. item={str(item)[:300]}")
                    self.root.after(
                        0,
                        lambda sid=session_id, rid=run_id, i=i, total=len(items): self._update_session_state(
                            sid, rid, state="downloading", detail=self._t("downloading_image", current=i + 1, total=total), running=True
                        ),
                    )
                    body, content_type = _download_url_bytes(url)
                    if _looks_like_html(body):
                        snippet = body[:300].decode("utf-8", errors="replace")
                        raise RuntimeError(
                            "Downloaded image url returned HTML instead of an image. "
                            f"content_type={content_type} url={url} body_preview={snippet}"
                        )
                    ext = _guess_image_ext(body)
                    if not ext:
                        preview = body[:80]
                        raise RuntimeError(
                            "Downloaded image url returned unknown/unsupported binary. "
                            f"content_type={content_type} url={url} head_bytes={preview!r}"
                        )
                    img_bytes = body

                if not img_bytes:
                    raise RuntimeError("Provider returned empty image bytes.")

                name = _default_filename(params.filename_prefix, ext or ".png")
                if params.n > 1:
                    # Keep stable naming for multiple images in one run
                    stem = Path(name).with_suffix("").name
                    name = f"{stem}_{i+1}{Path(name).suffix}"

                out_path = _next_available_path(params.out_dir / name)
                out_path.write_bytes(img_bytes)
                saved.append(str(out_path))

            session_hint = ""
            session_file_str = ""
            self._raise_if_cancelled(session_id, run_id)
            if should_record_session:
                session_file = self._save_session(params, saved)
                session_file_str = str(session_file)
                session_hint = f" | Session: {session_file}"

            self.root.after(
                0,
                lambda sid=session_id, rid=run_id, saved=list(saved), hint=session_hint, session_file=session_file_str: self._update_session_state(
                    sid,
                    rid,
                    state="done",
                    detail=self._t("done_saved", files=", ".join(saved), hint=hint),
                    running=False,
                    generated_files=saved,
                    error="",
                    session_file=session_file,
                ),
            )
            if should_open_out:
                self.root.after(0, lambda out_dir=params.out_dir: _open_folder(out_dir))
        except Exception as e:
            if str(e) == "__SESSION_CANCELLED__":
                self.root.after(
                    0,
                    lambda sid=session_id, rid=run_id: self._update_session_state(
                        sid,
                        rid,
                        state="cancelled",
                        detail=self._t("session_cancelled"),
                        running=False,
                    ),
                )
                return
            LOGGER.exception(
                "Generate failed | base_url=%s model=%s has_images=%s n=%s",
                params.base_url,
                params.model,
                bool(params.images),
                params.n,
            )
            # NOTE: In Python 3, exception variables are cleared at the end of the
            # except block. Capture the message now for Tk callbacks.
            err_msg = str(e)
            log_path = _log_file_path()
            self.root.after(
                0,
                lambda sid=session_id, rid=run_id, msg=err_msg, lp=log_path: self._update_session_state(
                    sid, rid, state="error", detail=self._t("error_with_log", msg=msg, log=lp), running=False, error=msg
                ),
            )
            self.root.after(
                0,
                lambda sid=session_id, msg=err_msg, lp=log_path: (
                    messagebox.showerror(
                        self._t("generate_failed"),
                        f"{self._get_session(sid).title if self._get_session(sid) else 'Session'}\n\n{msg}\n\nSee log: {lp}",
                    )
                ),
            )

    def run(self) -> None:
        self.root.mainloop()


def main() -> int:
    App().run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

