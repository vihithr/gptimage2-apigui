from __future__ import annotations

import base64
import json
import logging
import os
import sys
import threading
from urllib.parse import urlparse, urlunparse
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import (
    END,
    BOTH,
    LEFT,
    RIGHT,
    Button,
    Canvas,
    Checkbutton,
    Entry,
    Frame,
    IntVar,
    Label,
    Listbox,
    StringVar,
    Text,
    Tk,
    Toplevel,
)
from tkinter import filedialog, messagebox, ttk

from openai import OpenAI


DEFAULT_CONFIG = {
    "api_key": "",
    "base_url": "",
    "model": "",
    # Default output dir (optional). Empty => Pic_generate/image
    "out_dir": "",
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


class App:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("GPT Image GUI (gpt-image-2)")
        self.root.geometry("900x780")

        self.api_key_var = StringVar(value=_first_non_empty(CONFIG.get("api_key", ""), os.getenv("OPENAI_API_KEY", "")))
        self.base_url_var = StringVar(
            value=_first_non_empty(CONFIG.get("base_url", ""), os.getenv("OPENAI_BASE_URL", ""), "https://api.openai.com/v1")
        )
        self.model_var = StringVar(
            value=_first_non_empty(CONFIG.get("model", ""), os.getenv("OPENAI_IMAGE_MODEL", ""), "gpt-image-2")
        )
        self.quality_var = StringVar(value="high")
        self.size_var = StringVar(value="1024x1024")
        self.n_var = StringVar(value="1")
        default_out = _first_non_empty(CONFIG.get("out_dir", ""), str((Path("Pic_generate") / "image").resolve()))
        self.out_dir_var = StringVar(value=default_out)
        self.prefix_var = StringVar(value="gpt")
        self.status_var = StringVar(value="Ready.")
        self.open_out_after_var = IntVar(value=0)
        self.record_session_var = IntVar(value=1)

        self.selected_images: list[Path] = []
        self._image_preview_photo = None
        self._thumb_photos: list[object] = []
        self._thumb_labels: list[Label] = []
        self._pillow_available = self._check_pillow()
        self._selected_idx: int | None = None

        self._build_ui()

    def _check_pillow(self) -> bool:
        try:
            import PIL  # noqa: F401

            return True
        except Exception:
            return False

    def _build_ui(self) -> None:
        top = Frame(self.root)
        top.pack(fill=BOTH, padx=10, pady=10)

        def row(parent: Frame) -> Frame:
            f = Frame(parent)
            f.pack(fill=BOTH, pady=4)
            return f

        r = row(top)
        Label(r, text="API Key (OPENAI_API_KEY):", width=24, anchor="w").pack(side=LEFT)
        Entry(r, textvariable=self.api_key_var, show="*", width=80).pack(side=LEFT, fill=BOTH, expand=True)

        r = row(top)
        Label(r, text="Base URL (OPENAI_BASE_URL):", width=24, anchor="w").pack(side=LEFT)
        Entry(r, textvariable=self.base_url_var, width=80).pack(side=LEFT, fill=BOTH, expand=True)

        r = row(top)
        Label(r, text="Model:", width=24, anchor="w").pack(side=LEFT)
        Entry(r, textvariable=self.model_var, width=30).pack(side=LEFT)
        Label(r, text="Quality:", width=10, anchor="w").pack(side=LEFT, padx=(12, 0))
        ttk.Combobox(r, textvariable=self.quality_var, values=["low", "medium", "high", "auto"], width=10).pack(
            side=LEFT
        )
        Label(r, text="Size:", width=6, anchor="w").pack(side=LEFT, padx=(12, 0))
        ttk.Combobox(
            r,
            textvariable=self.size_var,
            values=["512x512", "1024x1024", "1024x1536", "1536x1024"],
            width=12,
        ).pack(side=LEFT)
        Label(r, text="N:", width=3, anchor="w").pack(side=LEFT, padx=(12, 0))
        Entry(r, textvariable=self.n_var, width=5).pack(side=LEFT)

        mid = Frame(self.root)
        mid.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

        left = Frame(mid)
        left.pack(side=LEFT, fill=BOTH, expand=True)
        right = Frame(mid)
        right.pack(side=RIGHT, fill=BOTH, expand=True, padx=(10, 0))

        Label(left, text="Prompt:").pack(anchor="w")
        self.prompt_text = Text(left, height=10)
        self.prompt_text.pack(fill=BOTH, expand=False)

        Label(left, text="Selected images (for image edit):").pack(anchor="w", pady=(10, 0))
        images_area = Frame(left)
        images_area.pack(fill=BOTH, expand=True)

        self.images_list = Listbox(images_area, height=10)
        self.images_list.pack(side=LEFT, fill=BOTH, expand=True)

        thumbs = Frame(images_area, width=120)
        thumbs.pack(side=RIGHT, fill=BOTH, padx=(8, 0))
        Label(thumbs, text="Preview").pack(anchor="w")
        if not self._pillow_available:
            Label(thumbs, text="(Install pillow for thumbnails)", fg="gray").pack(anchor="w")
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
            # Keep the inner frame width synced with canvas width.
            self.thumbs_canvas.itemconfigure(self._thumbs_window, width=event.width)

        self.thumbs_frame.bind("<Configure>", _on_thumbs_frame_configure)
        self.thumbs_canvas.bind("<Configure>", _on_thumbs_canvas_configure)

        # Mouse wheel scrolling (Windows)
        def _on_mousewheel(event) -> None:
            self.thumbs_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.thumbs_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        btns = Frame(left)
        btns.pack(fill=BOTH, pady=(6, 0))
        Button(btns, text="Add images...", command=self.on_add_images).pack(side=LEFT)
        Button(btns, text="Remove selected", command=self.on_remove_image).pack(side=LEFT, padx=(6, 0))
        Button(btns, text="Clear", command=self.on_clear_images).pack(side=LEFT, padx=(6, 0))
        Label(left, text="Tip: double-click an image to preview").pack(anchor="w", pady=(6, 0))
        self.images_list.bind("<Double-Button-1>", self.on_preview_selected_image)
        self.images_list.bind("<<ListboxSelect>>", self.on_list_selected)

        Label(right, text="Output:").pack(anchor="w")

        r = Frame(right)
        r.pack(fill=BOTH, pady=4)
        Entry(r, textvariable=self.out_dir_var, width=50).pack(side=LEFT, fill=BOTH, expand=True)
        Button(r, text="Browse...", command=self.on_pick_out_dir).pack(side=LEFT, padx=(6, 0))
        Button(r, text="Open folder", command=self.on_open_out_dir).pack(side=LEFT, padx=(6, 0))

        r = Frame(right)
        r.pack(fill=BOTH, pady=4)
        Label(r, text="Filename prefix:", width=16, anchor="w").pack(side=LEFT)
        Entry(r, textvariable=self.prefix_var, width=30).pack(side=LEFT)
        Checkbutton(r, text="Open folder after save", variable=self.open_out_after_var).pack(side=LEFT, padx=(10, 0))

        r = Frame(right)
        r.pack(fill=BOTH, pady=4)
        Checkbutton(r, text="Record session", variable=self.record_session_var).pack(side=LEFT)
        Button(r, text="Load session...", command=self.on_load_session).pack(side=LEFT, padx=(8, 0))
        Button(r, text="Open logs", command=self.on_open_logs).pack(side=LEFT, padx=(8, 0))

        run = Frame(right)
        run.pack(fill=BOTH, pady=(12, 0))
        Button(run, text="Generate", command=self.on_generate_clicked, height=2).pack(side=LEFT)
        Button(run, text="Open output folder", command=self.on_open_out_dir, height=2).pack(side=LEFT, padx=(8, 0))
        Button(run, text="Help", command=self.on_help).pack(side=LEFT, padx=(8, 0))

        status = Frame(self.root)
        status.pack(fill=BOTH, padx=10, pady=(0, 10))
        Label(status, textvariable=self.status_var, anchor="w").pack(fill=BOTH)

    def on_help(self) -> None:
        win = Toplevel(self.root)
        win.title("Help")
        msg = Text(win, height=18, width=100)
        msg.pack(fill=BOTH, expand=True)
        msg.insert(
            END,
            "Usage:\n"
            "- Fill prompt.\n"
            "- (Optional) Add images: if images are selected, the app uses images.edit (image-to-image).\n"
            "- If no images are selected, the app uses images.generate (text-to-image).\n"
            "- Set OPENAI_API_KEY and (optionally) OPENAI_BASE_URL for your provider.\n"
            "\n"
            "Notes:\n"
            "- Some providers return URLs instead of base64; this app expects b64_json.\n",
        )
        msg.configure(state="disabled")

    def on_add_images(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select images",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp"), ("All files", "*.*")],
        )
        if not paths:
            return
        for p in paths:
            path = Path(p)
            if path not in self.selected_images:
                self.selected_images.append(path)
                self.images_list.insert(END, str(path))
        self._refresh_thumbnails()

    def on_remove_image(self) -> None:
        idxs = list(self.images_list.curselection())
        if not idxs:
            return
        for idx in reversed(idxs):
            self.images_list.delete(idx)
            del self.selected_images[idx]
        self._refresh_thumbnails()

    def on_clear_images(self) -> None:
        self.images_list.delete(0, END)
        self.selected_images.clear()
        self._refresh_thumbnails()

    def on_pick_out_dir(self) -> None:
        d = filedialog.askdirectory(title="Choose output folder")
        if d:
            self.out_dir_var.set(d)

    def on_open_out_dir(self) -> None:
        try:
            _open_folder(Path(self.out_dir_var.get()))
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_open_logs(self) -> None:
        try:
            _open_folder(_logs_dir())
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_load_session(self) -> None:
        path = filedialog.askopenfilename(
            title="Load session file",
            initialdir=str(_sessions_dir()),
            filetypes=[("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            session = json.loads(Path(path).read_text(encoding="utf-8"))
            if not isinstance(session, dict):
                raise ValueError("Invalid session format.")

            params = session.get("params", {})
            if not isinstance(params, dict):
                raise ValueError("Session missing params object.")

            self.api_key_var.set(str(params.get("api_key", "")))
            self.base_url_var.set(str(params.get("base_url", "")))
            self.model_var.set(str(params.get("model", "")))
            self.quality_var.set(str(params.get("quality", "high")))
            self.size_var.set(str(params.get("size", "1024x1024")))
            self.n_var.set(str(params.get("n", "1")))
            self.out_dir_var.set(str(params.get("out_dir", self.out_dir_var.get())))
            self.prefix_var.set(str(params.get("filename_prefix", "gpt")))
            self.prompt_text.delete("1.0", END)
            self.prompt_text.insert("1.0", str(params.get("prompt", "")))

            images = params.get("images", [])
            self.images_list.delete(0, END)
            self.selected_images.clear()
            if isinstance(images, list):
                for p in images:
                    p_str = str(p).strip()
                    if p_str:
                        path_obj = Path(p_str)
                        self.selected_images.append(path_obj)
                        self.images_list.insert(END, str(path_obj))
            self._refresh_thumbnails()

            generated_files = session.get("generated_files", [])
            if isinstance(generated_files, list) and generated_files:
                self.status_var.set(
                    f"Session loaded. Referenced generated files: {', '.join(str(x) for x in generated_files)}"
                )
            else:
                self.status_var.set("Session loaded.")
        except Exception as e:
            LOGGER.exception("Load session failed: %s", path)
            messagebox.showerror("Load session failed", f"{e}\n\nSee log: {_log_file_path()}")

    def on_list_selected(self, _event=None) -> None:
        idxs = list(self.images_list.curselection())
        if not idxs:
            return
        self._select_image_index(idxs[0], scroll_thumb=True)

    def on_preview_selected_image(self, _event=None) -> None:
        idxs = list(self.images_list.curselection())
        if not idxs:
            return
        path = self.selected_images[idxs[0]]
        if not path.exists():
            messagebox.showerror("Preview failed", f"File not found:\n{path}")
            return

        # Best effort: if Pillow exists, show in-app preview; otherwise open with system viewer.
        try:
            from PIL import Image, ImageTk  # type: ignore

            img = Image.open(path)
            max_w, max_h = 700, 700
            img.thumbnail((max_w, max_h))

            win = Toplevel(self.root)
            win.title(f"Preview - {path.name}")

            tk_img = ImageTk.PhotoImage(img)
            self._image_preview_photo = tk_img  # prevent GC
            lbl = Label(win, image=tk_img)
            lbl.pack(fill=BOTH, expand=True)
            Button(win, text="Open in system viewer", command=lambda: _open_file(path)).pack(pady=8)
        except Exception:
            try:
                _open_file(path)
            except Exception as e:
                messagebox.showerror("Preview failed", str(e))

    def _select_image_index(self, idx: int, *, scroll_thumb: bool) -> None:
        self._selected_idx = idx
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

    def _refresh_thumbnails(self) -> None:
        if not hasattr(self, "thumbs_frame"):
            return
        if not self._pillow_available:
            return

        from PIL import Image, ImageTk  # type: ignore

        self._thumb_photos.clear()
        self._thumb_labels.clear()

        # Clear old thumbnail widgets
        for child in list(self.thumbs_frame.winfo_children()):
            child.destroy()

        thumb_size = (96, 96)
        pad_y = 6

        for idx, p in enumerate(self.selected_images):
            try:
                img = Image.open(p)
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

        # If list non-empty, keep selection consistent
        if self.selected_images:
            if self._selected_idx is None or self._selected_idx >= len(self.selected_images):
                self._select_image_index(0, scroll_thumb=False)
        else:
            self._selected_idx = None

    def _collect_params(self) -> GenerateParams:
        prompt = self.prompt_text.get("1.0", END).strip()
        if not prompt:
            raise ValueError("Prompt is empty.")

        api_key = self.api_key_var.get().strip()
        if not api_key:
            raise ValueError("Missing API key (set OPENAI_API_KEY or paste it here).")

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
            images=list(self.selected_images),
            out_dir=out_dir,
            filename_prefix=prefix,
        )

    def on_generate_clicked(self) -> None:
        try:
            params = self._collect_params()
        except Exception as e:
            messagebox.showerror("Invalid input", str(e))
            return

        should_open_out = self.open_out_after_var.get() == 1
        should_record_session = self.record_session_var.get() == 1
        self.status_var.set("Generating... (running in background)")
        t = threading.Thread(
            target=self._generate_worker,
            args=(params, should_open_out, should_record_session),
            daemon=True,
        )
        t.start()

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

    def _generate_worker(self, params: GenerateParams, should_open_out: bool, should_record_session: bool) -> None:
        try:
            _safe_mkdir(params.out_dir)
            client = OpenAI(api_key=params.api_key, base_url=params.base_url)

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
                )

            items = self._extract_response_items(resp)
            saved: list[str] = []
            for i, item in enumerate(items):
                b64 = self._extract_b64_from_item(item)
                img_bytes = base64.b64decode(b64)

                name = _default_filename(params.filename_prefix, ".png")
                if params.n > 1:
                    # Keep stable naming for multiple images in one run
                    stem = Path(name).with_suffix("").name
                    name = f"{stem}_{i+1}.png"

                out_path = _next_available_path(params.out_dir / name)
                out_path.write_bytes(img_bytes)
                saved.append(str(out_path))

            session_hint = ""
            if should_record_session:
                session_file = self._save_session(params, saved)
                session_hint = f" | Session: {session_file}"

            self.root.after(0, lambda: self.status_var.set(f"Done. Saved: {', '.join(saved)}{session_hint}"))
            if should_open_out:
                self.root.after(0, lambda: self.on_open_out_dir())
        except Exception as e:
            LOGGER.exception(
                "Generate failed | base_url=%s model=%s has_images=%s n=%s",
                params.base_url,
                params.model,
                bool(params.images),
                params.n,
            )
            self.root.after(0, lambda: self.status_var.set(f"Error: {e} (log: {_log_file_path()})"))
            self.root.after(0, lambda: messagebox.showerror("Generate failed", f"{e}\n\nSee log: {_log_file_path()}"))

    def run(self) -> None:
        self.root.mainloop()


def main() -> int:
    App().run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

