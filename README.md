# GptImage2-APIGUI

A small Windows GUI for OpenAI-compatible **Image Generate / Edit** APIs (supports providers via `base_url`), with:

- External `config.json` (auto-created if missing)
- Optional session recording to `sessions/`
- Logs to `logs/app.log` on errors
- Thumbnail preview (via Pillow)

## Run from source

### 1) Create venv & install deps

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install --upgrade pip
.\.venv\Scripts\python -m pip install openai pillow
```

### 2) Create config

Copy `config.example.json` to `config.json` and fill your settings:

- `api_key`: your API key
- `base_url`: OpenAI-compatible base url (usually ends with `/v1`)
- `model`: image model name
- `out_dir`: output directory (default: `image`)

### 3) Start

```powershell
.\.venv\Scripts\python .\GptImage2-APIGUI.py
```

## Build (PyInstaller)

```powershell
.\.venv\Scripts\pyinstaller --clean --noconfirm --onefile --windowed --name "GptImage2-APIGUI" "GptImage2-APIGUI.py"
```

Output:

- `dist/GptImage2-APIGUI.exe`

## Notes / Troubleshooting

- If you see errors, check `logs/app.log`.
- If your provider returns **HTML** or non-JSON, it usually means the `base_url` is wrong (not an API endpoint) or missing `/v1`.

## Security

Do **NOT** commit `config.json` to git. This repository ignores it by default.
