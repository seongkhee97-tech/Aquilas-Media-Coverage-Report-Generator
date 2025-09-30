# Deploy your FastAPI Clipping Report Builder

This bundle includes the minimal files to deploy your existing `main.py`, `ui.html`, and `template.docx` to Render.com.

## Repo layout (what you should put in GitHub)

```text
repo-root/
├─ main.py               # your current server code
├─ ui.html               # your front-end (optional)
├─ template.docx         # the Word template used by the app
├─ requirements.txt      # from this bundle
├─ render.yaml           # from this bundle
└─ (optional) static/    # if you have css/js assets for ui.html
```

## One-time code changes (serve the UI from FastAPI)

In `main.py`, add:

```py
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
import os

app = FastAPI(title="Clipping Report Builder")

# Serve /static if present
STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
if os.path.isdir(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

UI_PATH = os.path.join(os.path.dirname(__file__), "ui.html")

@app.get("/", response_class=HTMLResponse)
def home():
    if os.path.exists(UI_PATH):
        return HTMLResponse(open(UI_PATH, "r", encoding="utf-8").read())
    return HTMLResponse("<h1>Clipping Report Builder</h1><p>Backend is running.</p>")
```

(Leave your existing `/build_report` endpoint as-is. CORS is already enabled for `*` in your code.)

## Deploy steps (Render, no Docker)

1. Create a **new GitHub repo** and commit the files listed above.
2. Go to **render.com** → New → **Web Service**.
3. Connect your repo.
4. Render should read `render.yaml` and pre-fill:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `uvicorn main:app --host 0.0.0.0 --port $PORT`
5. Region: **Singapore** (closest to Malaysia), Plan: **Free** (or higher).
6. Click **Create Web Service**.

After the first deploy, you’ll get a public URL like:

`https://clipping-report-builder.onrender.com/`

- Visiting `/` serves your `ui.html`.
- The front-end calls your API at `/build_report` to generate the DOCX and downloads it.
- Live logs are available in the Render dashboard (works like an online terminal).

## Local run (Windows)
```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

## Local run (macOS/Linux)
```bash
python3 -m venv .venv
source ./.venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --host 0.0.0.0 --port 8000
```

## Notes
- If you also include a `Dockerfile` in the repo, Render will prefer Docker — either remove it or use one deployment method at a time.
- Make sure `template.docx` is committed so the server can read it at runtime.
- If your `ui.html` uses relative assets (CSS/JS), put them under `static/` and reference them with `/static/...` paths.
- The filesystem on Render is ephemeral; files saved at runtime are temporary only. That's fine for your `FileResponse` download flow.
- Logs/“terminal” are in the Render service dashboard.
