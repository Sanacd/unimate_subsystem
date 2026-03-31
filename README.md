# UniMate backend deploy bundle

This bundle keeps the original `flask/` folder unchanged and adds deployment-only wrapper files.

## Files added
- `deploy_lite.py` — lightweight deployment entrypoint that preserves transcript/chat flow without loading embeddings
- `deploy_adapter.py` — compatibility wrapper that points to the lightweight deploy app
- `wsgi.py` — WSGI entrypoint for hosting providers
- `Procfile` — Gunicorn startup command
- `requirements_deploy.txt` — deployment dependencies
- `frontend_with_transcript_upload.html` — your UI updated with a ChatGPT-style transcript upload button

## API routes to use
- `POST /api/upload-transcript`
  - multipart/form-data
  - field name: `file`
- `POST /api/chat`
  - JSON body: `{ "message": "...", "session_id": "..." }`
- `GET /health`

## Environment variables
- `PORT` — server port on the host
- `FRONTEND_ORIGIN` — your frontend URL for CORS, for example `https://your-frontend.example.com`
- `COOKIE_SECURE` — set `true` on HTTPS deployments
- `COOKIE_SAMESITE` — use `None` for cross-site frontend/backend deployments on HTTPS

## Local run
```bash
pip install -r requirements_deploy.txt
python deploy_lite.py
```

## Important notes
- Natural-language generation now uses the Gemini API via `GEMINI_API_KEY`.
- `GEMINI_MODEL` can override the default model (`gemini-2.5-flash`).
- If Gemini is unavailable, transcript upload still succeeds, Excel generation still completes, and the app returns a structured fallback summary instead of failing.
- If OCR is needed in your PDFs, some hosts also require system packages such as Tesseract and Poppler.

## Lightweight deployment path
- Deployment no longer imports `flask/app.py`, so startup does not load `SentenceTransformer` or any embedding model.
- The unused embedding endpoints were removed from the deploy path only; the original source inside `flask/` is preserved.
- Supported deploy endpoints remain:
  - `POST /api/upload-transcript`
  - `POST /api/chat`
  - `GET /health`
