# UniMate backend deploy bundle

This bundle keeps the original `flask/` folder unchanged and adds deployment-only wrapper files.

## Files added
- `deploy_adapter.py` — wraps the original Flask app and exposes deploy-friendly API routes
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
python deploy_adapter.py
```

## Important notes
- The original app still expects a local Ollama endpoint for answer generation (`http://localhost:11434/api/generate`).
- If your host does not provide Ollama, transcript upload can still work, but chat replies that depend on the LLM may fail or return the app's fallback message.
- If OCR is needed in your PDFs, some hosts also require system packages such as Tesseract and Poppler.
