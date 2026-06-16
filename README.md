# Arabic Invoice Parser API

FastAPI backend for parsing Arabic PDF invoices.

## Deploy on Render

1. Push this folder to a GitHub repository
2. Go to render.com → New → Web Service
3. Connect your GitHub repo
4. Render will auto-detect render.yaml and deploy

## API Endpoints

- `GET /` — Health check
- `POST /api/parse-invoice` — Upload PDF, returns JSON + Excel base64
- `GET /docs` — Swagger UI for testing

## Local Development

```bash
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

Then open: http://localhost:8000/docs
