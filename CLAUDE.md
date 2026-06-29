# Table Extractor — Project Knowledge Base

## What This Project Does
A local Flask web server that accepts scanned document images (JPG, PNG, WEBP, GIF, PDF),
sends them to the Google Gemini Vision API, extracts tabular data, and returns a growing
Excel file. The client (browser) owns the accumulated Excel blob — the server is stateless.

---

## Architecture

```
Browser (client owns Excel blob)
  │
  ├─ POST /upload  ─── file + current Excel blob
  │                         │
  │              server.py (Flask)
  │                         │
  │              ┌──────────▼──────────┐
  │              │  compress image     │  Pillow — resize ≤1280px, JPEG q75
  │              │  check MD5 cache    │  cache/{md5}.json — zero API calls on repeat
  │              │  extract GIF frames │  deduplicate via thumbnail hash
  │              │  rate limit 4.1s    │  stays under 15 RPM free tier
  │              │  call Gemini API    │  gemini-2.5-flash via REST
  │              │  parse JSON         │  extract headers + rows
  │              │  build/append Excel │  openpyxl — styled, frozen, filtered
  │              └─────────────────────┘
  │
  └─ receives updated Excel blob + row counts in headers
```

---

## Key Files

| File | Purpose |
|---|---|
| `server.py` | Everything — Flask routes, Gemini calls, Excel building, embedded HTML/JS UI |
| `.env` | `GEMINI_API_KEY=...` — never commit this |
| `server.log` | Full request/response log, viewable at `/logs` |
| `cache/*.json` | MD5-keyed API response cache — delete to force re-extraction |
| `requirements.txt` | `flask pandas openpyxl Pillow` |
| `start.bat` | First-run setup + launcher for Windows |

---

## Gemini API Details

- **Model**: `gemini-2.5-flash` (set via `GEMINI_MODEL` env var to override)
- **Key format**: `AQ.Ab8RN6...` keys are valid — the `AIzaSy` assumption was wrong
- **Free tier**: 15 RPM / 1,500 req/day — enforced by 4.1s rate limiter
- **Billing**: must be enabled on the GCP project even for free tier to work (limit: 0 error otherwise)
- **Large files**: anything > 3 MB goes through Gemini Files API (resumable upload protocol)
- **GIF limitation**: Files API does not accept `image/gif` — large GIFs are converted to PNG first

---

## Image Processing Pipeline

```
Input file
  │
  ├─ GIF?  ─── extract all frames via Pillow
  │              │  deduplicate: thumbnail 16×16 → MD5 hash
  │              │  each unique frame → compress → cache check → Gemini call
  │              └─ merge all frame tables → single output
  │
  └─ Other ─── compress: resize ≤1280px, JPEG quality 75
                 │  target: < 300 KB per API call
                 └─ cache check → Gemini call if miss
```

---

## Excel Output

- **Sheet "Data"**: all extracted rows, appended across uploads
- **Sheet "_Meta"** (hidden): tracks `source_file`, `extracted_at`, `rows_added` per run
- **Styling**: dark blue header row (#1F4E79), white bold text, frozen row 1, auto-filter
- **Dynamic columns**: new columns discovered in later pages are added automatically;
  existing rows get empty cells for the new columns

---

## Prompt Design

The extraction prompt tells Gemini to:
1. Extract EVERY column visible — no hardcoded schema
2. Match exact header text from the document (Arabic RTL aware)
3. Fill missing cells with `""` — never skip a column
4. On subsequent pages: include previously seen column names so it stays consistent

Known headers from the accumulated Excel are injected into the prompt on every upload.

---

## Endpoints

| Route | Method | Description |
|---|---|---|
| `/` | GET | Web UI |
| `/upload` | POST | Main extraction endpoint |
| `/resume` | GET | Restore last local session on page load |
| `/clear` | POST | Wipe local backup Excel |
| `/config` | GET | Returns `{model, is_local}` |
| `/logs` | GET | Last 200 lines of `server.log` |
| `/stats` | GET | Session stats: api_calls, cache_hits, rows, etc. |

---

## Common Errors & Fixes

| Error | Cause | Fix |
|---|---|---|
| `401 service account deleted` | Old/invalid API key | Create new key at aistudio.google.com/apikey |
| `429 limit: 0` | Billing not enabled on GCP project | Enable billing at console.cloud.google.com/billing |
| `429 quota exceeded` | Hit free tier daily limit | Wait until next day or create new project |
| `400 Bad Request` on GIF upload | Files API rejects image/gif | Already fixed: GIFs converted to PNG before upload |
| `TimeoutError` on large file | Inline base64 too large for 180s timeout | Already fixed: files > 3 MB use Files API |
| `Excel file is open` | File locked by Excel | Close the .xlsx file and retry |
| `model no longer available` | Gemini model deprecated | Set `GEMINI_MODEL=gemini-2.5-flash` in .env |

---

## Libraries & Skills Needed

| Library | Used For |
|---|---|
| `flask` | HTTP server and routing |
| `openpyxl` | Read/write/style Excel files |
| `pandas` | DataFrame construction for Excel rows |
| `Pillow (PIL)` | Image resize, GIF frame extraction, format conversion |
| `urllib` | Raw HTTP calls to Gemini (no SDK dependency) |
| `hashlib` | MD5 for cache keys and GIF frame dedup |

---

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `GEMINI_API_KEY` | (required) | Your Gemini API key |
| `GEMINI_MODEL` | `gemini-2.5-flash` | Model to use |
| `PORT` | `5000` | Server port (set by cloud platform) |

---

## Running Locally

```
start.bat          # first run: prompts for API key, saves to .env
python server.py   # subsequent runs
```

Server opens browser automatically at `http://localhost:5000`.
Check `http://localhost:5000/logs` for live logs after any upload.
Check `http://localhost:5000/stats` for session API usage summary.
