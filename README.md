# Table Extractor

A zero-dependency, single-file web app that uses Claude's vision to extract tables from images or PDFs and export them as `.xlsx` files.

## Stack

- Vanilla HTML / CSS / JS — no build step, no framework
- [SheetJS](https://sheetjs.com/) for Excel generation (CDN)
- [Tabler Icons](https://tabler.io/icons) (CDN)
- [Anthropic Messages API](https://docs.anthropic.com/en/api/messages) with vision

## Running locally

Just open `table-extractor.html` in a browser. No server needed.

```bash
open table-extractor.html          # macOS
start table-extractor.html         # Windows
xdg-open table-extractor.html      # Linux
```

Enter your Anthropic API key in the field at the top. It's saved to `localStorage` so you only need to do this once per device.

## Deploying

Since it's a single HTML file, you can host it anywhere static files are served.

**GitHub Pages**
1. Create a repo, push `table-extractor.html` as `index.html`
2. Go to Settings → Pages → Deploy from branch → `main` / `root`

**Netlify**
1. Drag & drop the file onto [netlify.com/drop](https://app.netlify.com/drop)

**Vercel**
```bash
npx vercel --prod table-extractor.html
```

**Any VPS / Nginx**
Drop the file in your webroot — it's just a static asset.

## API key security

The API key is entered by the user in the browser and stored only in their `localStorage`. It is sent directly from the browser to `api.anthropic.com`.

- Suitable for personal use or internal tools where you trust the users
- **Not suitable for public deployment** where you'd expose your own key
- For a public app, build a thin backend proxy that holds the key server-side

## Supported inputs

| Format | Notes |
|--------|-------|
| PNG / JPG / WEBP | Any image containing a table |
| PDF | Works best for PDFs that are images/scans; native-text PDFs also supported |

## Customization

**Change the model** — edit the `model` field in the `fetch` call inside the `extractBtn` listener:
```js
model: 'claude-opus-4-6',   // more accurate, slower, higher cost
model: 'claude-haiku-4-5-20251001', // faster, lower cost
```

**Change the output filename** — find `XLSX.writeFile(wb, 'extracted_tables.xlsx')` and update the second argument.

**Add a logo / branding** — edit the `.header` section in the HTML.

**Remove the API key field** (if you're hardcoding or proxying) — delete the `api-card` div and replace `apiInput.value || localStorage.getItem(STORAGE_KEY)` with your key source.

## How it works

1. User drops an image or PDF onto the upload zone
2. The file is read as Base64 via the `FileReader` API
3. A request is sent to `POST /v1/messages` with the file as a vision attachment and a prompt asking for JSON-formatted table data
4. The response is parsed and rendered as a preview table
5. SheetJS converts the data to `.xlsx` and triggers a browser download
