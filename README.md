# Perfume Cost & Margin Comparator

A small, single-purpose web app: upload the shop's master inventory sheet and the
latest supplier sheet, get back one formatted Excel workbook showing — for every
product — the cost difference (supplier's current cost vs the shop's last cost) and
the resulting profit margin, with color-coded flags.

Matching is done by **Barcode**, so it's exact (no fuzzy guessing).

## What it does

- Upload 2 files → download 1 comparison workbook
- Stateless: nothing is written to disk or stored; files exist only in memory during the request
- The output workbook has a **Summary** tab and a **Comparison** tab with live formulas

## Run locally

```bash
cd perfume-app
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open http://127.0.0.1:8000

## Make it a public URL (quick prototype)

With the app running locally, expose it with any tunnel, e.g.:

```bash
# Cloudflare Tunnel (no account needed for a quick trial URL)
cloudflared tunnel --url http://localhost:8000
```

That prints a public https URL you can send to the shop owner to try.

## Deploy to cloud later

A `Dockerfile` is included. The app is a standard stateless container listening on
port 8000 — it drops onto anything (ECS/Fargate, App Runner, a small EC2 box, Fly,
Render, etc.) with no database or volumes.

```bash
docker build -t perfume-comparator .
docker run -p 8000:8000 perfume-comparator
```

## File requirements

- **Shop master** must contain: `Barcode`, `Last Cost`, `Market/Selling Price`
  (plus any descriptive columns you want carried through: Brand, Perfume Name, etc.)
- **Supplier sheet** must contain: `Barcode`, `Current Cost`

Column names are matched case-insensitively and trimmed. If a required column is
missing, the app returns a clear message naming it.

## Project layout

```
perfume-app/
├── app/
│   ├── main.py         # FastAPI routes (serves UI + /api/compare)
│   ├── compare.py      # the matching + calculation engine (reusable)
│   └── static/
│       └── index.html  # single-page frontend
├── requirements.txt
├── Dockerfile
└── README.md
```
