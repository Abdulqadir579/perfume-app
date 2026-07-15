# Perfume Cost & Margin Comparator

Upload the shop's inventory **once** — it's stored server-side. Then drop in a
supplier sheet any time to download a formatted comparison workbook showing, for
every product, the cost change (supplier's current cost vs the shop's last cost)
and the resulting margin, with color-coded flags.

Matching is by **Barcode**, so it's exact. Matched items are listed first;
products the supplier doesn't carry appear below, marked NOT FOUND.

Protected by a single shared password.

## Setup

```bash
cd perfume-app
python3.12 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

### Required environment variables

| Var | Purpose |
|---|---|
| `APP_PASSWORD` | The shared login password. **Required** — the app refuses all requests (503) without it. |
| `SECRET_KEY` | Signs session cookies. If unset, one is generated at startup and everyone is logged out on restart. Set it in production. |
| `DATA_DIR` | Where the inventory is stored. Defaults to `./data`. |
| `COOKIE_SECURE` | Cookies are marked Secure by default. Set to `0` **only** for local plain-HTTP dev. |

Generate a secret key:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

## Run locally

```bash
APP_PASSWORD="your-password" COOKIE_SECURE=0 uvicorn app.main:app --reload
```

Open http://127.0.0.1:8000 — you'll be redirected to the login page.

`COOKIE_SECURE=0` is needed on plain HTTP localhost, otherwise the browser drops
the session cookie and login appears to silently fail. Over a tunnel (HTTPS) or
in production, omit it.

## Run over a tunnel (prototype link)

```bash
# terminal 1
APP_PASSWORD="your-password" uvicorn app.main:app
# terminal 2
cloudflared tunnel --url http://localhost:8000
```

Tunnels serve HTTPS, so leave `COOKIE_SECURE` alone.

## Deploy

```bash
docker build -t perfume-comparator .
docker run -p 8000:8000 \
  -v perfume-data:/data \
  -e APP_PASSWORD="your-password" \
  -e SECRET_KEY="$(python -c 'import secrets;print(secrets.token_hex(32))')" \
  perfume-comparator
```

Two things that bite if missed:
- **Mount a volume at `/data`** or the stored inventory dies with the container.
- **Serve over HTTPS.** Secure cookies won't survive plain HTTP.

## Security notes (honest scope)

- **One shared password, no user accounts.** There is one inventory, so per-user
  logins would add complexity without adding protection.
- Sessions are signed cookies (HMAC-SHA256), HttpOnly, 14-day expiry. The cookie
  holds no secret and can't be forged without `SECRET_KEY`.
- Password comparison is constant-time.
- **No rate limiting.** A determined attacker could brute-force the password.
  Use a strong one; add rate limiting before this is public-facing.
- The app **fails closed**: no `APP_PASSWORD` means every request is refused.

## API

All endpoints except `/health` and `/login` require a valid session.

| Method | Path | Purpose |
|---|---|---|
| GET | `/login` | Login page |
| POST | `/api/login` | Sign in (form field: `password`) |
| POST | `/api/logout` | Sign out |
| GET | `/api/inventory` | What inventory is stored |
| POST | `/api/inventory` | Upload/replace inventory (field: `inventory`) |
| DELETE | `/api/inventory` | Remove stored inventory |
| POST | `/api/compare-supplier` | Compare supplier sheet vs stored inventory (field: `supplier`) |
| POST | `/api/compare` | Original two-file compare (fields: `master`, `supplier`) |
| GET | `/health` | Health check (unauthenticated) |

## File requirements

- **Inventory** must contain: `Barcode`, `Last Cost`, `Market/Selling Price`
- **Supplier sheet** must contain: `Barcode`, `Current Cost`

Column names are matched case-insensitively and trimmed. Missing columns produce
a clear message naming them. A bad inventory file is validated *before* storing,
so it can never replace a good one.

## Layout

```
perfume-app/
├── app/
│   ├── main.py         # FastAPI routes
│   ├── auth.py         # shared-password auth + sessions
│   ├── compare.py      # matching + calculation engine
│   ├── storage.py      # inventory persistence
│   └── static/
│       ├── index.html  # app (gated)
│       └── login.html  # login page
├── data/               # runtime: stored inventory (gitignored)
├── .env.example
├── requirements.txt
├── Dockerfile
└── README.md
```
