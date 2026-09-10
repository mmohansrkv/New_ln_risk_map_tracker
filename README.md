# LN Risk Map — Productivity Tracker (Python / Flask version)

Same layout, tabs and workflow as the original single-file HTML tool, rebuilt
so the data lives on a **server** instead of the browser's localStorage.
That means the admin and every employee can log in from any device/browser
and see the same shared data.

## What changed vs. the HTML-only version
- **Backend:** `app.py` — a Flask app exposing a small JSON API
  (`/api/employees`, `/api/processes`, `/api/productivity`, `/api/notes`,
  `/api/leaves`, `/api/export/...`).
- **Storage:** `data.json` — created automatically on first run, sits next
  to `app.py`. All records (employees, processes, productivity entries,
  notes, leaves) live here instead of in the browser.
- **Frontend:** `templates/index.html` — the exact same visual design
  (dark card UI, tabs, forms, tables) but its JavaScript now calls the
  Flask API with `fetch()` instead of reading/writing `localStorage`.
- **Sessions:** login is a real server-side session (secure cookie) instead
  of a value stored in `localStorage`.
- **Excel export:** generated server-side with `openpyxl` (no CDN needed)
  and streamed back as a download — same 5 export buttons as before
  (Productivity, Notes, Leaves, Full Database).

## Setup

```bash
cd ln_tracker
python -m venv venv                 # optional but recommended
source venv/bin/activate            # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Then open **http://127.0.0.1:5000/** in your browser.

## Default admin login
- Email: `admin@lnriskmap.com`
- Password: `Admin@123`

Change these by editing `DEFAULT_DB["admin"]` in `app.py` before first run,
or by editing the `"admin"` block in `data.json` after it's created.

## Login links
Same pattern as before, just with a real URL instead of a local file path:
- `http://127.0.0.1:5000/#admin` → Admin login
- `http://127.0.0.1:5000/#employee` → Employee login (blank)
- `http://127.0.0.1:5000/#employee=EMP001` → Employee login, pre-filled
  with that employee's email

## Notes / things to know before you deploy this for real use
- `data.json` is a simple flat file — fine for a small team, but there's
  no row-level locking beyond a basic in-process lock, so it's not meant
  for heavy concurrent write traffic. For a bigger rollout, swap
  `load_db()`/`save_db()` for a real database (SQLite via `sqlite3`, or
  Postgres) — the rest of the app doesn't need to change.
- Employee passwords are stored in plain text in `data.json`, matching the
  original tool's behavior (the admin's Employee List shows the password
  column so it can be shared with new hires). If this will hold real
  people's credentials, consider hashing passwords and giving admins a
  "reset password" action instead of a visible password field.
- Change `app.secret_key` in `app.py` (or set the `LNRM_SECRET_KEY`
  environment variable) before deploying anywhere outside your own machine.
- To run this for a real team, host it somewhere reachable (a small VM,
  PythonAnywhere, Render, etc.) and share the site's URL instead of
  `127.0.0.1`.
