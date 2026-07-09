# Live Scheduling Tracker

*Memphis RPDC schedule & tour-coverage tracker — roster in, coverage gaps out,
one standalone weekly report.*

**Live app:** <https://andrewgh99.github.io/Live-Scheduling-Tracker/> ·
**Printable guide:** [HOW_TO_USE.pdf](HOW_TO_USE.pdf)
(regenerate with `python3 make_guide.py` after workflow changes)

Each FTC confirms the hours they can work (a simple Excel roster) → the tool
pops out the polished weekly board **and** scores coverage against the USPS
tours and the SOP station-priority model, so you can see exactly **where you're
short across all 3 tours × 7 days**.

Built as a reusable system: the roster→board core is isolated, so a future live
source (shared sheet / Hubstaff / Zoom presence) swaps only the *input*.

## Tours & the coverage logic (from the SOP)

Coverage is scored per USPS tour:

| Tour | Window |
|------|--------|
| **Tour 1** | 11:00 PM – 6:00 AM (overnight) |
| **Tour 2** | 6:00 AM – 2:00 PM (day) |
| **Tour 3** | 2:00 PM – 10:00 PM (evening) |

A person **counts toward a tour if they're on-clock for ≥ half of it.**
`Avail.` (the 24/7 lead) counts toward every tour as the always-on layer.

Headcount on a tour maps to stations manned, per the SOP *"Staffing Priority
24/7/365"* (fill in order #1→#4):

| Bodies | Stations manned | Status |
|--------|-----------------|--------|
| **4+** | #1 Floor Lead · #2 Docks · #3 Yard/Inbound · #4 Overflow | 🟢 Full |
| **3**  | #1 · #2 · #3 (Overflow down) | 🟡 |
| **2**  | #1 · #2 (Yard + Overflow down) | 🟠 |
| **1**  | #1 Floor Lead only | 🔴 Critical |
| **0**  | none — breaches the 24/7 Station #1 mandate | ⬛ |

Baseline = **4 bodies/tour** (all four stations covered).

## The two ways to use it

### 1. Live browser app — `rpdc_live.html`  (recommended)
Double-click it — opens in any browser, fully offline.

**The focus:** confirm coverage across the 3 tours × 7 days, **finalize the
week, and export ONE standalone report.** The end product is a finished
shareable report — not a tracker you keep open forever.

Tabs:
- **📖 How to Use** — the focus + the 5-step 7-day workflow (opens here first).
- **🔴 Live Now** — optional floor monitor. Auto-detects the **current tour**,
  shows live bodies-online → stations covered, flags who's scheduled now but off (**GAP**).
- **📊 Board & Gaps** — the 3-tour × 7-day grid + **gap list** (worst first), and
  the **Finalize the week** strip: click each day chip once its coverage is
  locked (or *Finalize all 7*); a counter tracks X/7.
- **📋 Roster (edit)** — edit inline or **drag an Excel roster on**. Each person's
  hours = confirmed availability. Saves in the browser.
- **ℹ️ Legend & SOP** — the 4 stations + color meanings.

**📄 Export Weekly Report** (header, or on Board & Gaps) → downloads
`RPDC_Week_Report_<date>.html`: a **single self-contained file** (no server, no
dependencies, no scripts) with the finalized week — coverage grid, gaps, roster,
and legend. Open / print / email / post it anywhere.

Other buttons: **Import roster (.xlsx)** · **Save roster (.xlsx)** (feeds
`convert.py` for the styled Excel board) · **Reset to crew**.

### 2. Excel round-trip — `convert.py`
```
python3 convert.py INPUT_TEMPLATE.xlsx RPDC_Board.xlsx
```
- **INPUT_TEMPLATE.xlsx** — the roster the team fills (pre-filled with the 11,
  has an INSTRUCTIONS tab).
- **RPDC_Board.xlsx** — styled weekly board + tour-coverage grid + gap table.
- It also prints the gap list to the terminal.

First-time setup: `python3 -m pip install openpyxl`

## Roster format — built to be filled in seconds
`Employee Name · Contact/Badge · Role · Shift · Station · Mon…Sun · Notes`

Every **day cell is a click-a-dropdown** — no typing needed to confirm coverage:
> `Off` · `Tour 1 (11p–6a)` · `Tour 2 (6a–2p)` · `Tour 3 (2p–10p)` · `Avail.`

Shift and Station are dropdowns too. Power users can still **type exact hours**
instead (`9a–6p`, `10p–8a`, `6:30a–3p`, split shift `10–11p / 8–9a`) — the
converter parses both. `Off` / blank = not working.

Both the template and the board carry a **“Legend & Definitions”** tab with the
4 stations (SOP) and what each color means — see below.

## Stations & colors (baked into every file — the “Legend & Definitions” tab)

**The 4 stations** (SOP — filled in this priority order):
1. **#1 Inside/Outside Floor Lead** — ALWAYS staffed 24/7/365; interior floor +
   docks, first responder, driver badging/PPE/check-in. Never empty.
2. **#2 Outside Docks** — dock patrol, hourly chock/trailer audits, meets every
   truck, strap collection, driver-compliance fines, trucks sealed before departure.
3. **#3 Trailer Yard / Inbound** — inbound four-photo inspection, yard audit,
   OOS flags, daily technician check.
4. **#4 Overflow / Support** — only when a 4th body is free; backs up #2/#3,
   carrier hunting, decaling, the 3 PM/3 AM reactivated-trailer list.

**Why the colors** (bodies on a tour → stations manned, per SOP Staffing Priority):
🟢 4+ all 4 stations · 🟡 3 = #1·#2·#3 (Overflow down) · 🟠 2 = #1·#2 (Yard at
risk) · 🔴 1 = #1 only (critical) · ⬛ 0 = 24/7 breach.

## Demo-crew read (verified — app and Python agree exactly)

The starter roster is a **fictional demo crew** with realistic shift patterns —
load your real roster via Excel drop or the shared Google Sheet. On the demo
data, worst gaps:
- **Tour 3 (2p–10p) Saturday → 2 bodies** (only #1·#2; Yard + Overflow down)
- **Tour 1 (11p–6a) Sunday → 2 bodies** (only #1·#2)
- 10 more tour/days sit at 3 (Overflow station down).

> **Data flag pattern to watch:** two demo rows are entered as `10–11p / 8–9a`,
> a split that overlaps **no** tour by half — so they don't register as covering
> any tour and their weekly hours read low. If a real person's cell looks like
> that but they actually work a full overnight shift, fix the hours in the
> roster and re-run; the night gaps will shrink.

## Notes on the model
- The 10–11 PM window sits between Tour 3 (ends 10p) and Tour 1 (starts 11p) per
  the tour definitions — the live view labels that hour a "seam."
- `Avail.` staff count toward every tour. To score raw crew-only coverage,
  change that behavior in `tour_overlap()` (convert.py) / `tourOverlap()`
  (rpdc_live.html).

## Load Coordinators (LC) — coming next
`make_template.py LC` produces `INPUT_TEMPLATE_LC.xlsx` — a blank starter with
the same easy-fill structure. **Not generated yet**: LCs are a different role
than the yard FTCs, so their stations/tours (if different) need defining before
the LC board is meaningful. Confirm the LC station model, then we build its
coverage view the same way.

## Files
- `rpdc_live.html` — live browser app · `xlsx.full.min.js` — SheetJS (offline Excel I/O)
- `convert.py` — roster → board core (+ shared Legend sheet) · `make_template.py` — regenerates the template (`LC` arg for the LC starter)
- `INPUT_TEMPLATE.xlsx` / `RPDC_Board.xlsx` — team-facing input & output (each has a Legend & Definitions tab)

## Distribution — no hosting service required

**`index.html` is the whole app in one file** (SheetJS inlined — rebuild with
`python3 build.py` after editing `rpdc_live.html`). Because it's self-contained,
you can skip hosting services entirely:

- **Shared drive / Google Drive / Dropbox** — drop `index.html` in the team
  folder; everyone double-clicks it. Updates = replace the file.
- **Telegram / email** — send the file once; it runs from Downloads forever.
- **The RPDC laptop** — keep it on the desktop as the site's scheduling tool.
- **GitHub Pages** — free hosting straight from this repo, no Netlify (see below).

Each copy keeps its own local edits/finalized days (browser storage), but when
connected to the shared Google Sheet, everyone reads the same roster — and the
exported report is the shared artifact anyway.

## Hosting — use it from anywhere, keep the roster current as people churn

The tool is a static file, so making it reachable "any time, any device" is free
and takes minutes. The roster stays current via a **shared Google Sheet** that HR
maintains — add a hire or remove someone in the Sheet, hit **↻ Refresh**, done.

**GitHub Pages (from this repo — no third-party service):**
1. Push this repo to GitHub (see below).
2. Repo → Settings → Pages → Source: `main` branch, `/ (root)` → Save.
3. Your URL is `https://<username>.github.io/<repo>/` — it serves `index.html`
   automatically.
   The repo ships only a **fictional demo crew**, so it is safe to make public
   for free-tier Pages. Keep real rosters out of the repo: they live in the
   shared Google Sheet (or local files under `private/`, which is git-ignored).

**Step 1 — put the roster in a Google Sheet.**
- Make a Google Sheet with the same header row as the template
  (`Employee Name, Contact/Badge, Role, Shift, Station, Mon…Sun, Notes`), first tab.
- Day cells accept the same values: `Off`, `Tour 1 (11p–6a)`, `Tour 2 (6a–2p)`,
  `Tour 3 (2p–10p)`, `Avail.`, or exact hours like `10p–8a`.
- Share it **“Anyone with the link → Viewer”** (or File → Share → Publish to web → CSV).

**Step 2 — host the app** (upload `rpdc_live.html` + `xlsx.full.min.js`):
- **Netlify Drop** (netlify.com/drop) — drag the two files onto the page → instant URL. Free.
- or **GitHub Pages** — push the folder, enable Pages. Free.
- or **Cloudflare Pages / Vercel** — same idea. Free.

**Step 3 — connect once.** Open the hosted URL → Roster tab → paste the Sheet
link → **Connect & Refresh**. The URL is remembered, so it auto-loads next time;
anyone opening the page and hitting Refresh sees the current crew.

Notes:
- The Sheet is the shared source of truth (multi-editor, edit history, permissions —
  all Google's, nothing to run). The board/report read from it live.
- Finalized-days and live on/off toggles are per-browser (the scheduler's view); the
  **exported report** is the shared, frozen artifact.
- A "link-viewable" Sheet is readable by anyone with the URL — fine for internal
  rosters. If names must stay private, keep the app + Sheet behind a login (e.g. host
  on an access-controlled Netlify/Cloudflare site) or use a private-sheet proxy.

## v2 — shared / auto-fed status
Single-screen today (one lead's laptop or a wall monitor). For multi-user
real-time, or to auto-feed on/off status instead of manual toggles, swap the
input reader for **Google Sheets / Hubstaff presence / Zoom presence** (the SOP
already mandates staying in the RPDC Zoom all shift). Coverage math and rendering
stay untouched.
