# UEN Autofill

A Streamlit web app that automatically looks up and fills in **Unique Entity Numbers (UENs)** for Singapore companies in your Excel or CSV spreadsheets.

Upload a spreadsheet → map your columns → download with UENs filled in.

---

## What it does

Many internal spreadsheets contain company names but are missing their UENs. This tool matches each company name against the full ACRA register (500 000+ entities) and writes the correct UEN back into your file. It handles:

- **Exact matches** — "Shopee Singapore Private Limited" → instant O(1) lookup
- **Partial / abbreviated names** — "Shopee Singapore" still matches
- **Alias / alternate names** — custom alias table for known variations
- **Token overlap** — "DBS Bank" matches "DBS Bank Ltd" even without the suffix
- **Dirty data** — placeholder values like `N/A`, `NIL`, `-`, `self` are normalised to `NA`; invalid UEN formats are detected and replaced; names accidentally pasted into the UEN column are fixed

---

## Features

- 📂 Upload `.xlsx`, `.xls`, or `.csv`
- 🗺️ Point-and-click column mapping with live preview (blue = company name column, green = UEN column)
- ⚡ Fast lookup — first-word inverted index + per-run deduplication cache
- 📊 Results summary (filled / replaced / already valid / marked NA / no match)
- 📥 Three download options: full sheet as Excel, full sheet as CSV, or a clean 2-column UEN results file
- 🔒 PDPA-compliant: no data is stored, logged, or transmitted — everything lives in session memory only

---

## Project structure

```
.
├── uen_autofill_app.py   # Main Streamlit app
├── requirements.txt
├── database_1.db
├── database_2.db
├── database_3.db
├── database_4.db
└── README.md
```

---

## How the matching works

The app uses a four-tier priority system to match typed company names against the ACRA register. Each tier only runs if the previous one failed.

| Tier | Method | Speed | Example |
|------|--------|-------|---------|
| 1 | Exact match (dict lookup) | O(1) | `"DBS Bank Ltd"` → exact hit |
| 2 | Substring match with length-ratio guard ≥ 0.5 | O(1) + tiny list | `"DBS Bank"` ↔ `"DBS Bank Ltd"` |
| 3 | Alias substring match | same list | `"Development Bank of Singapore"` via alias |
| 4 | Token overlap (all query tokens must appear in DB entry) | O(1) + small list | `"ST Engineering"` → `"Singapore Technologies Engineering Ltd"` |

**Stop words** (`pte`, `ltd`, `limited`, `private`, etc.) are excluded from token matching so that legal suffix differences don't block a match.

---

## Performance

| Optimisation | What it does | Impact |
|---|---|---|
| First-word inverted index | Maps every first word to its matching DB entries; replaces full-list scan | ~10–50× faster tiers 2–3 |
| Token inverted index | Maps each meaningful token to its entries; replaces LIKE scan in tier 4 | ~5–20× faster tier 4 |
| Per-run lookup cache | Caches results within a single run; skips re-lookup for repeated names | Instant for duplicates |
| `@st.cache_resource` | Indexes built once per process, shared across all sessions | No reload cost |

---

## Privacy & PDPA

This tool is designed for use in compliance with Singapore's Personal Data Protection Act (PDPA):

- Uploaded files are read into **session memory only**
- No data is written to disk, logged, or transmitted externally
- Session data is automatically discarded when the browser tab is closed or refreshed
- The ACRA register used for matching is publicly available data

---

## Contributing

Pull requests welcome. If you add support for new entity types, edge-case name formats, or additional alias sources, please include test cases with before/after examples.

---

## License

MIT
