import streamlit as st
import pandas as pd
import sqlite3
import re
import io
import os
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="UEN Autofill", page_icon="🔍", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.stApp { background-color: #F7F6F2; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 1100px; }
.element-container { margin-bottom: 0 !important; }
.stMarkdown { margin-bottom: 0 !important; }
.top-banner { background:#1A1A1A; color:#F7F6F2; padding:1.5rem 2rem; border-radius:12px; margin-bottom:1rem; }
.top-banner h1 { font-size:1.6rem; font-weight:600; margin:0; letter-spacing:-0.5px; }
.top-banner p  { font-size:0.85rem; color:#9A9A9A; margin:0.2rem 0 0 0; }
.badge { background:#3DBA6F; color:white; font-size:0.7rem; font-weight:600; padding:0.2rem 0.6rem; border-radius:20px; letter-spacing:0.5px; text-transform:uppercase; }
.pdpa-box { background:#F5F0FF; border-left:3px solid #7B5EA7; border-radius:0 8px 8px 0; padding:0.75rem 1rem; font-size:0.82rem; color:#3D2A6B; margin-bottom:1rem; line-height:1.6; }
.pdpa-box strong { color:#5A3D9A; }
.card-top { background:white; border:1px solid #E8E6DF; border-radius:12px 12px 0 0; border-bottom:none; padding:1.4rem 1.5rem 0.8rem 1.5rem; }
.card-mid { background:white; border-left:1px solid #E8E6DF; border-right:1px solid #E8E6DF; padding:0 1.5rem 1.4rem 1.5rem; }
.card-sep { background:white; border:1px solid #E8E6DF; border-bottom:none; padding:1.4rem 1.5rem 0.8rem 1.5rem; }
.card-bot { background:white; border:1px solid #E8E6DF; border-radius:0 0 12px 12px; border-top:none; padding:0 1.5rem 1.4rem 1.5rem; }
.step-label { font-size:0.72rem; font-weight:600; letter-spacing:1px; text-transform:uppercase; color:#9A9A9A; margin-bottom:0.4rem; }
.step-title { font-size:1.05rem; font-weight:600; color:#1A1A1A; margin-bottom:0; }
.stats-row { display:flex; gap:0.8rem; flex-wrap:wrap; margin-top:1rem; }
.stat-box { flex:1; min-width:120px; background:#F7F6F2; border-radius:8px; padding:0.9rem 1rem; border:1px solid #E8E6DF; text-align:center; }
.stat-number { font-family:'DM Mono',monospace; font-size:1.6rem; font-weight:500; color:#1A1A1A; }
.stat-label  { font-size:0.72rem; color:#9A9A9A; font-weight:500; margin-top:0.2rem; }
.stat-green .stat-number { color:#3DBA6F; }
.stat-orange .stat-number { color:#F5A623; }
.stat-blue .stat-number  { color:#4A90D9; }
.stat-red .stat-number   { color:#E85454; }
.cell-ref { font-family:'DM Mono',monospace; background:#F0EEE8; border:1px solid #DDD9CE; padding:0.1rem 0.5rem; border-radius:4px; font-size:0.8rem; color:#555; }
.stDownloadButton button { background:#1A1A1A !important; color:white !important; border:none !important; border-radius:8px !important; font-family:'DM Sans',sans-serif !important; font-weight:500 !important; padding:0.6rem 1.4rem !important; font-size:0.9rem !important; width:100% !important; }
.stDownloadButton button:hover { background:#333 !important; }
.stButton button { background:#3DBA6F !important; color:white !important; border:none !important; border-radius:8px !important; font-family:'DM Sans',sans-serif !important; font-weight:500 !important; padding:0.6rem 1.4rem !important; font-size:0.9rem !important; }
.stButton button:hover { background:#35A862 !important; }
.stSelectbox > div > div, .stNumberInput > div > div > input { border-radius:8px !important; border-color:#DDD9CE !important; font-family:'DM Sans',sans-serif !important; }
.info-box { background:#EBF5FF; border-left:3px solid #4A90D9; border-radius:0 8px 8px 0; padding:0.75rem 1rem; font-size:0.85rem; color:#2C5F8A; margin-bottom:0; }
.warn-box { background:#FFF8EC; border-left:3px solid #F5A623; border-radius:0 8px 8px 0; padding:0.75rem 1rem; font-size:0.85rem; color:#7A5500; margin-bottom:0; }
.dl-section { background:#F7F6F2; border:1px solid #E8E6DF; border-radius:10px; padding:1.2rem 1.4rem; margin-top:1rem; }
.dl-section-title { font-size:0.78rem; font-weight:600; letter-spacing:0.8px; text-transform:uppercase; color:#9A9A9A; margin-bottom:0.8rem; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ────────────────────────────────────────────────────────────────
STOP_WORDS = {"pte","ltd","limited","private","co","corp","sdn","bhd","llp","inc","and","the","of","in","for","by","at"}
NA_PATTERNS = [
    re.compile(r'^n\.?a\.?$',re.I), re.compile(r'^n/a$',re.I),
    re.compile(r'^nil$',re.I),       re.compile(r'^self$',re.I),
    re.compile(r'^-+$'),             re.compile(r"^'+$"),
    re.compile(r"^'[-@\s]*$"),       re.compile(r'^[^a-zA-Z0-9]+$'),
]
GENERIC_NON_COMPANIES = {"freelance","freelancer","self-employed","unemployed"}

DB_PATHS = [
    "./database_1.db",
    "./database_2.db",
    "./database_3.db",
    "./database_4.db",
]

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def is_na(value):
    v = value.strip()
    if not v: return False
    for p in NA_PATTERNS:
        if p.match(v): return True
    return v.lower() in GENERIC_NON_COMPANIES

def is_valid_uen(value):
    v = value.strip()
    if not v or ' ' in v: return False
    if len(v) < 6 or len(v) > 15: return False
    if not re.search(r'[A-Za-z]', v): return False
    if not re.match(r'^[0-9A-Za-z]+$', v): return False
    return True

def normalise(s):
    return re.sub(r'\s+',' ', re.sub(r'[.\-,()&\'/\\]',' ', s.lower())).strip()

def meaningful_tokens(norm):
    return [t for t in norm.split() if len(t) > 1 and t not in STOP_WORDS]

def col_letter_to_index(col_str):
    idx = 0
    for ch in col_str.upper().strip(): idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1

def col_index_to_letter(idx):
    result, idx = "", idx + 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result

def clean_val(v):
    s = str(v).strip()
    return "" if s in ("nan","None","<NA>","NaN") else s

def clear_session_data():
    for key in ["processed_df","stats","preview_two_col"]:
        st.session_state.pop(key, None)

# ─── MERGED IN-MEMORY DB ──────────────────────────────────────────────────────
# We load all shards once into a single in-memory SQLite database.
# This gives us:
#   • O(1) exact lookup via indexed WHERE norm_name = ?
#   • Fast LIKE prefix search via the index (WHERE norm_name LIKE 'abc%')
#   • Token-level candidate filtering purely in SQL
# The in-memory DB is held in a @st.cache_resource so it's built once per
# worker process and reused across all user sessions.

@st.cache_resource(show_spinner="Loading reference database…")
def build_memory_db(db_paths):
    """
    Reads all shards → builds a single :memory: SQLite DB with:
      TABLE companies(norm_name TEXT, uen TEXT, aliases TEXT, tokens TEXT)
      INDEX on norm_name
      INDEX on each token word (via a separate tokens table for fast token lookup)
    Returns an open sqlite3.Connection (kept alive by the cache).
    """
    mem = sqlite3.connect(":memory:", check_same_thread=False)
    mem.execute("""
        CREATE TABLE companies (
            norm_name TEXT NOT NULL,
            uen       TEXT NOT NULL,
            aliases   TEXT NOT NULL DEFAULT '',
            tokens    TEXT NOT NULL DEFAULT ''
        )
    """)

    rows_to_insert = []
    seen = set()

    for db_path in db_paths:
        if not os.path.exists(db_path):
            continue
        conn = sqlite3.connect(db_path)
        rows = conn.execute("SELECT company_name, uen, aliases FROM companies").fetchall()
        conn.close()

        for raw_name, uen, alias_raw in rows:
            raw_name  = (raw_name  or "").strip()
            uen       = (uen       or "").strip()
            alias_raw = (alias_raw or "").strip()
            if not raw_name and not uen:
                continue
            norm_name = normalise(raw_name)
            if not norm_name or norm_name in seen:
                continue
            seen.add(norm_name)
            toks = " ".join(meaningful_tokens(norm_name))  # space-joined for LIKE search
            rows_to_insert.append((norm_name, uen, alias_raw, toks))

    mem.executemany("INSERT INTO companies VALUES (?,?,?,?)", rows_to_insert)

    # Critical: index on norm_name enables O(log n) exact + prefix lookups
    mem.execute("CREATE INDEX idx_norm ON companies(norm_name)")
    mem.commit()

    # Also build an exact-lookup dict in Python for O(1) tier-1
    exact = {r[0]: r[1] for r in rows_to_insert}

    return mem, exact, len(rows_to_insert)


# ─── SQL-BACKED FIND UEN ──────────────────────────────────────────────────────
#
# Tier 1 — Python dict exact match:          O(1)
# Tier 2 — SQL prefix LIKE on indexed col:   O(log n) via index
#           + short candidate list checked in Python for ratio guard
# Tier 3 — Alias check on tier-2 candidates: negligible (small result set)
# Tier 4 — Token lookup: SQL WHERE norm_name LIKE '%token%' per token,
#           intersect candidate sets — much smaller than full table scan
#
# We NEVER iterate over the full substr_list. Every search hits at most a
# few hundred rows from the DB, not hundreds of thousands.

def find_uen(typed_name, mem_conn, exact):
    norm_query = normalise(typed_name)
    if not norm_query:
        return ""

    # ── Tier 1: O(1) exact ──────────────────────────────────────────────────
    if norm_query in exact:
        return exact[norm_query]

    nq_len = len(norm_query)
    best_uen, best_score = "", -1
    cur = mem_conn.cursor()

    # ── Tier 2: SQL prefix / suffix LIKE ────────────────────────────────────
    # Fetch rows where norm_name starts with the query OR where the query
    # starts with norm_name (covers both "ABC Pte Ltd" ↔ "ABC" directions).
    # The index makes the LIKE 'prefix%' branch fast; the second branch
    # (query contains norm_name) needs a full scan but is guarded by ratio.
    #
    # Strategy: use the first meaningful word as a prefix anchor so the index
    # is always used for at least one branch.
    words = norm_query.split()
    prefix = words[0] if words else norm_query  # first word as index anchor

    candidates = cur.execute(
        "SELECT norm_name, uen, aliases FROM companies "
        "WHERE norm_name LIKE ? LIMIT 200",
        (prefix + "%",)
    ).fetchall()

    for en, uen, alias_raw in candidates:
        en_len = len(en)
        if nq_len <= en_len:
            sl, ratio, hit = nq_len, nq_len / en_len, norm_query in en
        else:
            sl, ratio, hit = en_len, en_len / nq_len, en in norm_query
        if hit and ratio >= 0.5:
            score = 5000 + sl
            if score > best_score:
                best_score, best_uen = score, uen

    # Also check the reverse direction: query starts with a DB entry name
    # (e.g. query="abc" matches DB entry "abc pte ltd")
    if nq_len >= 4:
        reverse = cur.execute(
            "SELECT norm_name, uen FROM companies "
            "WHERE norm_name LIKE ? LIMIT 200",
            ("%" + norm_query + "%",)  # query is substring of DB name
        ).fetchall()
        for en, uen in reverse:
            en_len = len(en)
            sl = min(nq_len, en_len)
            ratio = sl / max(nq_len, en_len)
            if ratio >= 0.5 and norm_query in en:
                score = 5000 + sl
                if score > best_score:
                    best_score, best_uen = score, uen

    if best_score >= 5000:
        return best_uen   # tier-2 hit — no need to go further

    # ── Tier 3: alias check ──────────────────────────────────────────────────
    # Fetch rows that have aliases and whose first word matches
    alias_rows = cur.execute(
        "SELECT norm_name, uen, aliases FROM companies "
        "WHERE aliases != '' AND norm_name LIKE ? LIMIT 100",
        (prefix + "%",)
    ).fetchall()

    for _en, uen, alias_raw in alias_rows:
        aliases = [normalise(a.strip()) for a in alias_raw.split(",") if a.strip()]
        for alias in aliases:
            al_len = len(alias)
            if nq_len <= al_len:
                asl, ar, ah = nq_len, nq_len / al_len if al_len else 0, norm_query in alias
            else:
                asl, ar, ah = al_len, al_len / nq_len if nq_len else 0, alias in norm_query
            if ah and ar >= 0.5:
                score = 4000 + asl
                if score > best_score:
                    best_score, best_uen = score, uen
                break

    if best_score >= 4000:
        return best_uen

    # ── Tier 4: token overlap ────────────────────────────────────────────────
    # For each meaningful query token, fetch rows whose tokens field contains
    # that token. Intersect the result sets to find rows matching ALL tokens.
    query_toks = meaningful_tokens(norm_query)
    if not query_toks:
        return best_uen if best_score >= 0 else ""

    # Start with candidates matching the longest token (most selective)
    sorted_toks = sorted(query_toks, key=len, reverse=True)

    # Fetch candidates for the most selective token
    anchor = sorted_toks[0]
    tok_candidates = cur.execute(
        "SELECT norm_name, uen FROM companies "
        "WHERE tokens LIKE ? LIMIT 500",
        (f"% {anchor} %",)  # space-padded so we match whole tokens
    ).fetchall()

    # Also try prefix/suffix of tokens column (first/last token edge cases)
    tok_candidates += cur.execute(
        "SELECT norm_name, uen FROM companies "
        "WHERE tokens LIKE ? OR tokens LIKE ? LIMIT 200",
        (f"{anchor} %", f"% {anchor}")
    ).fetchall()

    # Exact single-token match
    tok_candidates += cur.execute(
        "SELECT norm_name, uen FROM companies WHERE tokens = ? LIMIT 100",
        (anchor,)
    ).fetchall()

    # De-duplicate candidates
    seen_names = {}
    for en, uen in tok_candidates:
        seen_names[en] = uen

    for en, uen in seen_names.items():
        en_toks = meaningful_tokens(en)
        if not en_toks:
            continue
        et_set = set(en_toks)
        matched = 0
        for qt in query_toks:
            qt_len = len(qt)
            found = False
            for et in en_toks:
                if et == qt or (qt_len >= 4 and qt in et) or (len(et) >= 4 and et in qt):
                    found = True; break
            if found:
                matched += 1
            else:
                break
        if matched == len(query_toks):
            score = 2000 - abs(len(en_toks) - len(query_toks)) * 10
            if score > best_score:
                best_score, best_uen = score, uen

    return best_uen if best_score >= 0 else ""


# ─── PROCESS ──────────────────────────────────────────────────────────────────
def process_df(df, name_col_idx, uen_col_idx, header_row_idx, end_row_1based,
               mem_conn, exact):
    result_df  = df.copy()
    filled = replaced = already = na_count = no_match = 0
    data_start = header_row_idx + 1
    data_end   = (end_row_1based - 1) if end_row_1based > 0 else (len(df) - 1)
    nc = df.columns[name_col_idx]
    uc = df.columns[uen_col_idx]

    for i in range(data_start, min(data_end + 1, len(df))):
        row = df.iloc[i]
        rn  = clean_val(row.iloc[name_col_idx]) if name_col_idx < len(row) else ""
        ru  = clean_val(row.iloc[uen_col_idx])  if uen_col_idx  < len(row) else ""
        tn, tu = rn.strip(), ru.strip()

        if not tn and not tu:
            result_df.at[i, nc] = "NA"; result_df.at[i, uc] = "NA"; na_count += 1; continue
        if not tn: continue
        if is_na(tn):
            result_df.at[i, nc] = "NA"; result_df.at[i, uc] = "NA"; na_count += 1; continue
        if tn.lower() == tu.lower():
            m = find_uen(tn, mem_conn, exact); result_df.at[i, uc] = m
            replaced += bool(m); no_match += not bool(m); continue
        if is_na(tu):
            m = find_uen(tn, mem_conn, exact); result_df.at[i, uc] = m
            filled += bool(m); no_match += not bool(m); continue
        if tu and not is_valid_uen(tu):
            m = find_uen(tn, mem_conn, exact); result_df.at[i, uc] = m
            replaced += bool(m); no_match += not bool(m); continue
        if tu and is_valid_uen(tu):
            if tu != rn: result_df.at[i, uc] = tu
            already += 1; continue
        m = find_uen(tn, mem_conn, exact); result_df.at[i, uc] = m
        filled += bool(m); no_match += not bool(m)

    return result_df, {"filled":filled,"replaced":replaced,"already":already,"na":na_count,"no_match":no_match}


# ─── EXCEL BUILDERS ───────────────────────────────────────────────────────────
def build_full_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, header=False)
    return buf.getvalue()

def build_uen_only_excel(df, name_col_idx, uen_col_idx, header_row_idx):
    out_df = df.copy().fillna("").replace({"nan":"","None":"","<NA>":""})
    hdr_name = clean_val(out_df.iloc[header_row_idx, name_col_idx]) or "Company Name"
    hdr_uen  = clean_val(out_df.iloc[header_row_idx, uen_col_idx])  or "UEN"
    data_rows = []
    for i in range(header_row_idx + 1, len(out_df)):
        nv = clean_val(out_df.iloc[i, name_col_idx])
        uv = clean_val(out_df.iloc[i, uen_col_idx])
        if nv == "NA" and uv == "NA": continue
        data_rows.append((nv, uv))

    wb = Workbook(); ws = wb.active; ws.title = "UEN Results"
    hf = Font(bold=True, color="F7F6F2", name="Arial", size=10)
    hfill = PatternFill("solid", fgColor="1A1A1A")
    ws.append([hdr_name, hdr_uen])
    for c in range(1, 3):
        cell = ws.cell(row=1, column=c)
        cell.font = hf; cell.fill = hfill
        cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    alt = PatternFill("solid", fgColor="F7F6F2")
    df_ = Font(name="Arial", size=10)
    uf  = Font(name="Courier New", size=10, color="1A6E3F")
    for rn, (nv, uv) in enumerate(data_rows, start=2):
        ws.append([nv, uv])
        fill = alt if rn % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        ws.cell(rn,1).font = df_; ws.cell(rn,1).fill = fill
        ws.cell(rn,2).font = uf;  ws.cell(rn,2).fill = fill
        ws.row_dimensions[rn].height = 18
    for col_cells in ws.columns:
        cl = get_column_letter(col_cells[0].column)
        ws.column_dimensions[cl].width = min(max((len(str(c.value or "")) for c in col_cells), default=10) + 4, 60)
    ws.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  UI
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="top-banner">
  <div>
    <div style="display:flex;align-items:center;gap:0.75rem;margin-bottom:0.3rem;">
      <h1>UEN Autofill</h1><span class="badge">Singapore</span>
    </div>
    <p>Upload a spreadsheet → map your columns → download with UENs filled in</p>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="pdpa-box">
  🔒 <strong>Privacy Notice (PDPA)</strong> &nbsp;·&nbsp;
  This tool processes your uploaded file solely for UEN matching.
  Uploaded data is <strong>not stored, logged, or shared</strong> — it exists only in temporary memory during your session and is discarded when you close or refresh this page.
</div>
""", unsafe_allow_html=True)

missing = [p for p in DB_PATHS if not os.path.exists(p)]
if len(missing) == len(DB_PATHS):
    st.markdown('<div class="warn-box">⚠️ No database shards found. Place database_1.db … database_4.db in the app root.</div>', unsafe_allow_html=True)
    st.stop()

mem_conn, exact, total_records = build_memory_db(tuple(DB_PATHS))
st.markdown(f'<div class="info-box">✅ Reference database loaded — <strong>{total_records:,}</strong> company records</div>', unsafe_allow_html=True)
st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

# ── STEP 1 ────────────────────────────────────────────────────────────────────
st.markdown('<div class="card-top"><div class="step-label">Step 1</div><div class="step-title">Upload your file</div></div>', unsafe_allow_html=True)
st.markdown('<div class="card-mid">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("file", type=["xlsx","xls","csv"], label_visibility="collapsed", on_change=clear_session_data)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file is None:
    st.markdown('<div class="card-bot"><div class="info-box">👆 Upload an Excel (.xlsx / .xls) or CSV file to get started.</div></div>', unsafe_allow_html=True)
    st.stop()

@st.cache_data(show_spinner="Reading file…")
def load_file(file_bytes, file_name):
    if file_name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=str)
    return pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str, engine="openpyxl")

file_bytes = uploaded_file.read()
try:
    raw_df = load_file(file_bytes, uploaded_file.name)
except Exception as e:
    st.error(f"Could not read file: {e}"); st.stop()

num_rows, num_cols = raw_df.shape
col_letters = [col_index_to_letter(i) for i in range(num_cols)]
row_options  = [f"Row {i}" for i in range(min(num_rows, 500))]

# ── STEP 2 ────────────────────────────────────────────────────────────────────
st.markdown('<div class="card-sep"><div class="step-label">Step 2</div><div class="step-title">Map your columns</div></div>', unsafe_allow_html=True)
st.markdown('<div class="card-mid">', unsafe_allow_html=True)
c1, c2, c3, c4 = st.columns(4)
with c1: name_col_letter = st.selectbox("Company Name column", col_letters, index=0)
with c2: uen_col_letter  = st.selectbox("UEN column", col_letters, index=min(1, len(col_letters)-1))
with c3: header_row_sel  = st.selectbox("Header row", row_options, index=0, help="Row 0 = first row of file")
with c4: end_row_input   = st.number_input("Last data row (0 = auto)", min_value=0, max_value=num_rows, value=0, step=1)
st.markdown('</div>', unsafe_allow_html=True)

name_col_idx   = col_letter_to_index(name_col_letter)
uen_col_idx    = col_letter_to_index(uen_col_letter)
header_row_idx = int(header_row_sel.split()[1])

# ── STEP 3: PREVIEW ───────────────────────────────────────────────────────────
st.markdown('<div class="card-sep"><div class="step-label">Step 3</div><div class="step-title">Preview</div></div>', unsafe_allow_html=True)
st.markdown('<div class="card-mid">', unsafe_allow_html=True)

preview_df = raw_df.copy().fillna("").replace({"nan":"","None":"","<NA>":""})
preview_df.columns = [f"{col_index_to_letter(i)}  (col {i+1})" for i in range(num_cols)]
name_col_label = f"{name_col_letter}  (col {name_col_idx+1})"
uen_col_label  = f"{uen_col_letter}  (col {uen_col_idx+1})"

def highlight_cols(df):
    s = pd.DataFrame("", index=df.index, columns=df.columns)
    mask = df.index >= header_row_idx
    if name_col_label in df.columns:
        s.loc[mask, name_col_label] = "background-color:#EBF5FF;color:#2C5F8A;font-weight:500;"
    if uen_col_label in df.columns:
        s.loc[mask, uen_col_label]  = "background-color:#EBFAF2;color:#1A6E3F;font-weight:500;"
    return s

st.dataframe(preview_df.style.apply(highlight_cols, axis=None), width='stretch',
             height=min(600, 38 + num_rows * 35))

pi1, pi2 = st.columns(2)
with pi1:
    st.markdown(f'🔵 <span class="cell-ref">{name_col_letter}</span> Company Name &nbsp;&nbsp;🟢 <span class="cell-ref">{uen_col_letter}</span> UEN', unsafe_allow_html=True)
with pi2:
    data_end_row = (end_row_input - 1) if end_row_input > 0 else (num_rows - 1)
    st.markdown(f'Rows to process: <span class="cell-ref">{max(0, data_end_row - header_row_idx)}</span> &nbsp;(header = <span class="cell-ref">{header_row_sel}</span>)', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ── STEP 4 ────────────────────────────────────────────────────────────────────
st.markdown('<div class="card-sep"><div class="step-label">Step 4</div><div class="step-title">Process &amp; Download</div></div>', unsafe_allow_html=True)
st.markdown('<div class="card-bot">', unsafe_allow_html=True)

if st.button("▶  Run UEN Autofill", use_container_width=True):
    with st.spinner("Looking up UENs…"):
        processed_df, stats = process_df(
            raw_df, name_col_idx, uen_col_idx,
            header_row_idx, int(end_row_input), mem_conn, exact)
    st.session_state["processed_df"]    = processed_df
    st.session_state["stats"]           = stats
    st.session_state["preview_two_col"] = True
    st.success("Done!")

if "processed_df" in st.session_state:
    s = st.session_state["stats"]
    st.markdown(f"""
    <div class="stats-row">
      <div class="stat-box stat-green"><div class="stat-number">{s['filled']}</div><div class="stat-label">✅ UENs filled</div></div>
      <div class="stat-box stat-blue"><div class="stat-number">{s['replaced']}</div><div class="stat-label">🔄 Replaced</div></div>
      <div class="stat-box"><div class="stat-number">{s['already']}</div><div class="stat-label">⚪ Already valid</div></div>
      <div class="stat-box stat-orange"><div class="stat-number">{s['na']}</div><div class="stat-label">🚫 Marked NA</div></div>
      <div class="stat-box stat-red"><div class="stat-number">{s['no_match']}</div><div class="stat-label">❌ No match</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    out_name = Path(uploaded_file.name).stem + "_processed"

    st.markdown('<div class="dl-section"><div class="dl-section-title">📥 Download results</div>', unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.markdown("**Full sheet** *(original columns + UEN filled)*")
        st.download_button("⬇  Excel (.xlsx)", build_full_excel(st.session_state["processed_df"]),
            file_name=out_name+".xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with dl2:
        st.markdown("**Full sheet** *(CSV)*")
        st.download_button("⬇  CSV", st.session_state["processed_df"].to_csv(index=False, header=False),
            file_name=out_name+".csv", mime="text/csv", use_container_width=True)
    with dl3:
        st.markdown("**Company Name + UEN only**")
        st.download_button("⬇  UEN Results (.xlsx)",
            build_uen_only_excel(st.session_state["processed_df"], name_col_idx, uen_col_idx, header_row_idx),
            file_name=out_name+"_uen_only.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    with st.expander("Preview processed output", expanded=True):
        two_col_view = st.toggle("Show Company Name & UEN columns only",
            value=st.session_state.get("preview_two_col", True), key="preview_toggle")
        st.session_state["preview_two_col"] = two_col_view

        out_df = st.session_state["processed_df"].copy().fillna("").replace({"nan":"","None":"","<NA>":""})

        if two_col_view:
            sliced = out_df.iloc[header_row_idx:, [name_col_idx, uen_col_idx]].copy()
            new_cols = [str(sliced.iloc[0,0]) or f"{name_col_letter} — Company Name",
                        str(sliced.iloc[0,1]) or f"{uen_col_letter} — UEN"]
            sliced = sliced.iloc[1:].copy(); sliced.columns = new_cols; sliced.index = range(len(sliced))
            st.dataframe(sliced, width='stretch', height=400)
        else:
            full = out_df.iloc[header_row_idx:].copy()
            hv = [str(v) if str(v) not in ("","nan","None") else f"Col {col_index_to_letter(i)}" for i,v in enumerate(full.iloc[0])]
            full = full.iloc[1:].copy(); full.columns = hv; full.index = range(len(full))
            st.dataframe(full, width='stretch', height=400)

    st.markdown('<div style="margin-top:1rem;font-size:0.78rem;color:#9A9A9A;text-align:center;">🔒 Session data is held in memory only and discarded when you close or refresh this page.</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
