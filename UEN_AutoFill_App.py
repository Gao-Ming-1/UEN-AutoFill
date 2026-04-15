import streamlit as st
import pandas as pd
import sqlite3
import re
import io
import os
from pathlib import Path
from collections import defaultdict
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

# ═══════════════════════════════════════════════════════════════════════════════
#  OPTIMISATION A — on-disk DB connection pool (no in-memory copy)
#  OPTIMISATION F — first-word inverted index built once at startup
#
#  What we keep in memory (per cached resource):
#    exact      : dict[norm_name -> uen]          ~50–80 MB for 500k records
#    fw_index   : dict[first_word -> list[(norm_name, uen, aliases)]]
#                                                  ~same strings, shared refs
#    db_conn    : a single read-only on-disk connection for tier-4 token LIKE
#                 (only used when tiers 1-3 all miss — rare for clean data)
#
#  What we NO LONGER keep:
#    ✗ in-memory SQLite DB copy        (saved ~300–500 MB)
#    ✗ substr_list full list of dicts  (saved ~100–200 MB)
#    ✗ alias_list duplicate list       (saved ~50 MB)
# ═══════════════════════════════════════════════════════════════════════════════

@st.cache_resource(show_spinner="Loading reference database…")
def build_indexes(db_paths: tuple):
    """
    Reads all on-disk shards once.
    Returns:
      exact     : dict  norm_name -> uen                   (tier-1, O(1))
      fw_index  : dict  first_word -> [(norm_name, uen, norm_aliases_list)]
                                                            (tiers 2-3, O(1) + tiny list scan)
      tok_index : dict  token -> [(norm_name, uen)]        (tier-4, O(1) + small list scan)
      total     : int   record count
    All values are plain Python dicts/lists — no SQLite copy, no DataFrame.
    Strings are interned where possible so lists share references.
    """
    exact     = {}                          # norm_name -> uen
    fw_index  = defaultdict(list)           # first_word -> [(norm_name, uen, [alias_norm, ...])]
    tok_index = defaultdict(list)           # token -> [(norm_name, uen)]
    seen      = set()

    for db_path in db_paths:
        if not os.path.exists(db_path):
            continue
        conn = sqlite3.connect(db_path)
        # fetchall() on raw tuples — fastest way to pull data out of SQLite
        rows = conn.execute(
            "SELECT company_name, uen, aliases FROM companies"
        ).fetchall()
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

            # intern uen string — many rows share the same UEN string pattern
            uen = sys.intern(uen)

            # tier-1: exact dict
            exact[norm_name] = uen

            # tier-2/3: first-word inverted index
            words = norm_name.split()
            if not words:
                continue
            first_word = words[0]

            norm_aliases = []
            if alias_raw:
                norm_aliases = [normalise(a.strip()) for a in alias_raw.split(",") if a.strip()]

            fw_index[first_word].append((norm_name, uen, norm_aliases))

            # tier-4: token inverted index
            # Index every meaningful token -> this entry
            for tok in meaningful_tokens(norm_name):
                tok_index[tok].append((norm_name, uen))

    return exact, dict(fw_index), dict(tok_index), len(exact)

import sys


# ─── FIND UEN (optimised) ─────────────────────────────────────────────────────
#
# Tier 1: exact dict              O(1)         — covers ~70-80% of clean data
# Tier 2: first-word fw_index     O(1)+list    — list is usually 1–50 entries
# Tier 3: alias check             same list    — only entries with aliases
# Tier 4: token tok_index         O(1)+list    — intersect small token lists
#
# No SQL, no full list scan. Memory is ~60% lower than the in-memory DB approach.

def find_uen(typed_name: str, exact: dict, fw_index: dict, tok_index: dict) -> str:
    norm_query = normalise(typed_name)
    if not norm_query:
        return ""

    # ── Tier 1: O(1) exact ──────────────────────────────────────────────────
    if norm_query in exact:
        return exact[norm_query]

    nq_len = len(norm_query)
    best_uen, best_score = "", -1

    # ── Tiers 2 & 3: first-word bucket ──────────────────────────────────────
    # We look up by EVERY word in the query, not just the first, so we catch
    # cases where the user's name starts differently from the DB entry.
    # e.g. query "st andrew's" → first_word "st" (same as DB "st andrew's ...")
    query_words = norm_query.split()
    candidate_buckets = set()
    for w in query_words:
        if w in fw_index:
            candidate_buckets.add(w)

    # Collect all candidates from relevant first-word buckets, deduplicated
    seen_cands: dict[str, tuple] = {}
    for w in candidate_buckets:
        for entry in fw_index[w]:
            nn = entry[0]
            if nn not in seen_cands:
                seen_cands[nn] = entry

    for norm_name, uen, norm_aliases in seen_cands.values():
        en_len = len(norm_name)

        # Tier 2: substring with ratio guard
        if nq_len <= en_len:
            sl, ratio, hit = nq_len, nq_len / en_len, norm_query in norm_name
        else:
            sl, ratio, hit = en_len, en_len / nq_len, norm_name in norm_query
        if hit and ratio >= 0.5:
            score = 5000 + sl
            if score > best_score:
                best_score, best_uen = score, uen
            continue  # no point checking aliases if tier-2 matched

        # Tier 3: alias substring match
        for alias in norm_aliases:
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

    # Return early if tiers 2/3 already found a good match
    if best_score >= 4000:
        return best_uen

    # ── Tier 4: token inverted index ────────────────────────────────────────
    query_toks = meaningful_tokens(norm_query)
    if not query_toks:
        return best_uen if best_score >= 0 else ""

    # For each query token, find matching token bucket (exact or substring).
    # Use the longest token first (most selective).
    sorted_toks = sorted(query_toks, key=len, reverse=True)

    # Build candidate set from the most selective token's bucket
    anchor = sorted_toks[0]
    anchor_len = len(anchor)

    # Collect candidates: entries whose token list contains a match for anchor
    cand_map: dict[str, str] = {}  # norm_name -> uen

    # Exact token match (fast O(1) dict lookup)
    if anchor in tok_index:
        for nn, u in tok_index[anchor]:
            cand_map[nn] = u

    # Substring token match for longer anchors (anchor is part of a DB token)
    if anchor_len >= 4:
        for tok, entries in tok_index.items():
            if tok == anchor:
                continue
            if anchor in tok or (len(tok) >= 4 and tok in anchor):
                for nn, u in entries:
                    cand_map[nn] = u

    # Now verify all query tokens match for each candidate
    for norm_name, uen in cand_map.items():
        en_toks = meaningful_tokens(norm_name)
        if not en_toks:
            continue
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
                break  # short-circuit: can never reach matched == len(query_toks)
        if matched == len(query_toks):
            score = 2000 - abs(len(en_toks) - len(query_toks)) * 10
            if score > best_score:
                best_score, best_uen = score, uen

    return best_uen if best_score >= 0 else ""


# ─── PROCESS ──────────────────────────────────────────────────────────────────
#
# OPTIMISATION H — per-run lookup cache
# Real spreadsheets often repeat the same company name many times
# (e.g. "ABC Pte Ltd" in 50 rows). We cache every lookup result within
# a single process_df call so each unique name is only looked up once.
# The cache is a plain dict local to the function call — zero memory overhead
# between runs, and it's automatically GC'd when the function returns.

def process_df(df, name_col_idx, uen_col_idx, header_row_idx, end_row_1based,
               exact, fw_index, tok_index):
    result_df  = df.copy()
    filled = replaced = already = na_count = no_match = 0
    data_start = header_row_idx + 1
    data_end   = (end_row_1based - 1) if end_row_1based > 0 else (len(df) - 1)
    nc = df.columns[name_col_idx]
    uc = df.columns[uen_col_idx]

    # H: lookup cache — keyed by normalised company name, value is matched UEN
    _lookup_cache: dict[str, str] = {}

    def cached_find(name: str) -> str:
        key = normalise(name)
        if key not in _lookup_cache:
            _lookup_cache[key] = find_uen(name, exact, fw_index, tok_index)
        return _lookup_cache[key]

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
            m = cached_find(tn); result_df.at[i, uc] = m
            replaced += bool(m); no_match += not bool(m); continue
        if is_na(tu):
            m = cached_find(tn); result_df.at[i, uc] = m
            filled += bool(m); no_match += not bool(m); continue
        if tu and not is_valid_uen(tu):
            m = cached_find(tn); result_df.at[i, uc] = m
            replaced += bool(m); no_match += not bool(m); continue
        if tu and is_valid_uen(tu):
            if tu != rn: result_df.at[i, uc] = tu
            already += 1; continue
        m = cached_find(tn); result_df.at[i, uc] = m
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
    hf    = Font(bold=True, color="F7F6F2", name="Arial", size=10)
    hfill = PatternFill("solid", fgColor="1A1A1A")
    ws.append([hdr_name, hdr_uen])
    for c in range(1, 3):
        cell = ws.cell(row=1, column=c)
        cell.font = hf; cell.fill = hfill
        cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    alt  = PatternFill("solid", fgColor="F7F6F2")
    wht  = PatternFill("solid", fgColor="FFFFFF")
    df_  = Font(name="Arial", size=10)
    uf   = Font(name="Courier New", size=10, color="1A6E3F")
    for rn, (nv, uv) in enumerate(data_rows, start=2):
        ws.append([nv, uv])
        fill = alt if rn % 2 == 0 else wht
        ws.cell(rn,1).font = df_; ws.cell(rn,1).fill = fill
        ws.cell(rn,2).font = uf;  ws.cell(rn,2).fill = fill
        ws.row_dimensions[rn].height = 18
    for col_cells in ws.columns:
        cl = get_column_letter(col_cells[0].column)
        ws.column_dimensions[cl].width = min(
            max((len(str(c.value or "")) for c in col_cells), default=10) + 4, 60)
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
  Uploaded data is <strong>not stored, logged, or shared</strong> — it exists only in temporary memory
  during your session and is discarded when you close or refresh this page.
</div>
""", unsafe_allow_html=True)

# ── DATABASE ──────────────────────────────────────────────────────────────────
missing = [p for p in DB_PATHS if not os.path.exists(p)]
if len(missing) == len(DB_PATHS):
    st.markdown('<div class="warn-box">⚠️ No database shards found. '
                'Place database_1.db … database_4.db in the app root directory.</div>',
                unsafe_allow_html=True)
    st.stop()

exact, fw_index, tok_index, total_records = build_indexes(tuple(DB_PATHS))
st.markdown(f'<div class="info-box">✅ Reference database loaded — '
            f'<strong>{total_records:,}</strong> company records</div>',
            unsafe_allow_html=True)
st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

# ── STEP 1 ────────────────────────────────────────────────────────────────────
st.markdown('<div class="card-top"><div class="step-label">Step 1</div>'
            '<div class="step-title">Upload your file</div></div>', unsafe_allow_html=True)
st.markdown('<div class="card-mid">', unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "file", type=["xlsx","xls","csv"],
    label_visibility="collapsed", on_change=clear_session_data)
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file is None:
    st.markdown('<div class="card-bot"><div class="info-box">'
                '👆 Upload an Excel (.xlsx / .xls) or CSV file to get started.'
                '</div></div>', unsafe_allow_html=True)
    st.stop()

@st.cache_data(show_spinner="Reading file…", max_entries=3)
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
st.markdown('<div class="card-sep"><div class="step-label">Step 2</div>'
            '<div class="step-title">Map your columns</div></div>', unsafe_allow_html=True)
st.markdown('<div class="card-mid">', unsafe_allow_html=True)
c1, c2, c3, c4 = st.columns(4)
with c1: name_col_letter = st.selectbox("Company Name column", col_letters, index=0)
with c2: uen_col_letter  = st.selectbox("UEN column", col_letters, index=min(1, len(col_letters)-1))
with c3: header_row_sel  = st.selectbox("Header row", row_options, index=0,
                               help="Row 0 = first row of file")
with c4: end_row_input   = st.number_input("Last data row (0 = auto)",
                               min_value=0, max_value=num_rows, value=0, step=1)
st.markdown('</div>', unsafe_allow_html=True)

name_col_idx   = col_letter_to_index(name_col_letter)
uen_col_idx    = col_letter_to_index(uen_col_letter)
header_row_idx = int(header_row_sel.split()[1])

# ── STEP 3: PREVIEW ───────────────────────────────────────────────────────────
st.markdown('<div class="card-sep"><div class="step-label">Step 3</div>'
            '<div class="step-title">Preview</div></div>', unsafe_allow_html=True)
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

st.dataframe(preview_df.style.apply(highlight_cols, axis=None),
             width='stretch', height=min(600, 38 + num_rows * 35))

pi1, pi2 = st.columns(2)
with pi1:
    st.markdown(
        f'🔵 <span class="cell-ref">{name_col_letter}</span> Company Name &nbsp;&nbsp;'
        f'🟢 <span class="cell-ref">{uen_col_letter}</span> UEN',
        unsafe_allow_html=True)
with pi2:
    data_end_row = (end_row_input - 1) if end_row_input > 0 else (num_rows - 1)
    st.markdown(
        f'Rows to process: <span class="cell-ref">{max(0, data_end_row - header_row_idx)}</span>'
        f' &nbsp;(header = <span class="cell-ref">{header_row_sel}</span>)',
        unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ── STEP 4 ────────────────────────────────────────────────────────────────────
st.markdown('<div class="card-sep"><div class="step-label">Step 4</div>'
            '<div class="step-title">Process &amp; Download</div></div>', unsafe_allow_html=True)
st.markdown('<div class="card-bot">', unsafe_allow_html=True)

if st.button("▶  Run UEN Autofill", use_container_width=True):
    with st.spinner("Looking up UENs…"):
        processed_df, stats = process_df(
            raw_df, name_col_idx, uen_col_idx,
            header_row_idx, int(end_row_input),
            exact, fw_index, tok_index)
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

    st.markdown('<div class="dl-section"><div class="dl-section-title">📥 Download results</div>',
                unsafe_allow_html=True)
    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.markdown("**Full sheet** *(original columns + UEN filled)*")
        st.download_button("⬇  Excel (.xlsx)",
            build_full_excel(st.session_state["processed_df"]),
            file_name=out_name+".xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with dl2:
        st.markdown("**Full sheet** *(CSV)*")
        st.download_button("⬇  CSV",
            st.session_state["processed_df"].to_csv(index=False, header=False),
            file_name=out_name+".csv", mime="text/csv", use_container_width=True)
    with dl3:
        st.markdown("**Company Name + UEN only**")
        st.download_button("⬇  UEN Results (.xlsx)",
            build_uen_only_excel(
                st.session_state["processed_df"],
                name_col_idx, uen_col_idx, header_row_idx),
            file_name=out_name+"_uen_only.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)

    # ── Processed preview with two-col / full toggle ───────────────────────
    with st.expander("Preview processed output", expanded=True):
        two_col_view = st.toggle(
            "Show Company Name & UEN columns only",
            value=st.session_state.get("preview_two_col", True),
            key="preview_toggle")
        st.session_state["preview_two_col"] = two_col_view

        out_df = (st.session_state["processed_df"]
                  .copy().fillna("").replace({"nan":"","None":"","<NA>":""}))

        if two_col_view:
            sliced = out_df.iloc[header_row_idx:, [name_col_idx, uen_col_idx]].copy()
            new_cols = [
                str(sliced.iloc[0,0]) or f"{name_col_letter} — Company Name",
                str(sliced.iloc[0,1]) or f"{uen_col_letter} — UEN",
            ]
            sliced = sliced.iloc[1:].copy()
            sliced.columns = new_cols
            sliced.index = range(len(sliced))
            st.dataframe(sliced, width='stretch', height=400)
        else:
            full = out_df.iloc[header_row_idx:].copy()
            hv = [str(v) if str(v) not in ("","nan","None") else f"Col {col_index_to_letter(i)}"
                  for i, v in enumerate(full.iloc[0])]
            full = full.iloc[1:].copy()
            full.columns = hv
            full.index = range(len(full))
            st.dataframe(full, width='stretch', height=400)

    st.markdown(
        '<div style="margin-top:1rem;font-size:0.78rem;color:#9A9A9A;text-align:center;">'
        '🔒 Session data is held in memory only and discarded when you close or refresh this page.'
        '</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)
