# backend/modules/bs_generators.py
import os
import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional

import pandas as pd
from openpyxl.utils import column_index_from_string  # Ø­Ø±Ù Ø¹Ù…ÙˆØ¯ -> Ø±Ù‚Ù…

# =============================================
# Paths / Config
# =============================================
SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent
INPUT_FILE = str(BASE_DIR / 'consolidation_inputs' / 'Consolidated_TBs.xlsx')
RAW_SHEET = 'RawTB'
MAP_SHEET = 'FS_Descriptions BS'
OUTPUT_FILE = str(BASE_DIR / 'outputs' / 'Balance_Sheet.xlsx')

# Ø¹Ù…ÙˆØ¯ Ø§Ù„Ø±ØµÙŠØ¯ Ø§Ù„Ù†Ù‡Ø§Ø¦ÙŠ ÙÙŠ RawTB
BALANCE_COL = "L"  # ØºÙŠÙ‘Ø± Ø§Ù„Ø­Ø±Ù Ù„Ùˆ Ø§Ù„Ø¹Ù…ÙˆØ¯ Ø§ØªØºÙŠÙ‘Ø±

CONTROL_TOL = 1
MULTIPERIOD_SHEET = 'BS Multi-Period'
DETAIL_SHEET = 'Merged_Detail'
COMPANY_SHEET_TPL = 'BS {company} by Period'

COMPANY_ORDER: List[str] = [
    'HO', 'Training', 'Schools', 'University', 'Business solution', 'Wellness',
    'Standalone',
    'Smart', 'Faisaliyah', 'Roqi school', 'Riyadah', 'Franklin', 'Fast Lane',
    'Mazaya', 'Jobzella', 'Cairo KH', 'Linguaphone',
    'Combination'
]
STANDALONE_PARTS = ['HO', 'Training', 'Schools', 'University', 'Business solution', 'Wellness']

# =============================================
# Aliases
# =============================================
COLUMN_ALIASES = {
    'ref_to_fs': ['ref to fs', 'ref_to_fs', 'ref', 'code', 'fs ref', 'fs code', 'line code', 'fs_code', 'fsref'],
    'fs_line':   ['fs line', 'line', 'line name', 'description', 'desc', 'display name', 'caption'],
    'order':     ['order', 'seq', 'sequence', 'sort', 'position', '#', 'no', 'index', 'idx'],
    'section':   ['section', 'group', 'bucket', 'category', 'side'],
    'sign':      ['sign', 'mult', 'multiplier'],
    'note_ref':  ['note ref', 'note', 'note_ref', 'note code']
}
TB_ALIASES = {
    'Statement': ['statement', 'stmt', 'fs', 'type'],
    'Company':   ['company', 'company name', 'entity', 'legal entity', 'co'],
    'REF to FS': ['ref to fs', 'fs ref', 'ref', 'fs code', 'fs_ref', 'fs code ref', 'fs_code', 'fsref'],
    'Reporting Date': ['reporting date', 'report date', 'date'],
}

# =============================================
# Helpers
# =============================================
def _clean_header(s: str) -> str:
    return str(s).replace('\u00A0', ' ').strip()

def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_clean_header(c) for c in df.columns]
    return df

def _rename_with_aliases(df: pd.DataFrame, alias_map: dict) -> pd.DataFrame:
    df = df.copy()
    lower_map = {c.lower(): c for c in df.columns}
    ren = {}
    for canon, aliases in alias_map.items():
        for a in [canon] + aliases:
            key = a.lower()
            if key in lower_map:
                ren[lower_map[key]] = canon
                break
    return df.rename(columns=ren)

def _ensure_columns(df: pd.DataFrame, must_have: list, where: str):
    missing = [c for c in must_have if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in '{where}': {missing} | Available: {list(df.columns)}")

def _norm_txt_spaces_dashes(s: str) -> str:
    return " ".join(
        str(s).lower()
              .replace('\u00A0', ' ')
              .replace('â€“', ' ').replace('â€”', ' ').replace('-', ' ')
              .split()
    )

# ---------- Company normalization ----------
def _company_alias_map():
    return {
        'business solution': ['business solutions', 'business Solution', 'Business Solutions', 'business solutions '],
        'smart': ['smart ', 'SMART', 'Smart ', ' smart'],
        'roqi school': ['Roqi School', 'roqi  school', 'Roqi  school'],
        'fast lane': ['Fastlane', 'FAST LANE', 'Fast  Lane'],
        'cairo kh': ['CairoKH', 'Cairo  KH', 'Cairo-KH'],
    }

def _normalize_company_name(name: str) -> str:
    base = str(name).strip()
    key = " ".join(base.lower().split())
    aliases = _company_alias_map()
    for canon, vars_ in aliases.items():
        if key == canon or key in [v.lower().strip() for v in vars_]:
            return canon.title() if canon != 'business solution' else 'Business solution'
    return base

def _normalize_company_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    new_cols = {}
    for c in df.columns:
        new_cols[c] = c if c == 'REF to FS' else _normalize_company_name(c)
    df.columns = [new_cols.get(c, c) for c in df.columns]
    if len(set(df.columns)) != len(df.columns):
        df = df.groupby(axis=1, level=0).sum(numeric_only=True)
    return df

# ---------- Period utilities ----------
def _build_period_column(tb: pd.DataFrame) -> Tuple[pd.Series, Dict[int, str], pd.Series]:
    if 'Reporting Date' not in tb.columns:
        return pd.Series([pd.NA] * len(tb)), {}, pd.Series(pd.NaT, index=tb.index)
    dt = pd.to_datetime(tb['Reporting Date'], errors='coerce', dayfirst=True)
    period = dt.dt.strftime('%b %Y').str.upper()
    labels: Dict[int, str] = {}
    if not dt.dropna().empty:
        df_y = pd.DataFrame({'dt': dt})
        df_y['year'] = df_y['dt'].dt.year
        for y, g in df_y.dropna().groupby('year'):
            max_dt = g['dt'].max()
            if pd.notna(max_dt):
                labels[int(y)] = max_dt.strftime('%b %Y').upper()
    return period, labels, dt

# =============================================
# Ù‚Ø±Ø§Ø¡Ø© Ø§Ù„Ø¨ÙŠØ§Ù†Ø§Øª: Ù‚ÙŠÙ…Ø© Ø§Ù„Ø­Ø³Ø§Ø¨ = Ø¹Ù…ÙˆØ¯ L
# =============================================
def _load_data_from(input_file: str) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[int, str]]:
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")

    tb = pd.read_excel(input_file, sheet_name=RAW_SHEET)
    tb = _normalize_headers(tb)
    tb = _rename_with_aliases(tb, TB_ALIASES)
    _ensure_columns(tb, ['Statement', 'Company', 'REF to FS'], RAW_SHEET)

    tb['Period'], year_labels, dt = _build_period_column(tb)
    tb['_dt'] = dt

    idx = column_index_from_string(BALANCE_COL) - 1  # 0-based
    if idx >= tb.shape[1]:
        raise ValueError(f"RawTB doesn't have column {BALANCE_COL}. It has only {tb.shape[1]} columns.")
    tb['Value'] = pd.to_numeric(tb.iloc[:, idx], errors='coerce').fillna(0.0)

    mapping = pd.read_excel(input_file, sheet_name=MAP_SHEET)
    mapping = _normalize_headers(mapping)

    # Ø·Ø¨Ù‘Ù‚ aliases Ù„Ù„Ø§Ø¹Ù…Ø¯Ø©
    def _normalize_columns(df: pd.DataFrame, alias_map: dict) -> pd.DataFrame:
        new_cols = {}
        lower_map = {c.lower().strip(): c for c in df.columns}
        for canon, aliases in alias_map.items():
            found = None
            for a in [canon] + aliases:
                key = a.lower().strip()
                if key in lower_map:
                    found = lower_map[key]; break
            if found is not None:
                new_cols[found] = canon
        return df.rename(columns=new_cols)

    mapping = _normalize_columns(mapping, COLUMN_ALIASES)

    if 'side' in mapping.columns:
        if 'section' not in mapping.columns:
            mapping['section'] = mapping['side']
        else:
            blank = mapping['section'].astype(str).str.strip().eq('')
            mapping.loc[blank, 'section'] = mapping.loc[blank, 'side']
        mapping.drop(columns=['side'], inplace=True)

    if 'order' not in mapping.columns:
        mapping['order'] = range(1, len(mapping) + 1)

    _ensure_columns(mapping, ['ref_to_fs', 'fs_line', 'order'], MAP_SHEET)

    # âœ… Ø¥ØµÙ„Ø§Ø­ sign: Ø¯Ø§ÙŠÙ…Ù‹Ø§ Series
    if 'sign' in mapping.columns:
        mapping['sign'] = pd.to_numeric(mapping['sign'], errors='coerce').fillna(1)
    else:
        mapping['sign'] = 1  # ÙŠØ¨Ù‚Ù‰ Ø¹Ù…ÙˆØ¯ Ø«Ø§Ø¨Øª Ù‚ÙŠÙ…ØªÙ‡ 1

    if 'section' not in mapping.columns:
        mapping['section'] = ''
    if 'note_ref' not in mapping.columns:
        mapping['note_ref'] = pd.NA

    mapping['order'] = pd.to_numeric(mapping['order'], errors='coerce')
    mapping = mapping.dropna(subset=['order']).sort_values('order').reset_index(drop=True)
    mapping['section'] = mapping['section'].ffill()

    return tb, mapping, year_labels

# =============================================
# Ø¨Ø§Ù‚Ù‰ Ø§Ù„Ù…Ù†Ø·Ù‚ (Pivot/Standalone/Combination/Totals)
# =============================================
def _sum_cols(df: pd.DataFrame, cols):
    present = [c for c in cols if c in df.columns]
    return df[present].sum(axis=1) if present else pd.Series(0.0, index=df.index)

def _combination(df: pd.DataFrame):
    cols = [c for c in df.columns if c not in ('Standalone', 'Combination')]
    return df[cols].sum(axis=1) if cols else pd.Series(0.0, index=df.index)

def _compute_totals_inplace(df: pd.DataFrame, value_cols: List[str]) -> None:
    def is_total(s: str) -> bool:
        return 'total' in str(s).casefold()
    df['_section_norm'] = df['section'].fillna('').astype(str)
    for _, gidx in df.groupby('_section_norm').groups.items():
        idx = list(gidx)
        start = 0
        for rel, row_idx in enumerate(idx):
            if is_total(df.at[row_idx, 'fs_line']):
                block_rows = [idx[i] for i in range(start, rel)]
                if block_rows:
                    mask_non_total = ~df.loc[block_rows, 'fs_line'].astype(str).str.casefold().str.contains('total')
                    block_rows = list(pd.Index(block_rows)[mask_non_total.values])
                sums = df.loc[block_rows, value_cols].sum(numeric_only=True) if block_rows else {c: 0.0 for c in value_cols}
                for c in value_cols:
                    df.at[row_idx, c] = sums.get(c, 0.0)
                start = rel + 1
    df.drop(columns=['_section_norm'], inplace=True)

def _apply_grand_totals(df: pd.DataFrame, value_cols: List[str]) -> None:
    aliases = {
        'total non current liabilities': ['total non current liability', 'total non-current liabilities', 'total non-current liability'],
        'total current liabilities': ['total current liability'],
        'total liabilities': ['total liability'],
        'total equity': ["total owner's equity", "total owners equity", "total owner equity", "total shareholders' equity", "total shareholders equity"],
        'total liabilities and equity': ['total liability and owner equity', 'total liabilities and owner equity', 'total liability and equity', 'total liabilities & equity', 'total liability & owner equity'],
        'assets classified as held for sale': ['assets classified as held-for-sale'],
        'total non current assets': ['total non-current assets'],
        'total current assets': ['total current asset'],
    }
    names_norm = [_norm_txt_spaces_dashes(x) for x in df['fs_line']]

    def find_rows(key: str):
        pats = [key] + aliases.get(key, [])
        rows = []
        for p in pats:
            target = _norm_txt_spaces_dashes(p)
            rx = re.compile(rf'\b{re.escape(target)}\b')
            for i, nm in enumerate(names_norm):
                if rx.search(nm):
                    rows.append(i)
        return list(dict.fromkeys(rows))

    def find_single_row(key: str):
        rows = find_rows(key)
        return [rows[-1]] if rows else []

    t_assets_parts = ['total non current assets', 'total current assets', 'assets classified as held for sale']
    t_assets_rows = find_rows('total assets')
    if t_assets_rows:
        comp_idx = []
        for p in t_assets_parts: comp_idx.extend(find_rows(p))
        if comp_idx:
            comp_sum = df.loc[comp_idx, value_cols].sum(numeric_only=True)
            for r in t_assets_rows:
                for c in value_cols: df.at[r, c] = comp_sum.get(c, 0.0)

    t_liab_rows = find_rows('total liabilities')
    if t_liab_rows:
        comp_idx = find_rows('total non current liabilities') + find_rows('total current liabilities')
        if comp_idx:
            comp_sum = df.loc[comp_idx, value_cols].sum(numeric_only=True)
            for r in t_liab_rows:
                for c in value_cols: df.at[r, c] = comp_sum.get(c, 0.0)

    t_le_rows = find_rows('total liabilities and equity')
    if t_le_rows:
        liab_idx = find_single_row('total liabilities')
        eq_idx   = find_single_row('total equity')
        comp_idx = liab_idx + eq_idx
        if comp_idx:
            comp_sum = df.loc[comp_idx, value_cols].sum(numeric_only=True)
            for r in t_le_rows:
                for c in value_cols: df.at[r, c] = comp_sum.get(c, 0.0)

    ctrl_rows = find_rows('control')
    if ctrl_rows and t_assets_rows and t_le_rows:
        a_sum = df.loc[find_single_row('total assets'), value_cols].sum(numeric_only=True)
        b_sum = df.loc[t_le_rows, value_cols].sum(numeric_only=True)
        diff = a_sum.subtract(b_sum, fill_value=0.0)
        for r in ctrl_rows:
            for c in value_cols: df.at[r, c] = diff.get(c, 0.0)

def _apply_total_equity_from_parts(df: pd.DataFrame, value_cols: List[str]) -> None:
    names = df['fs_line'].fillna('').map(_norm_txt_spaces_dashes)
    eq_attr_mask = names.str.contains(r'\btotal equity attributable to the shareholders\b') | names.str.contains(r'\btotal equity attributable to the shareholders of the co\b')
    nci_mask     = names.str.fullmatch(r'non[-\s]?controlling interests') | names.str.contains(r'\bnon[-\s]?controlling interests\b')
    total_eq_mask = names.str.fullmatch(r'total equity') | names.str.contains(r'\btotal equity\b')
    eq_attr_rows = list(df.index[eq_attr_mask])
    nci_rows     = list(df.index[nci_mask])
    total_eq_rows = list(df.index[total_eq_mask])
    if not total_eq_rows or not eq_attr_rows:
        return
    eq_attr = df.loc[eq_attr_rows, value_cols].sum(numeric_only=True)
    nci     = df.loc[nci_rows, value_cols].sum(numeric_only=True) if nci_rows else 0.0
    total   = eq_attr.add(nci, fill_value=0.0)
    for r in total_eq_rows:
        for c in value_cols:
            df.at[r, c] = float(total.get(c, 0.0))

def _flip_liab_equity_signs(df: pd.DataFrame, value_cols: List[str]) -> None:
    SECTION_LIAB_EQ = {'owner equity', 'equity', "owners' equity", 'owners equity', "shareholders' equity', 'shareholders equity",
                       'capital', 'liability', 'liabilities', 'current liability', 'current liabilities',
                       'non-current liability', 'non current liability', 'non-current liabilities', 'non current liabilities'}
    sec_norm = df['section'].fillna('').map(_norm_txt_spaces_dashes)
    mask_by_section = sec_norm.isin(SECTION_LIAB_EQ)
    KW_FALLBACK = ['owner equity', 'equity', "owners' equity", 'owners equity', "shareholders' equity", 'shareholders equity', 'capital',
                   'liability', 'liabilities', 'current liability', 'non-current liability']
    line_norm = df['fs_line'].fillna('').map(_norm_txt_spaces_dashes)
    mask_by_kw = sec_norm.eq('') & line_norm.apply(lambda s: any(k in s for k in KW_FALLBACK))
    mask = mask_by_section | mask_by_kw
    if mask.any():
        df.loc[mask, value_cols] = df.loc[mask, value_cols] * -1

# =============================================
# Ø§Ù„Ø¨Ù†Ø§Ø¡/Ø§Ù„ØªØµØ¯ÙŠØ±
# =============================================
def build_excel_ui(
    input_file: str,
    output_file: str,
    periods: Optional[List[str]] = None,
    label_mode: str = "PERIOD",
    company: Optional[str] = None,
):
    tb, mapping, year_labels = _load_data_from(input_file)

    tb['Company'] = tb['Company'].astype(str).map(_normalize_company_name)
    tb_bs = tb[tb['Statement'].astype(str).str.upper().eq('BS')].copy()

    avail = tb_bs[['Period', '_dt']].dropna().drop_duplicates().sort_values('_dt', ascending=False)
    available_periods = [str(p).upper() for p in avail['Period']]

    if periods is not None:
        wanted = [str(p).upper() for p in periods]
        selected_periods = [p for p in wanted if p in available_periods]
        if not selected_periods:
            raise ValueError(f"No matching periods found. Wanted={wanted}, Available={available_periods}")
    else:
        selected_periods = available_periods[:2]

    only_company = _normalize_company_name(company) if company else None

    def _build_pivot_for_period(tb_bs_, mapping_, period_, only_company_=None):
        cur = tb_bs_[tb_bs_['Period'].astype(str).str.upper().eq(period_)].copy()
        if cur.empty:
            base = pd.DataFrame({'ref_to_fs': mapping_['ref_to_fs'].unique()})
            if only_company_:
                base[period_] = 0.0
            else:
                for c in COMPANY_ORDER:
                    base[f"{c} {period_}"] = 0.0
            return base

        agg = (cur.groupby(['REF to FS', 'Company'], as_index=False)['Value'].sum())
        pvt = agg.pivot(index='REF to FS', columns='Company', values='Value').fillna(0.0)
        pvt = _normalize_company_columns(pvt)
        pvt = pvt.reset_index().rename(columns={'REF to FS': 'ref_to_fs'})

        if only_company_:
            col = only_company_
            if col not in pvt.columns:
                pvt[col] = 0.0
            out = pvt[['ref_to_fs', col]].copy().rename(columns={col: period_})
            return out
        else:
            for comp in COMPANY_ORDER:
                if comp in ('Standalone', 'Combination'):
                    continue
                if comp not in pvt.columns:
                    pvt[comp] = 0.0
            pvt['Standalone']  = _sum_cols(pvt, STANDALONE_PARTS)
            pvt['Combination'] = _combination(pvt)
            cols = ['ref_to_fs'] + [c for c in COMPANY_ORDER if c in pvt.columns]
            pvt = pvt[cols].copy()
            rename = {c: f"{c} {period_}" for c in cols if c != 'ref_to_fs'}
            pvt = pvt.rename(columns=rename)
            return pvt

    merged_all = mapping.copy()
    if 'REF to FS' in merged_all.columns:
        merged_all = merged_all.rename(columns={'REF to FS': 'ref_to_fs'})

    block_cols_per_period = {}
    for per in selected_periods:
        pvt = _build_pivot_for_period(tb_bs, mapping, per, only_company_=only_company)
        merged_all = merged_all.merge(pvt, how='left', on='ref_to_fs')
        if only_company:
            block_cols_per_period[per] = [per] if per in merged_all.columns else []
        else:
            block_cols = [f"{c} {per}" for c in COMPANY_ORDER if f"{c} {per}" in merged_all.columns]
            block_cols_per_period[per] = block_cols

    numeric_cols = [c for per in selected_periods for c in block_cols_per_period.get(per, [])]
    numeric_cols = [c for c in numeric_cols if c in merged_all.columns]

    if numeric_cols:
        # sign Ø£ØµØ¨Ø­ Series Ù…Ø¤ÙƒØ¯
        merged_all['sign'] = pd.to_numeric(merged_all['sign'], errors='coerce').fillna(1.0)
        for c in numeric_cols:
            merged_all[c] = pd.to_numeric(merged_all[c], errors='coerce').fillna(0.0) * merged_all['sign']
            merged_all[c] = merged_all[c].round(0)

        _flip_liab_equity_signs(merged_all, value_cols=numeric_cols)
        merged_all = merged_all.sort_values(['order', 'ref_to_fs'], kind='mergesort').reset_index(drop=True)
        _compute_totals_inplace(merged_all, value_cols=numeric_cols)
        _apply_total_equity_from_parts(merged_all, value_cols=numeric_cols)
        _apply_grand_totals(merged_all, value_cols=numeric_cols)

    out_cols = ['ref_to_fs', 'fs_line'] + numeric_cols
    out_df = merged_all[out_cols].copy()

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    sheet_name = COMPANY_SHEET_TPL.format(company=only_company) if only_company else MULTIPERIOD_SHEET
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        out_df.to_excel(writer, sheet_name=sheet_name, index=False)

    return {"periods": selected_periods, "output_file": output_file, "company": only_company or "ALL"}

# ====== APIs Ù…Ø³Ø§Ø¹Ø¯Ø© ======
def list_periods(input_file: str) -> List[str]:
    tb, _, _ = _load_data_from(input_file)
    tb_bs = tb[tb['Statement'].astype(str).str.upper().eq('BS')].copy()
    avail = tb_bs[['Period', '_dt']].dropna().drop_duplicates().sort_values('_dt', ascending=False)
    return [str(p) for p in avail['Period']]

def list_accounts(input_file: str, company: Optional[str] = None):
    tb, mapping, _ = _load_data_from(input_file)
    mapping_use = mapping[['ref_to_fs', 'fs_line', 'order']].drop_duplicates()
    if company:
        tb['Company'] = tb['Company'].astype(str).map(_normalize_company_name)
        company = _normalize_company_name(company)
        use = tb[tb['Company'].eq(company)][['REF to FS']].drop_duplicates()
        mapping_use = mapping_use.merge(use, how='inner', left_on='ref_to_fs', right_on='REF to FS')
    mapping_use = mapping_use.sort_values('order', kind='mergesort')
    return [{"ref": str(r['ref_to_fs']), "fs_line": str(r['fs_line'])} for _, r in mapping_use.iterrows()]

def series_by_account(input_file: str, company: Optional[str], ref: Optional[str], fs_line: Optional[str], periods: Optional[list] = None):
    tb, mapping, _ = _load_data_from(input_file)
    tb['Company'] = tb['Company'].astype(str).map(_normalize_company_name)
    tb_bs = tb[tb['Statement'].astype(str).str.upper().eq('BS')].copy()

    if company:
        company = _normalize_company_name(company)
        tb_bs = tb_bs[tb_bs['Company'].astype(str).eq(company)]

    mapping_use = mapping[['ref_to_fs', 'fs_line']].drop_duplicates()
    tb_bs = tb_bs.merge(mapping_use, how='left', left_on='REF to FS', right_on='ref_to_fs')

    if ref:
        tb_bs = tb_bs[tb_bs['ref_to_fs'].astype(str).str.casefold() == str(ref).strip().casefold()]
    elif fs_line:
        tb_bs = tb_bs[tb_bs['fs_line'].astype(str).str.casefold() == str(fs_line).strip().casefold()]
    else:
        raise ValueError("Either ref or fs_line must be provided")

    if periods:
        want = set([str(p).upper() for p in periods])
        tb_bs = tb_bs[tb_bs['Period'].astype(str).str.upper().isin(want)]

    agg = (tb_bs.groupby(['Period', '_dt'], as_index=False)['Value'].sum()).sort_values('_dt')
    return [{'period': str(r['Period']), 'value': float(r['Value'])} for _, r in agg.iterrows()]
# =========================
# SIMPLE EXPORTS FOR API
# Ø¶ÙŠÙ Ø§Ù„Ø¨Ù„ÙˆÙƒ Ø¯Ù‡ ÙÙŠ Ø¢Ø®Ø± modules/bs_generators.py
# =========================
from typing import List
import pandas as pd
from pathlib import Path

# Ø«ÙˆØ§Ø¨Øª Ø¢Ù…Ù†Ø© Ù„Ùˆ Ù…Ø´ Ù…ØªØ¹Ø§Ø±Ù Ø¹Ù„ÙŠÙ‡Ø§ ÙÙˆÙ‚
try:
    INPUT_FILE
except NameError:
    SCRIPT_DIR = Path(__file__).resolve().parent
    BASE_DIR = SCRIPT_DIR.parent
    INPUT_FILE = str(BASE_DIR / 'consolidation_inputs' / 'Consolidated_TBs.xlsx')

def _norm_co_compat(name: str) -> str:
    """Ø§Ø³ØªØ®Ø¯Ù… Ù†ÙˆØ±Ù…Ø§Ù„Ø§ÙŠØ²Ø± Ø§Ù„Ù…Ø´Ø±ÙˆØ¹ Ù„Ùˆ Ù…ÙˆØ¬ÙˆØ¯ØŒ ÙˆØ¥Ù„Ø§ Ø¨Ø³ Ù†Ø¸Ù‘Ù Ø§Ù„Ø§Ø³Ù…."""
    try:
        return _normalize_company_name(name)  # Ù…ÙˆØ¬ÙˆØ¯ Ø¹Ù†Ø¯Ùƒ ÙÙŠ Ø§Ù„Ù…Ù„Ù Ø§Ù„Ø£ØµÙ„ÙŠ
    except NameError:
        return " ".join(str(name).strip().split())

def list_companies_and_sectors(input_file: str = INPUT_FILE):
    """
    ØªØ±Ø¬Ø¹ (companies, sectors). Ø­Ø§Ù„ÙŠØ§Ù‹ sectors = Ù†ÙØ³ Ø§Ù„Ø´Ø±ÙƒØ§Øª (placeholder).
    Ø¨ØªØ´ØªØºÙ„ Ø³ÙˆØ§Ø¡ _load_data_from Ù…ÙˆØ¬ÙˆØ¯Ø© Ø£Ùˆ Ù„Ø§.
    """
    df = None
    # Ø¬Ø±Ù‘Ø¨ ØªÙ‚Ø±Ø§ Ø¹Ø¨Ø± Ø§Ù„Ø¯Ø§Ù„Ø© Ø§Ù„Ø¯Ø§Ø®Ù„ÙŠØ© Ù„Ùˆ Ù…ÙˆØ¬ÙˆØ¯Ø©
    try:
        tb, _, _ = _load_data_from(input_file)  # Ù…ÙˆØ¬ÙˆØ¯Ø© Ø¹Ù†Ø¯Ùƒ Ø£ØµÙ„Ø§Ù‹
        df = tb
    except Exception:
        pass

    if df is None:
        # fallback Ù…Ø¨Ø§Ø´Ø± Ù„ÙˆØ±Ù‚Ø© RawTB
        df = pd.read_excel(input_file, sheet_name='RawTB')

    if 'Company' not in df.columns:
        raise ValueError("Column 'Company' not found in RawTB")

    comps = sorted({
        _norm_co_compat(c)
        for c in df['Company'].dropna().astype(str)
        if str(c).strip() != ''
    })
    sectors = sorted(set(comps))  # Ù„Ø­Ø¯ Ù…Ø§ ÙŠØ¨Ù‚Ù‰ ÙÙŠ ØªØµÙ†ÙŠÙ Ø­Ù‚ÙŠÙ‚ÙŠ
    return comps, sectors

# (Ø§Ø®ØªÙŠØ§Ø±ÙŠ/ÙˆÙ‚Ø§Ø¦ÙŠ) Ù†ÙˆÙØ± list_periods Ù„Ùˆ Ø¨ÙÙ†ÙŠØª Ù†Ø³Ø®Ø© Ù‚Ø¯ÙŠÙ…Ø© Ø¨Ø¯ÙˆÙ†Ù‡Ø§
def list_periods(input_file: str = INPUT_FILE) -> List[str]:
    try:
        tb, _, _ = _load_data_from(input_file)
        tb_bs = tb[tb['Statement'].astype(str).str.upper().eq('BS')].copy()
        avail = (
            tb_bs[['Period', '_dt']]
            .dropna()
            .drop_duplicates()
            .sort_values('_dt', ascending=False)
        )
        return [str(p) for p in avail['Period']]
    except Exception:
        df = pd.read_excel(input_file, sheet_name='RawTB')
        df['Reporting Date'] = pd.to_datetime(df.get('Reporting Date'), errors='coerce', dayfirst=True)
        period = df['Reporting Date'].dt.strftime('%b %Y').str.upper()
        avail = (
            pd.DataFrame({'Period': period, '_dt': df['Reporting Date']})
            .dropna()
            .drop_duplicates()
            .sort_values('_dt', ascending=False)
        )
        return [str(p) for p in avail['Period']]

