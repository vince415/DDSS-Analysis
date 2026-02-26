import streamlit as st
import pandas as pd
import openpyxl
import io
import hashlib
from datetime import datetime, timedelta
import plotly.graph_objects as go
import plotly.express as px
import re

st.set_page_config(
    page_title="DDSS Forecast Analysis",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ═══════════════════════════════════════════════════════════════
#  Utility Functions
# ═══════════════════════════════════════════════════════════════

def normalize(s):
    """Normalize string: strip + uppercase"""
    return str(s).strip().upper() if s is not None else ""


def extract_mpa_year_week(filename):
    stem = re.sub(r'\.xlsx?$', '', filename, flags=re.IGNORECASE).strip()

    m = re.search(r'^(.*?)\s*(\d{4})\s*wk\s*(\d+)\s*$', stem, re.IGNORECASE)
    if m:
        mpa = m.group(1).strip().upper() if m.group(1).strip() else None
        year = int(m.group(2))
        week = int(m.group(3))
        return mpa, year, week

    m = re.search(r'^(.*?)\s*[_\-]?\s*wk\s*(\d+)\s*$', stem, re.IGNORECASE)
    if m:
        mpa = m.group(1).strip().upper() if m.group(1).strip() else None
        week = int(m.group(2))
        return mpa, None, week

    return None, None, None


def create_week_sort_key(year, week):
    if year is not None:
        return year * 100 + week
    else:
        return week  # Old format, will sort by week only


def find_all_ddss_sheets(wb):
    """Find all DDSS-like sheets (case-insensitive)"""
    results = []
    for name in wb.sheetnames:
        norm_name = name.strip().upper()
        if norm_name == 'DDSS' or norm_name.startswith("DDSS W"):
            results.append((name, wb[name]))
    return results


def detect_header_cols(row):
    """Detect column indices from header row (case-insensitive + aliases)"""
    norm = [normalize(c) for c in row]

    def find(*candidates):
        for c in candidates:
            cn = normalize(c)
            for i, v in enumerate(norm):
                if v == cn:
                    return i
        return None

    col_mpa = find('MPA')
    col_type = find('TYPE', 'DETAILS')
    col_part = find('CONSIGN PN', 'PART NUMBER')
    col_desc = find('DATA DESCRIPTION')
    col_oh = find('ON HAND (RM)', 'ON HAND (FG)', 'ON HAND')

    if col_mpa is None or col_part is None or col_desc is None:
        return None

    col_date = next((i for i, v in enumerate(row) if isinstance(v, datetime)), None)
    if col_date is None:
        return None

    return {
        'mpa': col_mpa, 'type': col_type, 'part': col_part,
        'desc': col_desc, 'oh': col_oh, 'date0': col_date
    }


# ═══════════════════════════════════════════════════════════════
#  File Parsing (Cached by file content hash)
# ═══════════════════════════════════════════════════════════════

@st.cache_data(show_spinner=False)
def parse_one_file(file_bytes: bytes, filename: str):
    """Parse a single Excel file, extract data from all DDSS sheets"""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ddss_sheets = find_all_ddss_sheets(wb)
        if not ddss_sheets:
            return None

        mpa_file, year, week = extract_mpa_year_week(filename)
        if week is None:
            return None

        # Create sort key for proper ordering
        week_sort_key = create_week_sort_key(year, week)

        all_records = []

        for sheet_name, sheet in ddss_sheets:
            records = []
            cur_part = cur_type = cols = start_date = None

            for row in sheet.iter_rows(values_only=True):
                if not any(row):
                    continue

                det = detect_header_cols(row)
                if det is not None:
                    cols = det
                    sd = row[cols['date0']]
                    start_date = sd if isinstance(sd, datetime) else None
                    if start_date is None:
                        cols = None
                    cur_part = cur_type = None
                    continue

                if not cols or not start_date:
                    continue

                L = len(row)

                if cols['part'] < L:
                    rp = row[cols['part']]
                    if rp and normalize(rp) not in ('CONSIGN PN', 'PART NUMBER', ''):
                        cur_part = str(rp).strip()

                if cols['type'] is not None and cols['type'] < L:
                    rt = row[cols['type']]
                    if rt and normalize(rt) not in ('TYPE', 'DETAILS', ''):
                        cur_type = str(rt).strip()

                if cols['desc'] >= L:
                    continue

                rd = row[cols['desc']]
                if not cur_part or not rd or normalize(rd) in ('DATA DESCRIPTION', ''):
                    continue

                mpa = mpa_file or (normalize(row[cols['mpa']])
                                   if cols['mpa'] < L and row[cols['mpa']] else None)

                # Create week label with year if available
                if year is not None:
                    week_label = f"{year}Wk{week}"
                else:
                    week_label = f"Wk{week}"

                base = {
                    'Week': week,
                    'Year': year,
                    'WeekSortKey': week_sort_key,
                    'WeekLabel': week_label,
                    'MPA': mpa,
                    'Sheet': sheet_name,
                    'Filename': filename,
                    'Type': cur_type,
                    'Consign_PN': cur_part,
                    'Data_Description': str(rd).strip(),
                }

                if cols['oh'] is not None and cols['oh'] < L:
                    oh = row[cols['oh']]
                    if isinstance(oh, (int, float)):
                        records.append({
                            **base,
                            'Date': start_date - timedelta(days=7),
                            'Column_Type': 'On hand',
                            'Value': float(oh),
                        })

                d0 = cols['date0']
                for off, ci in enumerate(range(d0, L)):
                    val = row[ci]
                    if isinstance(val, (int, float)):
                        records.append({
                            **base,
                            'Date': start_date + timedelta(days=off * 7),
                            'Column_Type': 'Forecast',
                            'Value': float(val),
                        })

            all_records.extend(records)

        return pd.DataFrame(all_records) if all_records else None

    except Exception:
        return None


def load_all_files(uploaded_files):
    """Load all uploaded files, use parse_one_file cache"""
    fingerprint = hashlib.md5(
        b''.join(f.name.encode() + str(f.size).encode() for f in uploaded_files)
    ).hexdigest()

    if (st.session_state.get('_file_fp') == fingerprint
            and 'combined_df' in st.session_state):
        return st.session_state['combined_df'], st.session_state.get('_parse_status', [])

    all_data, status = [], []
    for f in uploaded_files:
        f.seek(0)
        fb = f.read()
        df = parse_one_file(fb, f.name)
        if df is not None and len(df):
            all_data.append(df)
            status.append((f.name, '✓ Success', len(df)))
        else:
            status.append((f.name, '✗ Failed', 0))

    combined = pd.concat(all_data, ignore_index=True) if all_data else None
    st.session_state['combined_df'] = combined
    st.session_state['_file_fp'] = fingerprint
    st.session_state['_parse_status'] = status
    return combined, status


# ═══════════════════════════════════════════════════════════════
#  Select All / Clear + Multiselect
# ═══════════════════════════════════════════════════════════════

def make_filter(label: str, options: list, key: str) -> list:
    """Multiselect with Select All / Clear buttons"""
    if key not in st.session_state:
        st.session_state[key] = list(options)

    valid = set(options)
    st.session_state[key] = [v for v in st.session_state[key] if v in valid]

    col_a, col_b = st.columns(2)
    if col_a.button("Select All", key=f"_a_{key}", use_container_width=True):
        st.session_state[key] = list(options)
    if col_b.button("Clear", key=f"_b_{key}", use_container_width=True):
        st.session_state[key] = []

    return st.multiselect(label, options=options, key=key,
                          label_visibility="collapsed")


# ═══════════════════════════════════════════════════════════════
#  Build Pivot Tables (Split into two tables)
# ═══════════════════════════════════════════════════════════════

def is_wos_related(text: str) -> bool:
    """Check if text contains 'wos' (case-insensitive)"""
    return 'wos' in text.lower() if text else False


def build_pivot_tables(desc_agg: pd.DataFrame, metric_label: str = ''):
    """
    Build TWO pivot tables:
    1. Weekly data table (sorted by WeekSortKey)
    2. Delta table (consecutive week differences)
    """
    oh_df = desc_agg[desc_agg['Column_Type'] == 'On hand']
    fct_df = desc_agg[desc_agg['Column_Type'] == 'Forecast']

    if fct_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    # Get unique weeks with their sort keys and labels
    week_info = fct_df[['WeekSortKey', 'WeekLabel']].drop_duplicates()
    week_info = week_info.sort_values('WeekSortKey', ascending=False)  # Newest first

    # Pivot by WeekSortKey for aggregation, but display with WeekLabel
    pt = fct_df.pivot_table(index='WeekSortKey', columns='Date',
                            values='Value', aggfunc='sum').T.sort_index()

    # Create mapping from WeekSortKey to WeekLabel
    key_to_label = dict(zip(week_info['WeekSortKey'], week_info['WeekLabel']))

    # Table 1: Weekly data (newest at top)
    weekly_rows = []
    for sort_key in week_info['WeekSortKey']:
        label = key_to_label[sort_key]
        r = {'Metric': label}

        # Get On hand for this week
        oh_rows = oh_df[oh_df['WeekSortKey'] == sort_key]
        oh_val = oh_rows['Value'].sum() if not oh_rows.empty else None
        r['On hand'] = oh_val

        # Get forecast values for each date
        for dt in pt.index:
            if sort_key in pt.columns:
                v = pt.loc[dt, sort_key]
                date_str = dt.strftime('%Y-%m-%d')
                r[date_str] = v if pd.notna(v) else None
        weekly_rows.append(r)

    weekly_table = pd.DataFrame(weekly_rows)

    # Table 2: Delta table (consecutive week differences, newest first)
    delta_rows = []
    sort_keys = list(week_info['WeekSortKey'])

    for i in range(len(sort_keys) - 1):
        newer_key = sort_keys[i]
        older_key = sort_keys[i + 1]
        newer_label = key_to_label[newer_key]
        older_label = key_to_label[older_key]

        r = {'Metric': f'{newer_label}→{older_label}'}

        # Delta for On hand
        newer_oh = oh_df[oh_df['WeekSortKey'] == newer_key]
        older_oh = oh_df[oh_df['WeekSortKey'] == older_key]

        if not newer_oh.empty and not older_oh.empty:
            r['On hand'] = newer_oh['Value'].sum() - older_oh['Value'].sum()
        else:
            r['On hand'] = None

        # Delta for each date
        for dt in pt.index:
            newer_val = pt.loc[dt, newer_key] if newer_key in pt.columns else None
            older_val = pt.loc[dt, older_key] if older_key in pt.columns else None
            date_str = dt.strftime('%Y-%m-%d')

            if pd.notna(newer_val) and pd.notna(older_val):
                r[date_str] = newer_val - older_val
            else:
                r[date_str] = None
        delta_rows.append(r)

    delta_table = pd.DataFrame(delta_rows)

    return weekly_table, delta_table


def style_table(df: pd.DataFrame, is_wos: bool = False):
    """Apply styling to dataframe"""

    def apply_color(val):
        if isinstance(val, (int, float)) and val < 0:
            return 'color:#d62728;font-weight:bold'
        return ''

    styled = df.style.applymap(apply_color, subset=[c for c in df.columns if c != 'Metric'])
    numeric_cols = [c for c in df.columns if c != 'Metric']

    if is_wos:
        format_dict = {col: '{:.2f}' for col in numeric_cols}
    else:
        format_dict = {col: '{:.0f}' for col in numeric_cols}

    styled = styled.format(format_dict, na_rep='')
    return styled


# ═══════════════════════════════════════════════════════════════
#  Main Program
# ═══════════════════════════════════════════════════════════════

def main():
    # ────────── Sidebar ──────────────────────────────────────
    with st.sidebar:
        st.markdown("## 📁 Upload Files")
        uploaded = st.file_uploader(
            "Select Excel files (batch supported, naming: MPA YYYYwk00.xlsx)",
            type=['xlsx'],
            accept_multiple_files=True,
            label_visibility="collapsed",
        )

        if not uploaded:
            st.info("Please upload files\n\nFormat: `MPA YYYYwk00.xlsx`")
            st.session_state.pop('combined_df', None)
            return

        with st.spinner("Parsing files..."):
            combined, status = load_all_files(uploaded)

        if combined is None:
            st.error("Parse failed, please check file format")
            return

        st.success(f"Loaded {len(uploaded)} files, total {len(combined):,} records")
        st.divider()

        # ── Filter Section ──────────────────────────────────────
        st.markdown("## 🔍 Filters")

        # Sheet Filter
        st.markdown("**Sheet**")
        all_sheets = sorted(combined['Sheet'].dropna().unique())
        sel_sheet = make_filter("Sheet", all_sheets, "f_sheet")
        if not sel_sheet:
            st.warning("Please select Sheet")
            return
        w = combined[combined['Sheet'].isin(sel_sheet)]
        st.divider()

        # MPA Filter
        st.markdown("**MPA**　*(Multi-select will sum data)*")
        all_mpas = sorted(w['MPA'].dropna().unique())
        sel_mpa = make_filter("MPA", all_mpas, "f_mpa")
        if not sel_mpa:
            st.warning("Please select MPA")
            return
        w = w[w['MPA'].isin(sel_mpa)]
        st.divider()

        # Type Filter
        all_types = sorted(w['Type'].dropna().unique())
        if all_types:
            st.markdown("**Type / Details**")
            sel_type = make_filter("Type", all_types, "f_type")
            if not sel_type:
                st.warning("Please select Type")
                return
            w = w[w['Type'].isin(sel_type)]
            st.divider()

        # Consign PN Filter
        all_parts = sorted(w['Consign_PN'].dropna().unique())
        st.markdown("**Consign PN / Part Number**")
        sel_part = make_filter("Part", all_parts, "f_part")
        if not sel_part:
            st.warning("Please select Part")
            return
        w = w[w['Consign_PN'].isin(sel_part)]
        st.divider()

        # Data Description Filter
        all_descs = sorted(w['Data_Description'].dropna().unique())
        st.markdown("**Data Description**")
        sel_desc = make_filter("Desc", all_descs, "f_desc")
        if not sel_desc:
            st.warning("Please select Data Description")
            return
        st.divider()

        # Month Filter
        st.markdown("**Month**")
        w['YearMonth'] = pd.to_datetime(w['Date']).dt.to_period('M')
        all_months = sorted(w['YearMonth'].dropna().unique().astype(str))
        sel_months = make_filter("Month", all_months, "f_month")
        if not sel_months:
            st.warning("Please select Month")
            return
        w = w[w['YearMonth'].astype(str).isin(sel_months)]

    # ────────── Main Area ───────────────────────────────────────
    st.title("📊 DDSS Forecast Analysis")
    st.caption("File naming: `MPA YYYYwk00.xlsx` (year+week format)")

    if 'combined_df' not in st.session_state or st.session_state['combined_df'] is None:
        st.info("👈 Please upload files in the left sidebar")
        return

    filtered = w[w['Data_Description'].isin(sel_desc)]
    if filtered.empty:
        st.warning("No matching data, please adjust filters")
        return

    # Overview Metrics - show week range using WeekLabel
    week_labels = sorted(combined['WeekLabel'].dropna().unique())
    wk_min_label = week_labels[0] if week_labels else "N/A"
    wk_max_label = week_labels[-1] if week_labels else "N/A"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Files", len(uploaded))
    c2.metric("Week Range", f"{wk_min_label} – {wk_max_label}")
    c3.metric("Selected MPA", f"{len(sel_mpa)} MPA(s)")
    c4.metric("Selected Parts", f"{len(sel_part)} Part(s)")

    with st.expander("📄 File Parse Status", expanded=False):
        st.dataframe(
            pd.DataFrame(status, columns=['Filename', 'Status', 'Records']),
            use_container_width=True,
            hide_index=True,
        )

    st.divider()

    mpa_label = " + ".join(sel_mpa)
    sheet_label = ", ".join(sel_sheet)
    part_label = (", ".join(sel_part) if len(sel_part) <= 3
                  else f"{len(sel_part)} Parts")

    download_tables = {}

    # Generate charts for each Data Description
    for desc in sel_desc:
        is_wos = is_wos_related(desc)

        st.subheader(f"📈  {desc}")
        st.caption(f"MPA: **{mpa_label}**　　Sheet: **{sheet_label}**　　Part: {part_label}")

        sub = filtered[filtered['Data_Description'] == desc]
        agg = sub.groupby(['WeekSortKey', 'WeekLabel', 'Date', 'Column_Type'], as_index=False)['Value'].sum()

        # Line Chart
        fct = agg[agg['Column_Type'] == 'Forecast']
        if not fct.empty:
            pp = (fct.pivot_table(index='Date', columns='WeekLabel',
                                  values='Value', aggfunc='sum')
                  .reset_index().sort_values('Date'))

            week_labels_chart = [c for c in pp.columns if c != 'Date']
            colors = px.colors.qualitative.Set2
            fig = go.Figure()

            # Determine hover template based on WOS status
            if is_wos:
                hover_template = '<b>%{fullData.name}</b><br>Date: %{x|%Y-%m-%d}<br>Value: %{y:.2f}<extra></extra>'
            else:
                hover_template = '<b>%{fullData.name}</b><br>Date: %{x|%Y-%m-%d}<br>Value: %{y:,.0f}<extra></extra>'

            for i, wk_label in enumerate(week_labels_chart):
                tmp = pp[['Date', wk_label]].dropna()
                fig.add_trace(go.Scatter(
                    x=tmp['Date'], y=tmp[wk_label],
                    mode='lines+markers',
                    name=wk_label,
                    line=dict(width=2.5, color=colors[i % len(colors)]),
                    marker=dict(size=6),
                    hovertemplate=hover_template,
                ))

            # Reduce date label density
            actual_dates = sorted(pp['Date'].unique())
            if len(actual_dates) > 10:
                display_dates = actual_dates[::3]
            elif len(actual_dates) > 5:
                display_dates = actual_dates[::2]
            else:
                display_dates = actual_dates

            fig.update_layout(
                margin=dict(t=30, b=60, l=50, r=20), height=420,
                xaxis_title=None, yaxis_title="Value",
                hovermode='x unified',
                legend=dict(orientation="h", y=1.05, x=1,
                            xanchor="right", yanchor="bottom"),
                xaxis=dict(
                    tickformat='%Y-%m-%d',
                    tickmode='array',
                    tickvals=display_dates,
                    tickangle=-45,
                    tickfont=dict(size=10),
                ),
                yaxis=dict(separatethousands=True),
            )
            st.plotly_chart(fig, use_container_width=True)

        # Split tables
        weekly_table, delta_table = build_pivot_tables(agg, metric_label=desc)

        # Table 1: Weekly data
        with st.expander(f"📋 {desc}", expanded=True):
            if weekly_table.empty:
                st.info("No data")
            else:
                h = min(420, (len(weekly_table) + 1) * 38 + 12)
                st.dataframe(
                    style_table(weekly_table, is_wos=is_wos),
                    use_container_width=True,
                    height=h,
                )
                download_tables[desc] = weekly_table

        # Table 2: Delta table
        with st.expander(f"📋 {desc} Delta", expanded=True):
            if delta_table.empty:
                st.info("No delta data")
            else:
                h = min(420, (len(delta_table) + 1) * 38 + 12)
                st.dataframe(
                    style_table(delta_table, is_wos=is_wos),
                    use_container_width=True,
                    height=h,
                )
                download_tables[f"{desc} Delta"] = delta_table

        st.divider()

    # Export
    st.subheader("💾 Export Data")
    col1, col2 = st.columns(2)

    with col1:
        if download_tables:
            parts = []
            for d, tbl in download_tables.items():
                t = tbl.copy()
                t.insert(0, 'Data_Description', d)
                parts.append(t)
            st.download_button(
                "⬇️ Download Detailed Tables (CSV)",
                data=pd.concat(parts).to_csv(index=False, encoding='utf-8-sig'),
                file_name=f"{'_'.join(sel_mpa)}_detailed_tables.csv",
                mime="text/csv",
                use_container_width=True,
            )

    with col2:
        st.download_button(
            "⬇️ Download Filtered Raw Data (CSV)",
            data=filtered.to_csv(index=False, encoding='utf-8-sig'),
            file_name="filtered_raw_data.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with st.expander("🔍 Raw Data Preview"):
        st.dataframe(filtered, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
