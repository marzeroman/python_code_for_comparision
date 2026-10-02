"""
Data Tool - click version
=========================
Check, clean, compare and pivot ANY tabular data from menus.
Same engine as compare_tool.py (keep both files in the same folder).

Run:
    pip install -r requirements.txt
    streamlit run data_tool.py
"""

import json
import os

import pandas as pd
import streamlit as st

from compare_tool import (AGG_FUNCS, DEFAULT_CLEANING, FILE_TYPES, MATCH_METHODS, MERGE_HOW,
                          SettingsError, check_quality, clean_table, column_profile, compare_schema,
                          compare_tables, compare_two_rows, demo_tables, find_row, fuzzy_compare,
                          logs_to_frames, make_pivot, read_file_bytes, rowwise_compare, to_excel_bytes,
                          validate_cleaning, value_overlap)

RULE_CHOICES = ["required", "unique", "min", "max", "allowed", "pattern", "not_future", "not_after"]


# =============================================================================
# HELPERS
# =============================================================================
@st.cache_data(show_spinner=False)
def load_file(name: str, data: bytes) -> dict:
    """{table_name: DataFrame}. Each Excel sheet becomes its own table."""
    base = os.path.splitext(name)[0]
    return {(base if s == "" else f"{base} [{s}]"): df
            for s, df in read_file_bytes(name, data).items() if not df.empty}


@st.cache_data(show_spinner="Cleaning data...")
def cached_clean(df: pd.DataFrame, cfg_json: str, label: str):
    return clean_table(df, json.loads(cfg_json), label)


@st.cache_data(show_spinner="Checking data quality...")
def cached_check(df: pd.DataFrame, cfg_json: str, label: str):
    return check_quality(df, label, json.loads(cfg_json))


def sample_tables() -> dict:
    a, b = demo_tables()
    return {"Sample: ERP suppliers": a, "Sample: Supplier statement": b}


def safe(df: pd.DataFrame) -> pd.DataFrame:
    """Make mixed-type columns display-safe in Streamlit."""
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(lambda x: "" if (not isinstance(x, (list, dict)) and pd.isna(x)) else str(x))
    return df


def show(df: pd.DataFrame, container=st):
    container.dataframe(safe(df), width="stretch", hide_index=True)


def download_buttons(sheets: dict, filename: str, csv_label="⬇ Download CSV (main result)"):
    c1, c2 = st.columns(2)
    c1.download_button("⬇ Download Excel (all results)", to_excel_bytes(sheets), f"{filename}.xlsx",
                       key=f"x_{filename}",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    first = next(iter(sheets.values()))
    c2.download_button(csv_label, first.to_csv(index=False).encode("utf-8-sig"),
                       f"{filename}.csv", mime="text/csv", key=f"c_{filename}")


def _rule_value(check, value):
    """Turn the text typed in the rules table into the right Python value."""
    v = "" if value is None else str(value).strip()
    if check in ("required", "unique", "not_future"):
        return True
    if check in ("min", "max"):
        try:
            return float(v)
        except ValueError:
            return v                                   # a date like 2020-01-01
    if check == "allowed":
        return [x.strip() for x in v.split(",") if x.strip()]
    return v


# =============================================================================
# SIDEBAR: CLEANING SETTINGS
# =============================================================================
def cleaning_sidebar(all_columns: list) -> dict | None:
    with st.sidebar.expander("🧹 3. Cleaning (runs before every task)", expanded=False):
        on = st.toggle("Apply cleaning", True, key="cl_on",
                       help="Fixes run on every table before checking, comparing or pivoting. "
                            "Every change is listed in the Cleaning log.")
        markers = st.text_input("Values that mean 'empty' (comma separated)",
                                ", ".join(m for m in DEFAULT_CLEANING["empty_markers"] if m), key="cl_mk")
        trim = st.checkbox("Trim spaces & invisible characters", True, key="cl_trim")
        unify = st.checkbox("Unify case/spacing variants ('ERBIL' → 'Erbil')", True, key="cl_unify")
        to_num = st.selectbox("Text → numbers", ["auto", "off"], key="cl_num")
        decimal = st.radio("Decimal separator", [".", ","], horizontal=True, key="cl_dec",
                           help="'.' for 1,234.56   |   ',' for 1.234,56")
        to_date = st.selectbox("Text → dates", ["auto", "off"], key="cl_date")
        day_first = st.checkbox("Day first (01/03/2026 = 1 March)", True, key="cl_dayfirst")
        dups = st.checkbox("Remove 100% identical duplicate rows", True, key="cl_dups")

        st.markdown("**Replace values** (fix known typos/variants)")
        rep = st.data_editor(pd.DataFrame({"Column": pd.Series(dtype=str), "Find": pd.Series(dtype=str),
                                           "Replace with": pd.Series(dtype=str)}),
                             num_rows="dynamic", key="cl_replace", hide_index=True,
                             column_config={"Column": st.column_config.SelectboxColumn(options=all_columns)})
        st.markdown("**Rules** (rows that break them are listed, not deleted)")
        rules_df = st.data_editor(pd.DataFrame({"Column": pd.Series(dtype=str), "Check": pd.Series(dtype=str),
                                                "Value": pd.Series(dtype=str)}),
                                  num_rows="dynamic", key="cl_rules", hide_index=True,
                                  column_config={"Column": st.column_config.SelectboxColumn(options=all_columns),
                                                 "Check": st.column_config.SelectboxColumn(options=RULE_CHOICES)})
        st.caption("Rule values: min/max → a number or date · allowed → Open, Closed · "
                   "pattern → e.g. ^\\d{11}$ · not_after → another column name · "
                   "required/unique/not_future → leave empty")
        st.markdown("**Merge rows with the same key** (optional)")
        mkey = st.multiselect("Key column(s)", all_columns, key="cl_mkey")
        mhow = {}
        if mkey:
            for col in st.multiselect("Columns to combine", [c for c in all_columns if c not in mkey], key="cl_mcols"):
                mhow[col] = st.selectbox(f"{col}", sorted(MERGE_HOW), index=sorted(MERGE_HOW).index("sum"),
                                         key=f"cl_mhow_{col}")
    if not on:
        return None

    replace = {}
    for _, r in rep.dropna(subset=["Column", "Find"]).iterrows():
        replace.setdefault(r["Column"], {})[str(r["Find"])] = "" if pd.isna(r["Replace with"]) else str(r["Replace with"])
    rules = []
    for _, r in rules_df.dropna(subset=["Column", "Check"]).iterrows():
        rules.append({"column": r["Column"], r["Check"]: _rule_value(r["Check"], r["Value"])})
    return {**DEFAULT_CLEANING,
            "empty_markers": [""] + [m.strip() for m in markers.split(",") if m.strip()],
            "trim": trim, "unify": unify, "to_number": "auto" if to_num == "auto" else [],
            "decimal": decimal, "to_date": "auto" if to_date == "auto" else [], "day_first": day_first,
            "replace": replace, "rules": rules, "remove_dups": dups,
            "merge": {"key": mkey, "how": mhow} if mkey else {}}


# =============================================================================
# PAGES
# =============================================================================
def page_check(raw: dict, cleaned: dict, cfg):
    st.subheader("🩺 Check data quality")
    st.caption("Looks at the values themselves, so it works on any kind of data. Nothing is changed here.")
    name = st.selectbox("Table", list(raw), key="ck_table")
    after = cfg is not None and st.toggle("Check the cleaned version instead", False, key="ck_after")
    df = cleaned[name] if after else raw[name]
    cfg_json = json.dumps(cfg or DEFAULT_CLEANING, default=str)
    r = cached_check(df, cfg_json, name)

    iss = r["issues"]
    m = st.columns(4)
    m[0].metric("Rows × columns", f"{len(df):,} × {len(df.columns)}")
    m[1].metric("Problems to fix", int((iss["Level"] == "Fix").sum()))
    m[2].metric("Things to check", int((iss["Level"] == "Check").sum()))
    m[3].metric("Possible typos", len(r["similar"]))
    if iss.empty:
        st.success("No problems found.")

    t1, t2, t3, t4 = st.tabs([f"Issues ({len(iss)})", "Columns", f"Similar spellings ({len(r['similar'])})",
                              f"Duplicate rows ({len(r['duplicates'])})"])
    t1.caption("Level: **Fix** = cleaning can fix it · **Check** = look at it yourself · **Info** = for your knowledge. "
               "Setting names in 'How to fix' match the sidebar cleaning options and compare_tool.py.")
    show(iss.drop(columns=["Table"]), t1)
    show(r["columns"].drop(columns=["Table"]), t2)
    t3.caption("If two values are really the same thing, add them under **Replace values** in the sidebar.")
    show(r["similar"].drop(columns=["Table"]), t3)
    show(r["duplicates"], t4)
    download_buttons({"Overview": r["overview"], "Issues": iss, "Columns": r["columns"],
                      "Similar spellings": r["similar"], "Duplicate rows": r["duplicates"]},
                     "data_quality_report", "⬇ Download CSV (issues)")


def page_clean(raw: dict, cleaned: dict, log_df, prob_df, cfg):
    st.subheader("🧹 Clean data")
    if cfg is None:
        st.info("Cleaning is off. Turn it on under **3. Cleaning** in the sidebar.")
        return
    st.caption("Change the options under **3. Cleaning** in the sidebar - the results update here, "
               "and every other page uses the cleaned data.")
    name = st.selectbox("Table", list(raw), key="cl_table")
    before, after = raw[name], cleaned[name]
    lg, pr = log_df[log_df["Table"] == name], prob_df[prob_df["Table"] == name]
    m = st.columns(4)
    m[0].metric("Rows", f"{len(after):,}", delta=len(after) - len(before) or None)
    m[1].metric("Cells changed", int(lg["Cells changed"].sum()) if len(lg) else 0)
    m[2].metric("Data problems", len(pr))
    m[3].metric("Columns", len(after.columns))
    t1, t2, t3, t4 = st.tabs([f"Cleaning log ({len(lg)})", f"Data problems ({len(pr)})",
                              "Cleaned data", "Original data"])
    show(lg.drop(columns=["Table"]), t1)
    t2.caption("Rows that break your rules, or values that could not be read. They are listed, not deleted.")
    show(pr.drop(columns=["Table"]), t2)
    t3.dataframe(safe(after.head(1000)), width="stretch")
    t4.dataframe(safe(before.head(1000)), width="stretch")

    c1, c2 = st.columns(2)
    c1.download_button("⬇ Cleaned data (Excel)", to_excel_bytes({"Data": after}), f"{name}_cleaned.xlsx",
                       key="cl_dl_x", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    c2.download_button("⬇ Cleaned data (CSV)", after.to_csv(index=False).encode("utf-8-sig"),
                       f"{name}_cleaned.csv", mime="text/csv", key="cl_dl_c")
    st.download_button("⬇ Cleaning report (log + problems)",
                       to_excel_bytes({"Cleaning log": lg, "Data problems": pr}),
                       f"{name}_cleaning_report.xlsx", key="cl_dl_r",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_pivot(tables: dict):
    st.subheader("📊 Pivot table")
    name = st.selectbox("Table", list(tables), key="pv_table")
    df = tables[name]
    cols = list(df.columns)
    c1, c2, c3 = st.columns(3)
    rows = c1.multiselect("Rows", cols, key="pv_rows")
    columns = c2.multiselect("Columns", [c for c in cols if c not in rows], key="pv_cols")
    values = c3.multiselect("Values (empty = count rows)", [c for c in cols if c not in rows + columns], key="pv_vals")
    c4, c5, c6 = st.columns(3)
    agg_label = c4.selectbox("Calculation", list(AGG_FUNCS), key="pv_agg")
    totals = c5.checkbox("Show totals", True, key="pv_tot")
    fill0 = c6.checkbox("Fill empty cells with 0", True, key="pv_fill")
    with st.expander("Filter data first (optional)"):
        fcol = st.selectbox("Filter column", ["(none)"] + cols, key="pv_fcol")
        if fcol != "(none)":
            opts = sorted(df[fcol].dropna().astype(str).unique())
            keep = st.multiselect("Keep these values", opts, default=opts, key="pv_fvals")
            df = df[df[fcol].astype(str).isin(keep)]
    if not rows and not columns:
        st.info("Pick at least one field for **Rows** or **Columns**.")
        return
    try:
        result = make_pivot(df, rows, columns, values, AGG_FUNCS[agg_label], totals, fill0)
    except Exception as e:
        st.error(f"Could not build this pivot: {e}")
        return
    show(result)
    if len(rows) == 1 and st.checkbox("Show as bar chart", key="pv_chart"):
        chart = result[result[rows[0]].astype(str) != "Total"].set_index(rows[0])
        chart = chart.drop(columns=["Total"], errors="ignore").select_dtypes("number")
        if not chart.empty:
            st.bar_chart(chart)
    download_buttons({"Pivot": result}, "pivot_result")


def page_compare_tables(tables: dict):
    st.subheader("🔀 Compare two tables")
    names = list(tables)
    c1, c2 = st.columns(2)
    na = c1.selectbox("Table A", names, key="ct_a")
    nb = c2.selectbox("Table B", names, index=min(1, len(names) - 1), key="ct_b")
    a, b = tables[na], tables[nb]
    if na == nb:
        st.warning("A and B are the same table - pick two different ones.")
    with st.expander("Column structure (which columns exist where)"):
        show(compare_schema(a, b))
    mode = st.radio("How should rows be matched?", ["Exact key (IDs, invoice numbers, codes)",
                                                    "Fuzzy (names spelled differently)"],
                    horizontal=True, key="ct_mode")
    fuzzy_section(a, b) if mode.startswith("Fuzzy") else exact_section(a, b)


def exact_section(a, b):
    common = [c for c in a.columns if c in b.columns]
    if not common:
        st.error("These tables have no column names in common. Use **Fuzzy** matching, "
                 "which lets you pick differently-named key columns.")
        return
    keys = st.multiselect("Key column(s) - what identifies a row?", common, key="ct_keys")
    if not keys:
        st.info("Choose at least one key column to match rows between the tables.")
        return
    cols = st.multiselect("Columns to compare", [c for c in common if c not in keys],
                          default=[c for c in common if c not in keys], key="ct_cols")
    c3, c4 = st.columns(2)
    tol = c3.number_input("Number tolerance (0 = exact)", min_value=0.0, value=0.0, step=0.01, key="ct_tol")
    ignore_case = c4.checkbox("Ignore upper/lower case", True, key="ct_case")
    r = compare_tables(a, b, keys, cols, tol, ignore_case)
    if r["dup_a"] or r["dup_b"]:
        st.warning(f"Keys are not unique (duplicates: A={r['dup_a']}, B={r['dup_b']}). Add another key column, "
                   "or use **Merge rows with the same key** in the sidebar.")
    m = st.columns(5)
    m[0].metric("Rows in A", len(a))
    m[1].metric("Rows in B", len(b))
    m[2].metric("Only in A", len(r["only_a"]))
    m[3].metric("Only in B", len(r["only_b"]))
    m[4].metric("Matched but changed", f"{r['matched_changed']} / {r['matched']}")
    t1, t2, t3 = st.tabs([f"Changed values ({len(r['diffs'])})", f"Only in A ({len(r['only_a'])})",
                          f"Only in B ({len(r['only_b'])})"])
    show(r["diffs"], t1)
    show(r["only_a"], t2)
    show(r["only_b"], t3)
    download_buttons({"Changed values": r["diffs"], "Only in A": r["only_a"],
                      "Only in B": r["only_b"], "Columns": compare_schema(a, b)}, "table_comparison")


def fuzzy_section(a, b):
    c1, c2 = st.columns(2)
    key_a = c1.multiselect("Key column(s) in A (e.g. a name)", list(a.columns), key="fz_ka")
    key_b = c2.multiselect("Matching key column(s) in B - same order", list(b.columns), key="fz_kb")
    if not key_a or len(key_a) != len(key_b):
        st.info("Pick the key column(s) in both tables - the same number on each side.")
        return
    st.markdown("**Columns to check after matching** (A column → B column)")
    others_b = ["(skip)"] + [c for c in b.columns if c not in key_b]
    pairs, grid = {}, st.columns(3)
    for n, ca in enumerate([c for c in a.columns if c not in key_a]):
        cb = grid[n % 3].selectbox(ca, others_b, index=others_b.index(ca) if ca in others_b else 0,
                                   key=f"fz_pair_{ca}")
        if cb != "(skip)":
            pairs[ca] = cb
    with st.expander("Matching settings", expanded=True):
        s1, s2, s3 = st.columns(3)
        threshold = s1.slider("Minimum similarity %", 50, 100, 85, key="fz_thr",
                              help="Pairs below this are not matched. 85-90 is a good start.")
        review = s2.slider("Flag for review below %", 50, 100, 95, key="fz_rev")
        method = s3.selectbox("Method", list(MATCH_METHODS), key="fz_method",
                              format_func=lambda k: f"{k} - {MATCH_METHODS[k][1]}")
        s4, s5, s6 = st.columns(3)
        words = s4.text_input("Words to ignore (comma separated)", "co, company, ltd, llc, inc, the", key="fz_words")
        tol = s5.number_input("Number tolerance", min_value=0.0, value=0.0, step=0.01, key="fz_tol")
        ignore_case = s6.checkbox("Ignore upper/lower case", True, key="fz_case")
        mm_a = st.selectbox("Optional: only pair rows that also share this exact value (column in A)",
                            ["(none)"] + list(a.columns), key="fz_mma")
        must = {}
        if mm_a != "(none)":
            mm_b = st.selectbox("...and this column in B", list(b.columns), key="fz_mmb",
                                index=list(b.columns).index(mm_a) if mm_a in b.columns else 0)
            must = {mm_a: mm_b}
    try:
        r = fuzzy_compare(a, b, key_a, key_b, pairs, threshold, method, must,
                          [w.strip() for w in words.split(",") if w.strip()], review, tol, ignore_case)
    except Exception as e:
        st.error(f"Could not run fuzzy matching: {e}")
        return
    s = r["stats"]
    m = st.columns(5)
    m[0].metric("Matched pairs", s["Matched pairs"])
    m[1].metric("Need review", s["Matches needing review"])
    m[2].metric("Matched, values differ", s["Matched with different values"])
    m[3].metric("Only in A", s["Only in A (no match)"])
    m[4].metric("Only in B", s["Only in B (no match)"])
    t1, t2, t3, t4 = st.tabs([f"Matches ({len(r['matches'])})", f"Different values ({len(r['diffs'])})",
                              f"Only in A ({len(r['only_a'])})", f"Only in B ({len(r['only_b'])})"])
    show(r["matches"], t1)
    show(r["diffs"], t2)
    t3.caption("'Closest %' shows the best candidate that was below your minimum - "
               "if it's clearly the same, lower the minimum a little.")
    show(r["only_a"], t3)
    show(r["only_b"], t4)
    download_buttons({"Summary": pd.DataFrame(list(s.items()), columns=["Item", "Value"]),
                      "Matches": r["matches"], "Different values": r["diffs"],
                      "Only in A": r["only_a"], "Only in B": r["only_b"]}, "fuzzy_comparison")


def page_compare_columns(tables: dict):
    st.subheader("↔️ Compare two columns")
    names = list(tables)
    c1, c2 = st.columns(2)
    ta = c1.selectbox("Table A", names, key="cc_ta")
    ca = c1.selectbox("Column A", list(tables[ta].columns), key="cc_ca")
    tb = c2.selectbox("Table B", names, key="cc_tb")
    cols_b = list(tables[tb].columns)
    cb = c2.selectbox("Column B", cols_b, index=min(1, len(cols_b) - 1) if ta == tb else 0, key="cc_cb")
    sa, sb = tables[ta][ca], tables[tb][cb]
    la, lb = f"A: {ca}", f"B: {cb}"
    ignore_case = st.checkbox("Ignore upper/lower case", True, key="cc_case")
    st.markdown("**Summary side by side**")
    prof = pd.DataFrame({la: pd.Series(column_profile(sa)), lb: pd.Series(column_profile(sb))})
    prof = prof.reset_index().rename(columns={"index": "Metric"})
    show(prof)
    st.markdown("**Distinct values: overlap**")
    both, only_a, only_b = value_overlap(sa, sb, ignore_case)
    t1, t2, t3 = st.tabs([f"In both ({len(both)})", f"Only in A ({len(only_a)})", f"Only in B ({len(only_b)})"])
    show(pd.DataFrame({"Value": both}), t1)
    show(pd.DataFrame({"Value": only_a}), t2)
    show(pd.DataFrame({"Value": only_b}), t3)
    sheets = {"Summary": prof, "In both": pd.DataFrame({"Value": both}),
              "Only in A": pd.DataFrame({"Value": only_a}), "Only in B": pd.DataFrame({"Value": only_b})}
    if ta == tb or st.checkbox("Also compare row by row by position (only if both tables are in the same order)",
                               key="cc_pos"):
        tol = st.number_input("Number tolerance (0 = exact)", min_value=0.0, value=0.0, step=0.01, key="cc_tol")
        rw = rowwise_compare(sa, sb, la, lb, tol, ignore_case)
        view = rw[~rw["Same?"]] if st.checkbox("Show only rows that differ", True, key="cc_only") else rw
        st.markdown(f"**Row by row** - {int((~rw['Same?']).sum())} of {len(rw)} rows differ")
        show(view)
        sheets = {"Row by row": rw, **sheets}
    download_buttons(sheets, "column_comparison")


def page_compare_rows(tables: dict):
    st.subheader("↕️ Compare two rows")
    names = list(tables)
    picked = []
    for side, col in zip(["A", "B"], st.columns(2)):
        t = col.selectbox(f"Table {side}", names, index=0 if side == "A" else min(1, len(names) - 1),
                          key=f"cr_t{side}")
        df = tables[t]
        by = col.selectbox(f"Find row {side} by", ["Row number"] + list(df.columns), key=f"cr_by{side}")
        default = ("0" if side == "A" else "1") if by == "Row number" else ""
        val = col.text_input(f"Value for row {side}", default, key=f"cr_v{side}")
        row, msg = find_row(df, by, val) if val else (None, "Enter a value.")
        if msg:
            col.caption(msg)
        picked.append((row, df))
    (ra, dfa), (rb, dfb) = picked
    if ra is None or rb is None:
        return
    with st.expander("Pair differently-named columns (optional)"):
        pairs = {}
        for ca in [c for c in dfa.columns if c not in dfb.columns]:
            cb = st.selectbox(f"{ca} (A) ↔", ["(skip)"] + [c for c in dfb.columns if c not in dfa.columns],
                              key=f"cr_pair_{ca}")
            if cb != "(skip)":
                pairs[ca] = cb
    ignore_case = st.checkbox("Ignore upper/lower case", True, key="cr_case")
    out = compare_two_rows(ra, rb, ignore_case, 0.0, pairs)
    st.markdown(f"**{int((~out['Same?']).sum())} of {len(out)} compared columns differ**")
    show(out)
    download_buttons({"Row comparison": out}, "row_comparison")


# =============================================================================
# MAIN
# =============================================================================
def main():
    st.set_page_config(page_title="Data Tool", page_icon="📊", layout="wide")
    st.title("📊 Data Tool - check, clean, compare & pivot")

    with st.sidebar:
        st.header("1. Load data")
        files = st.file_uploader("CSV, Excel, JSON or Parquet (several allowed)",
                                 type=FILE_TYPES, accept_multiple_files=True)
        st.caption("Power BI: on a visual click **⋯ → Export data**, then upload the file here.")
        use_sample = st.checkbox("Use sample data to try it out", value=not files)
        raw = {}
        for f in files or []:
            try:
                raw.update(load_file(f.name, f.getvalue()))
            except Exception as e:
                st.error(f"{f.name}: {e}")
        if use_sample:
            raw.update(sample_tables())

        st.header("2. Choose a task")
        mode = st.radio("Task", ["Check data", "Clean data", "Pivot table", "Compare two tables",
                                 "Compare columns", "Compare rows"], label_visibility="collapsed")

    all_columns = list(dict.fromkeys(c for df in raw.values() for c in df.columns))
    cfg = cleaning_sidebar(all_columns)

    if not raw:
        st.info("⬅ Upload one or more files in the sidebar (or tick 'Use sample data').")
        return

    # Clean every table once (cached) - all pages below use the cleaned versions
    cleaned, logs, probs = dict(raw), [], []
    if cfg is not None:
        try:
            validate_cleaning(cfg, raw)
            cfg_json = json.dumps(cfg, default=str)
            for name, df in raw.items():
                cdf, lg, pr = cached_clean(df, cfg_json, name)
                cleaned[name] = cdf
                logs += lg
                probs += pr
        except SettingsError as e:
            st.error(f"Cleaning settings problem:\n\n{e}")
            cfg = None
    log_df, prob_df = logs_to_frames(logs, probs)

    if mode not in ("Check data", "Clean data"):
        if cfg is not None:
            st.caption(f"🧹 Using cleaned data: {int(log_df['Cells changed'].sum()) if len(log_df) else 0} cells changed, "
                       f"{len(prob_df)} data problems (see **Clean data**).")
        with st.expander(f"Preview tables ({len(cleaned)})"):
            pv = st.selectbox("Table", list(cleaned), key="preview")
            st.caption(f"{cleaned[pv].shape[0]} rows × {cleaned[pv].shape[1]} columns")
            st.dataframe(safe(cleaned[pv].head(100)), width="stretch")

    if mode == "Check data":
        page_check(raw, cleaned, cfg)
    elif mode == "Clean data":
        page_clean(raw, cleaned, log_df, prob_df, cfg)
    else:
        {"Pivot table": page_pivot, "Compare two tables": page_compare_tables,
         "Compare columns": page_compare_columns, "Compare rows": page_compare_rows}[mode](cleaned)


if __name__ == "__main__":
    main()