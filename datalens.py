"""DataLens — generic Excel/CSV analyzer, compare & merge tool for Streamlit.

Upload any spreadsheet and it auto-detects columns, profiles them, and builds
KPI tiles, a column schema, category charts, a group/aggregate explorer, and a
low-data-records finder. A second file can be uploaded to compare and merge
against the first, with duplicate-key detection and resolution.
"""

import re

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

KPI_KEYWORDS = re.compile(
    r"total|amount|sum|population|revenue|due|paid|balance|variance|count|number|qty|quantity|value|score",
    re.I,
)
CAT_KEYWORDS = re.compile(
    r"region|district|status|type|network|gender|sex|service|category|department|zone|level|grade|class|batch",
    re.I,
)
ID_LIKE = re.compile(r"^(no\.?|s/?n\.?|sl\.?\s?no\.?|id|#|index|row)$", re.I)
IDENTIFIER_NUMERIC = re.compile(
    r"mobile|phone|contact|tel(ephone)?|cell|whatsapp|fax|zip\s*code|postal\s*code|\bpin\b|\bcode\b",
    re.I,
)
HEADER_STOPWORDS = {
    "total", "number", "no", "of", "the", "a", "an", "count", "onboarded",
    "and", "amount", "sum", "qty", "quantity", "value",
}


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def slugify(s):
    s = re.sub(r"[^A-Za-z0-9]+", "_", str(s)).strip("_").lower()
    return s or "data"


def trim_zero(x):
    s = f"{x:.1f}"
    return s[:-2] if s.endswith(".0") else s


def format_compact(n):
    if n is None or (isinstance(n, float) and np.isnan(n)):
        return "—"
    abs_n = abs(n)
    if abs_n >= 1e9:
        return f"{trim_zero(n / 1e9)}B"
    if abs_n >= 1e6:
        return f"{trim_zero(n / 1e6)}M"
    if abs_n >= 1e3:
        return f"{trim_zero(n / 1e3)}K"
    if float(n).is_integer():
        return f"{int(n):,}"
    return f"{n:.2f}"


def is_blank(v):
    if v is None:
        return True
    if isinstance(v, float) and np.isnan(v):
        return True
    if pd.isna(v):
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def filled_mask(df):
    """A same-shape boolean mask: True where a cell counts as non-empty
    (NaN and blank/whitespace-only strings both count as empty)."""
    mask = df.notna()
    obj_cols = df.select_dtypes(include="object").columns
    for c in obj_cols:
        mask[c] = mask[c] & (df[c].astype(str).str.strip() != "")
    return mask


def normalize_key(v, pad_width=0):
    if is_blank(v):
        return ""
    if isinstance(v, pd.Timestamp):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, (int, float, np.integer, np.floating)):
        padded = pad_recovered(v, pad_width)
        if padded is not None:
            return padded
        f = float(v)
        return str(int(f)) if f.is_integer() else str(f)
    s = str(v).strip().upper()
    return re.sub(r"\s+", " ", s)


def header_tokens(h):
    words = re.split(r"[^a-z0-9]+", str(h).lower())
    return [w for w in words if w and w not in HEADER_STOPWORDS]


def header_token_score(h1, h2):
    t1 = header_tokens(h1)
    t2 = set(header_tokens(h2))
    return sum(1 for w in t1 if w in t2)


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------

def _dedupe_headers(cols):
    seen = {}
    out = []
    for i, c in enumerate(cols):
        name = "" if c is None else str(c).strip()
        if name == "" or name.lower().startswith("unnamed"):
            name = f"Column {i + 1}"
        if name not in seen:
            seen[name] = 0
            out.append(name)
        else:
            seen[name] += 1
            out.append(f"{name} ({seen[name]})")
    return out


def _coerce_csv_types(df):
    """pd.read_csv(dtype=str) reads every cell as text -- necessary because
    unlike Excel, CSV carries no per-cell type metadata, so pandas' normal
    auto-detection will happily convert a whole column like "0030407" to the
    number 30407, permanently losing the leading zeros. This converts each
    column back to numeric/boolean/datetime, but ONLY where doing so is safe
    for every value in that column -- a column with even one leading-zero
    numeric-looking value (an ID, a phone number) is left as text."""
    out = df.copy()
    for col in out.columns:
        s = out[col]
        non_null = s[s.notna()]
        non_null = non_null[non_null.str.strip() != ""]
        if non_null.empty:
            continue
        stripped = non_null.str.strip()

        if stripped.str.match(r"^0\d+$").any():
            continue  # a pure digit string starting with 0, e.g. "0030407",
            # "0244123456" -- keep as text. Anchored so a date like
            # "01/15/2024" (has separators) doesn't get caught here too.

        numeric = pd.to_numeric(stripped, errors="coerce")
        if numeric.notna().all():
            out[col] = pd.to_numeric(s, errors="coerce")
            continue

        lowered = stripped.str.lower()
        if lowered.isin(["true", "false"]).all():
            out[col] = s.str.strip().str.lower().map({"true": True, "false": False})
            continue

        try:
            parsed = pd.to_datetime(stripped, errors="raise")
            if parsed.notna().all():
                out[col] = pd.to_datetime(s, errors="coerce")
        except (ValueError, TypeError):
            pass
    return out


def load_dataframe(uploaded_file, key_prefix):
    """Read an uploaded file into a DataFrame, with a sheet picker for Excel
    files with more than one sheet. Returns None (and shows an error) on
    failure."""
    name = uploaded_file.name
    try:
        if name.lower().endswith((".xlsx", ".xls")):
            xls = pd.ExcelFile(uploaded_file)
            sheet_name = xls.sheet_names[0]
            if len(xls.sheet_names) > 1:
                sheet_name = st.selectbox(
                    "Sheet", xls.sheet_names, key=f"{key_prefix}_sheet"
                )
            df = xls.parse(sheet_name)
        elif name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file, dtype=str)
            df = _coerce_csv_types(df)
        else:
            st.error("Unsupported file type. Please upload .xlsx, .xls, or .csv.")
            return None
    except Exception as e:
        st.error(f"Couldn't parse '{name}': {e}")
        return None

    df = df.dropna(how="all")
    if df.empty:
        st.error(f"'{name}' has headers but no data rows.")
        return None
    df.columns = _dedupe_headers(df.columns)
    return df.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Column analysis
# ---------------------------------------------------------------------------

def classify_columns(df):
    meta = {}
    n = len(df)
    fmask = filled_mask(df)
    for col in df.columns:
        s = df[col]
        mask = fmask[col]
        filled = int(mask.sum())
        values = s[mask]

        # If a column mixes text cells like "0030407" with true-number cells
        # for the same field (a common Excel data-entry inconsistency), the
        # numeric cells lose their leading zeros. pad_width records the widest
        # all-digit *text* value seen in this column, so display/export/match
        # logic can zero-pad any stray numeric cell back to that width.
        pad_width = 0
        for v in values:
            if isinstance(v, str):
                t_stripped = v.strip()
                if t_stripped.isdigit() and len(t_stripped) > pad_width:
                    pad_width = len(t_stripped)

        if filled == 0:
            meta[col] = {"type": "empty", "filled": 0, "total": n, "pad_width": 0}
            continue

        if pd.api.types.is_bool_dtype(s):
            t = "boolean"
        elif pd.api.types.is_datetime64_any_dtype(s):
            t = "date"
        elif pd.api.types.is_numeric_dtype(s):
            t = "number"
        else:
            t = "text"

        info = {"type": t, "filled": filled, "total": n, "pad_width": pad_width}

        if t == "number":
            nums = pd.to_numeric(values, errors="coerce").dropna()
            info["no_grouping"] = bool(ID_LIKE.match(col.strip()) or IDENTIFIER_NUMERIC.search(col))
            if len(nums):
                info["sum"] = float(nums.sum())
                info["avg"] = float(nums.mean())
                info["min"] = float(nums.min())
                info["max"] = float(nums.max())
            if info["no_grouping"]:
                info["unique"] = int(nums.astype(str).nunique())
        elif t == "date":
            info["min_date"] = values.min()
            info["max_date"] = values.max()
        elif t == "boolean":
            info["true_share"] = float(values.mean())
        elif t == "text":
            vals = values.astype(str).str.strip().str.upper()
            unique_count = int(vals.nunique())
            ratio = unique_count / filled if filled else 0
            info["unique"] = unique_count
            info["classification"] = "category" if (unique_count <= 30 and ratio <= 0.6) else "text"
            counts = vals.value_counts()
            info["top_value"] = counts.index[0]
            info["top_share"] = float(counts.iloc[0] / filled)
            info["counts"] = counts

        meta[col] = info
    return meta


def pick_kpi_columns(meta, limit=4):
    order = list(meta.keys())
    numeric_cols = [c for c, m in meta.items() if m["type"] == "number" and not m.get("no_grouping")]
    numeric_cols.sort(key=lambda c: (0 if KPI_KEYWORDS.search(c) else 1, order.index(c)))
    return numeric_cols[:limit]


def pad_recovered(v, pad_width):
    """If v is a non-negative whole number shorter than pad_width, return it
    zero-padded back to that width (recovering a leading zero Excel dropped
    when it stored a sibling cell as a number instead of text). Returns None
    when no padding applies."""
    if not pad_width or isinstance(v, bool) or not isinstance(v, (int, float, np.integer, np.floating)):
        return None
    f = float(v)
    if not f.is_integer() or f < 0:
        return None
    digits = str(int(f))
    return digits.zfill(pad_width) if len(digits) < pad_width else None


def to_display_df(df, meta):
    """Uppercase text columns for display; leave numeric/date/boolean columns
    in their native dtype so Streamlit renders them natively. Any numeric
    stray in a column that also holds zero-padded text values (e.g. a mixed
    ID/phone column) is recovered back to the observed width."""
    disp = df.copy()
    for col, info in meta.items():
        pad_width = info.get("pad_width", 0)
        if info.get("type") == "text":
            def _fmt_text(v, pad_width=pad_width):
                if is_blank(v):
                    return ""
                padded = pad_recovered(v, pad_width)
                if padded is not None:
                    return padded
                return str(v).strip().upper()
            disp[col] = disp[col].apply(_fmt_text)
        elif pad_width:
            def _fmt_num(v, pad_width=pad_width):
                if is_blank(v):
                    return v
                padded = pad_recovered(v, pad_width)
                return padded if padded is not None else v
            disp[col] = disp[col].apply(_fmt_num)
    disp.columns = [str(c).upper() for c in disp.columns]
    return disp


def search_mask(df, query):
    if not query:
        return pd.Series(True, index=df.index)
    q = query.strip()
    if not q:
        return pd.Series(True, index=df.index)
    return df.astype(str).apply(lambda row: row.str.contains(q, case=False, na=False, regex=False).any(), axis=1)


# ---------------------------------------------------------------------------
# Render: stats, schema, category charts, explorer, low-data
# ---------------------------------------------------------------------------

def render_stats(df, meta):
    total_records = len(df)
    total_cols = len(df.columns)
    filled_cells = sum(m.get("filled", 0) for m in meta.values())
    total_cells = total_records * total_cols
    completeness = round(filled_cells / total_cells * 100) if total_cells else 0

    tiles = [("Total records", f"{total_records:,}", f"{total_cols} columns detected")]
    kpis = pick_kpi_columns(meta)
    for c in kpis:
        m = meta[c]
        tiles.append((
            c.upper(),
            format_compact(m["sum"]),
            f"avg {format_compact(m['avg'])} · range {format_compact(m['min'])}–{format_compact(m['max'])}",
        ))
    if not kpis:
        cat_count = sum(1 for m in meta.values() if m.get("classification") == "category")
        tiles.append(("Category columns", str(cat_count), "auto-detected"))
    tiles.append(("Data completeness", f"{completeness}%", f"{filled_cells:,} of {total_cells:,} cells filled"))

    cols = st.columns(len(tiles))
    for widget, (label, value, sub) in zip(cols, tiles):
        with widget:
            st.metric(label, value)
            st.caption(sub)


def render_schema(meta):
    rows = []
    for col, m in meta.items():
        t = m["type"]
        filled_pct = round(m["filled"] / m["total"] * 100) if m.get("total") else 0
        if t == "number" and m.get("no_grouping"):
            type_label, summary = "NUMBER", f"Identifier / contact number · {m.get('unique', 0):,} unique"
        elif t == "number":
            type_label, summary = "NUMBER", f"Σ {format_compact(m['sum'])} · avg {format_compact(m['avg'])}"
        elif t == "date":
            type_label = "DATE"
            summary = f"{m['min_date']:%d/%m/%Y} – {m['max_date']:%d/%m/%Y}"
        elif t == "boolean":
            type_label, summary = "BOOLEAN", f"{round(m['true_share'] * 100)}% true"
        elif t == "text" and m.get("classification") == "category":
            type_label, summary = "CATEGORY", f"{m['top_value']} ({round(m['top_share'] * 100)}%) top"
        elif t == "text":
            type_label, summary = "TEXT", f"{m['unique']:,} unique values"
        else:
            type_label, summary = "EMPTY", "—"
        rows.append({"Column": col.upper(), "Type": type_label, "Filled": f"{filled_pct}%", "Summary": summary})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def render_category_charts(meta):
    cats = [(c, m) for c, m in meta.items() if m.get("classification") == "category"]
    cats.sort(key=lambda item: (0 if CAT_KEYWORDS.search(item[0]) else 1, item[1]["unique"]))
    cats = cats[:4]

    if not cats:
        st.info("No clear category columns were detected in this sheet.")
        return

    cols = st.columns(2)
    for i, (c, m) in enumerate(cats):
        with cols[i % 2]:
            counts = m["counts"]
            top = counts.head(7)
            if len(counts) > 7:
                top = pd.concat([top, pd.Series({"OTHER": counts.iloc[7:].sum()})])
            chart_df = top.reset_index()
            chart_df.columns = ["Value", "Count"]
            fig = px.bar(chart_df, x="Count", y="Value", orientation="h", title=f"{c.upper()} ({m['filled']:,})")
            fig.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False, margin=dict(t=40, l=10, r=10, b=10))
            st.plotly_chart(fig, use_container_width=True)
            csv = chart_df.to_csv(index=False).encode("utf-8")
            st.download_button(f"Download {c.upper()} breakdown", csv, file_name=f"{slugify(c)}_breakdown.csv", key=f"catdl_{i}_{slugify(c)}")


def render_explorer(df, meta, key_prefix):
    st.subheader("Group & aggregate")
    all_cols = list(meta.keys())
    numeric_cols = [c for c, m in meta.items() if m["type"] == "number" and not m.get("no_grouping")]
    cat_cols = [c for c, m in meta.items() if m.get("classification") == "category"]
    default_group = cat_cols[0] if cat_cols else all_cols[0]

    c1, c2, c3 = st.columns(3)
    with c1:
        group_col = st.selectbox("Group by", all_cols, index=all_cols.index(default_group), key=f"{key_prefix}_group")
    with c2:
        measure_options = ["Number of records"] + numeric_cols
        measure = st.selectbox("Measure", measure_options, key=f"{key_prefix}_measure")
    with c3:
        agg = st.selectbox(
            "Aggregation", ["Sum", "Average", "Minimum", "Maximum"],
            key=f"{key_prefix}_agg", disabled=(measure == "Number of records"),
        )

    group_pad_width = meta[group_col].get("pad_width", 0)

    def _group_key(v):
        if is_blank(v):
            return "(BLANK)"
        padded = pad_recovered(v, group_pad_width)
        if padded is not None:
            return padded
        return str(v).strip().upper()

    work = df.copy()
    work["_group"] = work[group_col].apply(_group_key)

    if measure == "Number of records":
        result = work.groupby("_group").size().sort_values(ascending=False)
    else:
        agg_func = {"Sum": "sum", "Average": "mean", "Minimum": "min", "Maximum": "max"}[agg]
        result = work.groupby("_group")[measure].agg(agg_func).sort_values(ascending=False)

    cap = 12
    top = result.head(cap)
    if len(result) > cap:
        rest_keys = result.index[cap:]
        if measure == "Number of records" or agg == "Sum":
            other_val = result.iloc[cap:].sum()
        elif agg == "Average":
            other_val = work.loc[work["_group"].isin(rest_keys), measure].mean()
        elif agg == "Minimum":
            other_val = result.iloc[cap:].min()
        else:
            other_val = result.iloc[cap:].max()
        top = pd.concat([top, pd.Series({"OTHER": other_val})])

    chart_df = top.reset_index()
    chart_df.columns = ["Group", "Value"]
    fig = px.bar(chart_df, x="Value", y="Group", orientation="h")
    fig.update_layout(yaxis={"categoryorder": "total ascending"}, showlegend=False, margin=dict(t=20, l=10, r=10, b=10))
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(chart_df, use_container_width=True, hide_index=True)

    csv = chart_df.to_csv(index=False).encode("utf-8")
    st.download_button("Download this breakdown as CSV", csv, file_name="group_by.csv", mime="text/csv", key=f"{key_prefix}_dl_explorer")


def render_low_data(df, meta, key_prefix):
    st.subheader("Low-data records")
    pct = st.slider("Flag records with less than this % of fields filled", 1, 99, 50, key=f"{key_prefix}_lowdata_pct")
    threshold = pct / 100

    fmask = filled_mask(df)
    completeness = fmask.sum(axis=1) / len(df.columns)
    flagged_completeness = completeness[completeness < threshold].sort_values()
    flagged = df.loc[flagged_completeness.index].copy()
    flagged_display = to_display_df(flagged, meta)
    flagged_display["FILLED %"] = (flagged_completeness * 100).round().astype(int).astype(str) + "%"

    st.write(f"{len(flagged_display):,} of {len(df):,} records flagged")
    st.dataframe(flagged_display, use_container_width=True, hide_index=True)

    csv = flagged_display.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download low-data records", csv, file_name="low_data_records.csv",
        mime="text/csv", key=f"{key_prefix}_dl_lowdata", disabled=flagged_display.empty,
    )


# ---------------------------------------------------------------------------
# Compare & merge
# ---------------------------------------------------------------------------

def looks_like_id_column(name):
    nm = name.strip()
    if re.search(r"ID$", nm):  # camelCase/ALLCAPS suffix, e.g. SubscriberID (case-sensitive)
        return True
    if re.search(r"(^|[^a-zA-Z])id$", nm, re.I):  # standalone "id" word, e.g. "School Id", "ID"
        return True
    return False


def guess_key_column(meta):
    for c in meta:
        if looks_like_id_column(c):
            return c
    for c in meta:
        if re.search(r"name", c, re.I):
            return c
    for c, m in meta.items():
        if m.get("classification") == "category" or m.get("type") == "text":
            return c
    return next(iter(meta))


def resolve_duplicates(df, key_col, key_pad_width=0):
    """Group rows by normalized key. Groups with more than one row are
    collapsed into a single row: the most-complete row is kept as the base,
    and any of its blank fields are filled in from the other rows in the
    group. Returns (resolved_df indexed by key, list of dup-group summaries,
    dataframe of the raw duplicate rows)."""
    work = df.copy()
    work["_key"] = work[key_col].apply(lambda v: normalize_key(v, key_pad_width))
    work = work[work["_key"] != ""]
    work["_filled_count"] = filled_mask(df.loc[work.index]).sum(axis=1)

    resolved_rows = []
    dup_groups = []
    dup_indices = []

    for key, group in work.groupby("_key", sort=False):
        if len(group) == 1:
            resolved_rows.append(group.iloc[0])
            continue
        dup_groups.append({"key": key, "count": len(group)})
        dup_indices.extend(group.index.tolist())

        primary = group.loc[group["_filled_count"].idxmax()].copy()
        for col in df.columns:
            if is_blank(primary[col]):
                for _, row in group.iterrows():
                    if not is_blank(row[col]):
                        primary[col] = row[col]
                        break
        resolved_rows.append(primary)

    resolved_df = pd.DataFrame(resolved_rows).drop(columns=["_filled_count"], errors="ignore")
    resolved_df = resolved_df.set_index("_key")
    dup_rows_df = work.loc[dup_indices].drop(columns=["_filled_count"], errors="ignore") if dup_indices else work.iloc[0:0].drop(columns=["_filled_count"], errors="ignore")
    return resolved_df, dup_groups, dup_rows_df


def render_compare_and_merge(df_a, meta_a, name_a, key_prefix="cmp"):
    st.header("Compare with another file")
    st.caption("Upload a second spreadsheet to cross-check it against the file above — matched by a key column you pick, with numeric columns diffed automatically.")

    uploaded_b = st.file_uploader("Choose file to compare", type=["xlsx", "xls", "csv"], key=f"{key_prefix}_upload")
    if uploaded_b is None:
        return

    df_b = load_dataframe(uploaded_b, key_prefix=f"{key_prefix}_b")
    if df_b is None:
        return
    meta_b = classify_columns(df_b)
    name_b = uploaded_b.name

    st.write(f"**{name_b.upper()}** — {len(df_b):,} rows × {len(df_b.columns)} columns")

    default_key_a = guess_key_column(meta_a)
    default_key_b = guess_key_column(meta_b)
    c1, c2 = st.columns(2)
    with c1:
        key_a = st.selectbox(
            "Match rows using (this file)", list(df_a.columns),
            index=list(df_a.columns).index(default_key_a), key=f"{key_prefix}_key_a",
        )
    with c2:
        key_b = st.selectbox(
            "Match rows using (second file)", list(df_b.columns),
            index=list(df_b.columns).index(default_key_b), key=f"{key_prefix}_key_b",
        )

    resolved_a, dup_a, dup_rows_a = resolve_duplicates(df_a, key_a, meta_a[key_a].get("pad_width", 0))
    resolved_b, dup_b, dup_rows_b = resolve_duplicates(df_b, key_b, meta_b[key_b].get("pad_width", 0))

    keys_a, keys_b = set(resolved_a.index), set(resolved_b.index)
    matched = keys_a & keys_b
    only_a = keys_a - keys_b
    only_b = keys_b - keys_a

    stat_cols = st.columns(5)
    stats = [
        ("Rows in this file", f"{len(df_a):,}", f"{len(keys_a):,} with a usable key"),
        ("Rows in second file", f"{len(df_b):,}", f"{len(keys_b):,} with a usable key"),
        ("Matched", f"{len(matched):,}", "present in both files"),
        ("Only in this file", f"{len(only_a):,}", "not found in second file"),
        ("Only in second file", f"{len(only_b):,}", "not found in this file"),
    ]
    for widget, (label, value, sub) in zip(stat_cols, stats):
        with widget:
            st.metric(label, value)
            st.caption(sub)

    if dup_a or dup_b:
        parts = []
        if dup_a:
            parts.append(f"{len(dup_a)} duplicate key(s) in this file")
        if dup_b:
            parts.append(f"{len(dup_b)} duplicate key(s) in the second file")
        st.warning(
            "**Duplicates found:** " + " and ".join(parts) +
            ". Each duplicate was auto-merged into a single row — keeping the most complete "
            "entry and filling any blanks from the others. Use the duplicate-row downloads "
            "below if you'd rather review and merge them yourself."
        )

    dl_cols = st.columns(5)
    with dl_cols[0]:
        st.download_button(
            "Only in this file",
            to_display_df(resolved_a.loc[list(only_a)].reset_index(drop=True), meta_a).to_csv(index=False).encode("utf-8"),
            file_name=f"{slugify(name_a)}_only.csv", key=f"{key_prefix}_dl_only_a", disabled=not only_a,
        )
    with dl_cols[1]:
        st.download_button(
            "Only in second file",
            to_display_df(resolved_b.loc[list(only_b)].reset_index(drop=True), meta_b).to_csv(index=False).encode("utf-8"),
            file_name=f"{slugify(name_b)}_only.csv", key=f"{key_prefix}_dl_only_b", disabled=not only_b,
        )
    with dl_cols[2]:
        st.download_button(
            "Matched rows",
            to_display_df(resolved_a.loc[list(matched)].reset_index(drop=True), meta_a).to_csv(index=False).encode("utf-8"),
            file_name=f"{slugify(name_a)}_matched.csv", key=f"{key_prefix}_dl_matched", disabled=not matched,
        )
    with dl_cols[3]:
        st.download_button(
            "Duplicate rows (this file)",
            to_display_df(dup_rows_a.drop(columns=["_key"], errors="ignore"), meta_a).to_csv(index=False).encode("utf-8"),
            file_name=f"{slugify(name_a)}_duplicates.csv", key=f"{key_prefix}_dl_dup_a", disabled=not dup_a,
        )
    with dl_cols[4]:
        st.download_button(
            "Duplicate rows (second file)",
            to_display_df(dup_rows_b.drop(columns=["_key"], errors="ignore"), meta_b).to_csv(index=False).encode("utf-8"),
            file_name=f"{slugify(name_b)}_duplicates.csv", key=f"{key_prefix}_dl_dup_b", disabled=not dup_b,
        )

    # ---- Numeric column comparison ----
    st.subheader("Numeric column comparison")
    numeric_a = [c for c, m in meta_a.items() if m["type"] == "number" and not m.get("no_grouping")]
    numeric_b = [c for c, m in meta_b.items() if m["type"] == "number" and not m.get("no_grouping")]

    if not numeric_a or not numeric_b:
        st.info("No comparable numeric columns found in one or both files.")
    else:
        st.caption("Auto-matched by column name — change any pairing, or set it to \"Don't compare\" to drop it.")
        pairs = []
        for ca in numeric_a:
            best, best_score = None, 0
            for cb in numeric_b:
                score = header_token_score(ca, cb)
                if score > best_score:
                    best, best_score = cb, score
            options = ["Don't compare"] + numeric_b
            default_idx = options.index(best) if best else 0
            chosen = st.selectbox(f"{ca.upper()} compares to", options, index=default_idx, key=f"{key_prefix}_pair_{slugify(ca)}")
            if chosen != "Don't compare":
                pairs.append((ca, chosen))

        if not pairs:
            st.caption("Pick at least one column pair above to see a diff.")
        elif not matched:
            st.caption("No matched records to compare.")
        else:
            records = []
            for key in matched:
                row_a = resolved_a.loc[key]
                row_b = resolved_b.loc[key]
                rec = {"KEY": key}
                total_abs = 0.0
                for ca, cb in pairs:
                    va = row_a[ca] if not is_blank(row_a[ca]) else None
                    vb = row_b[cb] if not is_blank(row_b[cb]) else None
                    diff = (vb - va) if (va is not None and vb is not None) else None
                    rec[f"{ca.upper()} (A)"] = va
                    rec[f"{cb.upper()} (B) — vs {ca.upper()}"] = vb
                    rec[f"DIFF ({cb.upper()} − {ca.upper()})"] = diff
                    if diff is not None:
                        total_abs += abs(diff)
                rec["_total_abs"] = total_abs
                records.append(rec)
            diff_df = pd.DataFrame(records).sort_values("_total_abs", ascending=False).drop(columns=["_total_abs"])
            st.dataframe(diff_df, use_container_width=True, hide_index=True)
            st.download_button(
                "Download comparison CSV", diff_df.to_csv(index=False).encode("utf-8"),
                file_name="comparison.csv", key=f"{key_prefix}_dl_diff",
            )

    # ---- Merge ----
    st.subheader("Merged data")
    st.caption("Every record from both files, combined into one table using the key columns picked above.")

    a_other = [c for c in df_a.columns if c != key_a]
    b_other = [c for c in df_b.columns if c != key_b]
    a_names_norm = {re.sub(r"[^a-z0-9]", "", c.lower()) for c in df_a.columns}

    all_keys = sorted(keys_a | keys_b)
    merged_records = []
    for key in all_keys:
        row_a = resolved_a.loc[key] if key in keys_a else None
        row_b = resolved_b.loc[key] if key in keys_b else None
        rec = {key_a: key}
        for c in a_other:
            rec[c] = row_a[c] if row_a is not None else None
        for c in b_other:
            norm = re.sub(r"[^a-z0-9]", "", c.lower())
            col_name = f"{c} (FILE B)" if norm in a_names_norm else c
            rec[col_name] = row_b[c] if row_b is not None else None
        rec["SOURCE"] = "BOTH" if (row_a is not None and row_b is not None) else ("FILE A ONLY" if row_a is not None else "FILE B ONLY")
        merged_records.append(rec)
    merged_df = pd.DataFrame(merged_records)
    meta_merged = classify_columns(merged_df)
    merged_display = to_display_df(merged_df, meta_merged)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total merged records", f"{len(merged_df):,}")
    m2.metric("In both files", f"{len(matched):,}")
    m3.metric("Only in this file", f"{len(only_a):,}")
    m4.metric("Only in second file", f"{len(only_b):,}")

    search = st.text_input("Search merged data", key=f"{key_prefix}_merged_search")
    shown = merged_display[search_mask(merged_df, search)]
    st.write(f"{len(shown):,} of {len(merged_df):,} rows")
    st.dataframe(shown, use_container_width=True, hide_index=True)

    selected_cols = st.multiselect(
        "Columns to include in the download", list(merged_display.columns),
        default=list(merged_display.columns), key=f"{key_prefix}_merge_cols",
    )
    csv = merged_display[selected_cols].to_csv(index=False).encode("utf-8") if selected_cols else b""
    st.download_button(
        "Download merged CSV", csv,
        file_name=f"merged_{slugify(name_a)}_{slugify(name_b)}.csv",
        key=f"{key_prefix}_dl_merged", disabled=not selected_cols,
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run_app():
    """Entry point without page config, for embedding into a multi-tool app."""
    st.title("DataLens")
    st.write("Upload a spreadsheet and it automatically reads the columns, profiles them, and builds charts, an explorer, and a searchable table.")

    uploaded = st.file_uploader("Choose an Excel or CSV file", type=["xlsx", "xls", "csv"], key="datalens_main_upload")
    if uploaded is None:
        st.info("Upload a file to get started.")
        return

    df = load_dataframe(uploaded, key_prefix="main")
    if df is None:
        return
    meta = classify_columns(df)

    st.success(f"Loaded {len(df):,} rows × {len(df.columns)} columns from **{uploaded.name}**")

    with st.expander("🔍 Debug: raw column types (temporary, remove once the zero-padding bug is found)"):
        debug_col = st.selectbox("Column to inspect", list(df.columns), key="debug_col")
        debug_series = df[debug_col]
        st.write(f"pandas dtype: `{debug_series.dtype}`")
        m = meta.get(debug_col, {})
        st.write(f"classify_columns type: `{m.get('type')}` · pad_width: `{m.get('pad_width')}`")
        sample = debug_series.dropna().head(15)
        st.write("Raw values and their Python types (first 15 non-null rows):")
        st.table(pd.DataFrame({
            "repr(value)": [repr(v) for v in sample],
            "python type": [type(v).__name__ for v in sample],
        }))

    render_stats(df, meta)
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Column overview")
        st.caption("Detected automatically from the header row.")
        render_schema(meta)
    with col2:
        render_explorer(df, meta, key_prefix="main")

    st.divider()
    st.subheader("Category breakdown")
    render_category_charts(meta)

    st.divider()
    render_low_data(df, meta, key_prefix="main")

    st.divider()
    st.subheader("Data")
    search = st.text_input("Search all columns", key="main_search")
    display_df = to_display_df(df, meta)
    mask = search_mask(df, search)
    shown = display_df[mask]
    st.write(f"{len(shown):,} of {len(df):,} rows")
    st.dataframe(shown, use_container_width=True, hide_index=True)

    selected_cols = st.multiselect(
        "Columns to include in the download", list(display_df.columns),
        default=list(display_df.columns), key="main_dl_cols",
    )
    csv = display_df[selected_cols].to_csv(index=False).encode("utf-8") if selected_cols else b""
    st.download_button(
        "Download data as CSV", csv, file_name=f"{slugify(uploaded.name)}.csv",
        mime="text/csv", key="main_dl_data", disabled=not selected_cols,
    )

    st.divider()
    render_compare_and_merge(df, meta, uploaded.name, key_prefix="cmp")


def main():
    """Entry point when running this file directly."""
    st.set_page_config(page_title="DataLens", layout="wide")
    run_app()


if __name__ == "__main__":
    main()
