"""Dynamic Excel/CSV dashboard. Run with run_app.bat or streamlit run app.py."""

from datetime import timedelta
from hashlib import sha256
from io import BytesIO
from pathlib import Path
import re
import sqlite3

import pandas as pd
import plotly.express as px
import streamlit as st

COLORS = px.colors.qualitative.Safe


@st.cache_data(show_spinner=False)
def sheet_names(content: bytes, suffix: str):
    if suffix == ".csv":
        return []
    with pd.ExcelFile(BytesIO(content), engine="xlrd" if suffix == ".xls" else "openpyxl") as book:
        return book.sheet_names


@st.cache_data(show_spinner=False)
def load_data(content: bytes, suffix: str, sheet=0) -> pd.DataFrame:
    if suffix == ".csv":
        try:
            data = pd.read_csv(BytesIO(content))
        except UnicodeDecodeError:
            data = pd.read_csv(BytesIO(content), encoding="cp1252")
    else:
        data = pd.read_excel(BytesIO(content), sheet_name=sheet,
                             engine="xlrd" if suffix == ".xls" else "openpyxl")
    # Keep every source column, making blank/duplicate display names unambiguous.
    names, used = [], set()
    for i, col in enumerate(data.columns):
        base = str(col).strip() or f"Column {i + 1}"
        name, n = base, 2
        while name in used:
            name = f"{base} ({n})"
            n += 1
        names.append(name)
        used.add(name)
    data.columns = names
    if len(data.columns) == 0:
        raise ValueError("This sheet has no columns.")
    return data


def normalized(name):
    return re.sub(r"[^a-z0-9]+", "_", str(name).lower()).strip("_")


def infer_columns(data):
    """Return derived analysis series without altering source values in the table."""
    numbers, dates = {}, {}
    for col in data.columns:
        values = data[col]
        key = normalized(col)
        tokens = key.split("_")
        identifier = any(t in tokens for t in ("id", "code", "phone", "zip", "postal", "year"))
        if pd.api.types.is_datetime64_any_dtype(values):
            dates[col] = pd.to_datetime(values, errors="coerce", utc=True).dt.tz_convert(None)
        elif any(t in tokens for t in ("date", "datetime", "timestamp")) and not pd.api.types.is_numeric_dtype(values):
            parsed = pd.to_datetime(values, errors="coerce", format="mixed", utc=True).dt.tz_convert(None)
            if parsed.notna().any():
                dates[col] = parsed
        if pd.api.types.is_bool_dtype(values) or identifier or col in dates:
            continue
        if pd.api.types.is_numeric_dtype(values):
            numbers[col] = pd.to_numeric(values, errors="coerce").replace([float("inf"), -float("inf")], float("nan"))
        elif values.notna().any():
            converted = pd.to_numeric(values, errors="coerce")
            if converted[values.notna()].notna().all():
                numbers[col] = converted.replace([float("inf"), -float("inf")], float("nan"))
    return numbers, dates


def preferred(columns, aliases):
    for alias in aliases:
        for col in columns:
            if normalized(col) == alias:
                return col
    return None


def chart(fig):
    fig.update_layout(margin=dict(l=8, r=8, t=12, b=8), height=350, legend_title_text="")
    st.plotly_chart(fig, width="stretch")


def run_manual_query(data, query):
    """Query a private in-memory copy; permit reads only after loading it."""
    # Support SQL Server's common TOP syntax with SQLite's LIMIT equivalent.
    top = re.match(r"\s*SELECT\s+TOP\s*\(?([0-9]+)\)?\s+", query, re.IGNORECASE)
    if top:
        count = min(int(top.group(1)), 10000)
        query = "SELECT " + query[top.end():].strip().rstrip(";") + f" LIMIT {count}"
    with sqlite3.connect(":memory:") as connection:
        data.to_sql("data", connection, index=False)
        allowed = {sqlite3.SQLITE_SELECT, sqlite3.SQLITE_READ, sqlite3.SQLITE_FUNCTION,
                   sqlite3.SQLITE_RECURSIVE}
        connection.set_authorizer(lambda action, *args: sqlite3.SQLITE_OK
                                  if action in allowed else sqlite3.SQLITE_DENY)
        remaining = 1000
        def progress():
            nonlocal remaining
            remaining -= 1
            return int(remaining <= 0)
        connection.set_progress_handler(progress, 10000)
        cursor = connection.execute(query)
        if cursor.description is None:
            raise ValueError("Enter a SELECT query that returns rows.")
        rows = cursor.fetchmany(10001)
        return pd.DataFrame.from_records(rows[:10000], columns=[c[0] for c in cursor.description]), len(rows) > 10000


def main():
    st.set_page_config(page_title="Data dashboard", page_icon="📊", layout="wide")
    st.title("Data dashboard")
    with st.sidebar:
        st.header("Data and filters")
        upload = st.file_uploader("Upload Excel or CSV", type=["xlsx", "xlsm", "xls", "csv"])
    try:
        source = Path(__file__).with_name("motor_industry_sales_data.xlsx")
        content = upload.getvalue() if upload is not None else source.read_bytes()
        name = upload.name if upload is not None else source.name
        suffix = Path(name).suffix.lower()
        sheets = sheet_names(content, suffix)
        fingerprint = sha256(content).hexdigest()[:16]
        sheet = st.sidebar.selectbox("Worksheet", sheets, key=f"sheet_{fingerprint}") if sheets else 0
        data = load_data(content, suffix, sheet)
    except Exception as exc:
        st.error(f"Unable to read this file: {exc}")
        st.info("Upload a readable Excel/CSV file. Excel sheets should have column names in the first row.")
        st.stop()

    scope = f"{fingerprint}_{sheet}"
    st.caption(f"{name}" + (f" · {sheet}" if sheets else "") +
               f" · {len(data):,} rows · {len(data.columns):,} columns")
    if data.empty:
        st.info("This sheet contains no data rows. Select another worksheet or file.")
        st.dataframe(data, width="stretch", hide_index=True)
        st.stop()

    numbers, dates = infer_columns(data)
    categories = [col for col in data.columns if col not in numbers and col not in dates]
    amount_default = preferred(numbers, ["total_amount_pkr", "total_amount", "revenue", "sales", "amount", "total", "net_sales"])
    id_default = preferred(data.columns, ["order_id", "invoice_id", "transaction_id", "sale_id"])
    quantity_default = preferred(numbers, ["quantity", "qty", "units", "units_sold"])
    category_default = preferred(categories, ["vehicle_type", "category", "company", "department", "product"])
    group_default = preferred(categories, ["region", "country", "city", "branch"])

    def picker(label, options, default, key):
        choices = [None] + list(options)
        return st.selectbox(label, choices, index=choices.index(default) if default in choices else 0,
                            format_func=lambda v: "None" if v is None else str(v), key=f"{scope}_{key}")

    with st.sidebar:
        with st.expander("Detected columns / chart settings", expanded=amount_default is None):
            amount = picker("Numeric measure", numbers, amount_default, "measure")
            date = picker("Date column", dates, next(iter(dates), None), "date")
            order_id = picker("Order / transaction ID", data.columns, id_default, "id")
            quantity = picker("Quantity column", numbers, quantity_default, "quantity")
            category = picker("Category / breakdown", categories, category_default or next(iter(categories), None), "category")
            region = picker("Comparison group", categories, group_default, "group")
            aggregation = st.selectbox("Measure aggregation", ["Sum", "Mean", "Median", "Count"], key=f"{scope}_agg")
        defaults = [col for col in categories if 1 < data[col].nunique() <= 40][:6]
        filter_columns = st.multiselect("Filter columns", list(data.columns), default=defaults, key=f"{scope}_filters")

    mask = pd.Series(True, index=data.index)
    for col in filter_columns:
        values = data[col].astype("string").fillna("(Missing)")
        options = sorted(values.unique().tolist())
        with st.sidebar:
            if len(options) <= 200:
                chosen = st.multiselect(col, options, default=options, key=f"{scope}_filter_{col}")
                mask &= values.isin(chosen)
            else:
                text = st.text_input(f"{col} contains", key=f"{scope}_filter_text_{col}")
                if text:
                    mask &= values.str.contains(text, case=False, regex=False, na=False)

    if date:
        date_values = dates[date]
        valid = date_values.dropna()
        if not valid.empty:
            with st.sidebar:
                bounds = (valid.min().date(), valid.max().date())
                selected_dates = st.date_input("Date range", value=bounds, min_value=bounds[0],
                                               max_value=bounds[1], key=f"{scope}_{date}_range")
                keep_missing = st.checkbox("Include missing / invalid dates", value=True, key=f"{scope}_{date}_missing")
            if len(selected_dates) != 2:
                st.info("Select both a start and end date.")
                st.stop()
            lower = pd.Timestamp(selected_dates[0])
            upper = pd.Timestamp(selected_dates[1]).to_pydatetime() + timedelta(days=1)
            date_mask = date_values.ge(lower) & date_values.lt(upper)
            mask &= date_mask | (date_values.isna() & keep_missing)

    filtered = data.loc[mask]
    cards = st.columns(4)
    cards[0].metric("Rows", f"{len(filtered):,}")
    if order_id:
        cards[1].metric(f"Distinct {order_id}", f"{filtered[order_id].nunique():,}")
    else:
        cards[1].metric("Columns", f"{len(data.columns):,}")
    agg = aggregation.lower()
    measure_label = f"{aggregation} of {amount}" if amount else "Row count"
    if amount:
        numeric = numbers[amount].loc[mask]
        result = numeric.sum(min_count=1) if agg == "sum" else getattr(numeric, agg)()
        cards[2].metric(measure_label, f"{result:,.2f}" if pd.notna(result) else "—")
    else:
        cards[2].metric("Numeric columns", len(numbers))
    if quantity:
        total_quantity = numbers[quantity].loc[mask].sum(min_count=1)
        cards[3].metric(f"Total {quantity}", f"{total_quantity:,.0f}" if pd.notna(total_quantity) else "—")
    else:
        cards[3].metric("Missing cells", f"{int(filtered.isna().sum().sum()):,}")
    st.caption(f"{len(filtered):,} of {len(data):,} source rows. Charts use {measure_label.lower()}. "
               "All statuses are included unless filtered.")
    if amount and numbers[amount].loc[mask].isna().any():
        st.warning(f"{numbers[amount].loc[mask].isna().sum():,} missing or non-finite values in {amount} are excluded from numeric calculations.")

    # Internal chart names cannot collide with uploaded column names.
    plot = pd.DataFrame(index=filtered.index)
    if amount:
        plot["value"] = numbers[amount].loc[mask]

    def aggregate(frame, keys):
        if amount:
            grouped = frame.groupby(keys, dropna=False)["value"]
            result = grouped.sum(min_count=1) if agg == "sum" else getattr(grouped, agg)()
        else:
            result = frame.groupby(keys, dropna=False).size()
        return result.rename("measure").reset_index()

    if filtered.empty:
        st.info("No records match these filters.")
    else:
        if date:
            plot["month"] = dates[date].loc[mask].dt.to_period("M").dt.to_timestamp()
            monthly = aggregate(plot.loc[plot["month"].notna()], ["month"])
            if not monthly.empty:
                st.subheader(f"Monthly {measure_label.lower()}")
                chart(px.line(monthly, x="month", y="measure", markers=True,
                              labels={"month": date, "measure": measure_label}, color_discrete_sequence=[COLORS[0]]))
            else:
                st.info("No valid dates in the filtered records for a monthly chart.")
        if category:
            plot["category"] = filtered[category].astype("string").fillna("(Missing)")
            totals = aggregate(plot, ["category"]).sort_values("measure", ascending=False)
            st.subheader(f"{measure_label} by {category}")
            if len(totals) <= 12 and totals["measure"].ge(0).all() and totals["measure"].sum() > 0:
                chart(px.pie(totals, names="category", values="measure", hole=.4,
                             labels={"category": category, "measure": measure_label}, color_discrete_sequence=COLORS))
            else:
                st.caption("Showing the top 20 groups by the selected measure.")
                chart(px.bar(totals.head(20), x="category", y="measure",
                             labels={"category": category, "measure": measure_label}, color_discrete_sequence=[COLORS[0]]))
        if region and region != category:
            plot["group"] = filtered[region].astype("string").fillna("(Missing)")
            group_totals = aggregate(plot, ["group"]).sort_values("measure", ascending=False)
            top_groups = group_totals.head(20)["group"]
            comparison = plot.loc[plot["group"].isin(top_groups)]
            keys = ["group", "category"] if category else ["group"]
            grouped = aggregate(comparison, keys)
            st.subheader(f"{measure_label} by {region}")
            if len(group_totals) > 20:
                st.caption("Showing the top 20 comparison groups.")
            chart(px.bar(grouped, x="group", y="measure", color="category" if category else None,
                         barmode="group", color_discrete_sequence=COLORS,
                         labels={"group": region, "category": category or "Category", "measure": measure_label}))
        if not date and not category and not region:
            st.info("Choose a breakdown column in chart settings to create a chart. All source data is available below.")

    st.subheader("Manual SQL query")
    st.caption("Query the table named data. Local SQLite SQL supports SELECT, JOIN, GROUP BY, "
               "and window functions, plus SELECT TOP and [column names]. Full T-SQL procedures "
               "and SQL Server functions require a SQL Server connection.")
    use_filtered = st.checkbox("Apply dashboard filters to query data", value=True,
                               key=f"{scope}_sql_filtered")
    with st.expander("Available SQL columns"):
        st.code(", ".join("[" + c.replace("]", "]]") + "]" for c in data.columns), language="sql")
    with st.form(f"{scope}_sql_form"):
        query = st.text_area("SQL query", value="SELECT TOP 100 * FROM data;",
                             height=160, key=f"{scope}_sql_text")
        submitted = st.form_submit_button("Run query")
    if submitted:
        try:
            result, truncated = run_manual_query(filtered if use_filtered else data, query)
            st.session_state[f"{scope}_sql_result"] = (result, truncated)
        except Exception as exc:
            st.session_state.pop(f"{scope}_sql_result", None)
            st.error(f"Query failed: {exc}")
    saved_result = st.session_state.get(f"{scope}_sql_result")
    if saved_result is not None:
        result, truncated = saved_result
        st.caption(f"{len(result):,} result rows. Results reflect the last query run.")
        if truncated:
            st.info("Results limited to the first 10,000 rows.")
        st.dataframe(result, width="stretch", hide_index=True)
        st.download_button("Download query results", result.to_csv(index=False).encode("utf-8-sig"),
                           file_name="query_results.csv", mime="text/csv", key=f"{scope}_sql_download")

    st.subheader("Data preview")
    query = st.text_input("Search filtered records", placeholder="Search any column…", key=f"{scope}_search")
    preview = filtered
    if query:
        match = filtered.astype(str).apply(lambda column: column.str.contains(query, case=False, regex=False)).any(axis=1)
        preview = filtered.loc[match]
    st.caption(f"{len(preview):,} matching rows · All {len(data.columns)} source columns · Search applies to table and download.")
    st.dataframe(preview, width="stretch", hide_index=True)
    st.download_button("Download displayed records", preview.to_csv(index=False).encode("utf-8-sig"),
                       file_name="filtered_records.csv", mime="text/csv")
    with st.expander("Source and detected column types"):
        st.write("Column roles are detected from headers and values. Change them in chart settings. "
                 "No columns are required. Source values are preserved in the table and CSV download. "
                 "Numeric IDs are excluded from automatic numeric measures. Totals should only be used for additive measures; "
                 "select Mean or Median for rates and other non-additive values. No currency or customer acquisition is inferred.")
        st.dataframe(pd.DataFrame({"Column": data.columns,
                                  "Detected type": ["Date" if c in dates else "Numeric" if c in numbers else "Text / identifier" for c in data.columns],
                                  "Missing values": [int(data[c].isna().sum()) for c in data.columns]}),
                     width="stretch", hide_index=True)


if __name__ == "__main__":
    main()
