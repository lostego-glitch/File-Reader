
import base64
import io
import json
import math
import os
from datetime import date, datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from django.contrib.auth.decorators import login_required
from django.http import HttpResponse
from django.shortcuts import render


SESSION_DATA_KEY = "uploaded_excel"
SESSION_NAME_KEY = "uploaded_excel_name"
SUPPORTED_EXTENSIONS = (".xlsx", ".xls", ".csv", ".json")


def clean_column_name(value):
    name = str(value).strip()
    return name if name else "Unnamed"


def make_unique_columns(columns):
    result = []
    used = {}

    for column in columns:
        name = clean_column_name(column)
        if name not in used:
            used[name] = 0
            result.append(name)
        else:
            used[name] += 1
            candidate = f"{name}_{used[name]}"
            while candidate in used:
                used[name] += 1
                candidate = f"{name}_{used[name]}"
            used[candidate] = 0
            result.append(candidate)

    return result


def json_safe_value(value):
    if value is None:
        return None

    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()

    if isinstance(value, np.generic):
        value = value.item()

    if isinstance(value, (dict, list, tuple, set)):
        try:
            return json.dumps(value, ensure_ascii=False, default=str)
        except (TypeError, ValueError):
            return str(value)

    if isinstance(value, float) and not math.isfinite(value):
        return None

    try:
        missing = pd.isna(value)
        if isinstance(missing, (bool, np.bool_)) and missing:
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, (str, int, float, bool)):
        return value

    return str(value)


def dataframe_to_records(df):
    if df is None or df.empty:
        return []

    records = []

    for _, row in df.iterrows():
        records.append({
            str(column): json_safe_value(row[column])
            for column in df.columns
        })

    return records


def normalize_nested_values(df):
    if df is None or df.empty:
        return df

    result = df.copy()

    for column in result.columns:
        result[column] = result[column].map(
            lambda value: (
                json.dumps(value, ensure_ascii=False, default=str)
                if isinstance(value, (dict, list, tuple, set))
                else value
            )
        )

    return result


def dataframe_to_session(request, uploaded_file):
    uploaded_file.seek(0)
    data = uploaded_file.read()

    request.session[SESSION_DATA_KEY] = base64.b64encode(data).decode("ascii")
    request.session[SESSION_NAME_KEY] = os.path.basename(uploaded_file.name)
    request.session.modified = True


def save_dataframe_to_session(request, uploaded_file):
    dataframe_to_session(request, uploaded_file)


def get_uploaded_bytes(request):
    encoded = request.session.get(SESSION_DATA_KEY)

    if not encoded:
        return None

    try:
        return base64.b64decode(encoded)
    except (ValueError, TypeError):
        return None


def get_file_extension(request):
    filename = request.session.get(SESSION_NAME_KEY, "")
    return os.path.splitext(filename)[1].lower()


def get_workbook(request):
    data = get_uploaded_bytes(request)
    return io.BytesIO(data) if data is not None else None


def get_sheet_names(request):
    if get_file_extension(request) not in (".xlsx", ".xls"):
        return []

    workbook = get_workbook(request)

    if workbook is None:
        return []

    try:
        workbook.seek(0)

        with pd.ExcelFile(workbook) as excel:
            return excel.sheet_names
    except Exception:
        return []


def json_records_to_dataframe(records, column_name="value"):
    if records is None:
        return pd.DataFrame()

    if isinstance(records, dict):
        return normalize_nested_values(pd.json_normalize(records, sep="."))

    if not isinstance(records, list):
        return pd.DataFrame({column_name: [records]})

    if not records:
        return pd.DataFrame()

    if all(isinstance(item, dict) for item in records):
        return normalize_nested_values(pd.json_normalize(records, sep="."))

    if all(isinstance(item, (list, tuple)) for item in records):
        max_length = max((len(item) for item in records), default=0)

        padded_records = [
            list(item) + [None] * (max_length - len(item))
            for item in records
        ]

        return normalize_nested_values(pd.DataFrame(padded_records))

    if any(isinstance(item, dict) for item in records):
        normalized = []

        for item in records:
            if isinstance(item, dict):
                normalized.append(item)
            else:
                normalized.append({column_name: item})

        return normalize_nested_values(pd.json_normalize(normalized, sep="."))

    return pd.DataFrame({column_name: records})


def read_json_dataframe(data):
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("latin-1")

    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        records = []

        try:
            for line in text.splitlines():
                if line.strip():
                    records.append(json.loads(line))
        except json.JSONDecodeError as exc:
            raise ValueError(
                f"The JSON file is invalid or could not be parsed: {exc}"
            ) from exc

        if not records:
            raise ValueError("The JSON file contains no readable records.")

        return json_records_to_dataframe(records)

    if isinstance(payload, list):
        return json_records_to_dataframe(payload)

    if not isinstance(payload, dict):
        return pd.DataFrame({"value": [payload]})

    preferred_keys = [
        "source_data",
        "records",
        "rows",
        "items",
        "results",
        "airports",
        "data",
        "filtered_data",
        "pivot_table",
    ]

    for key in preferred_keys:
        if key not in payload:
            continue

        value = payload[key]

        if isinstance(value, list):
            return json_records_to_dataframe(value, key)

        if isinstance(value, dict):
            for nested_key in preferred_keys:
                nested_value = value.get(nested_key)

                if isinstance(nested_value, (list, dict)):
                    return json_records_to_dataframe(nested_value, nested_key)

            return normalize_nested_values(pd.json_normalize(value, sep="."))

    return normalize_nested_values(pd.json_normalize(payload, sep="."))


def load_dataframe(request, sheet_name=None):
    data = get_uploaded_bytes(request)

    if data is None:
        return pd.DataFrame()

    extension = get_file_extension(request)
    buffer = io.BytesIO(data)

    if extension in (".xlsx", ".xls"):
        sheets = get_sheet_names(request)

        if not sheets:
            raise ValueError(
                "The Excel workbook does not contain a readable worksheet."
            )

        if sheet_name not in sheets:
            sheet_name = sheets[0]

        buffer.seek(0)
        df = pd.read_excel(buffer, sheet_name=sheet_name)

    elif extension == ".csv":
        try:
            df = pd.read_csv(buffer, encoding="utf-8-sig")
        except UnicodeDecodeError:
            buffer.seek(0)
            df = pd.read_csv(buffer, encoding="latin-1")

    elif extension == ".json":
        df = read_json_dataframe(data)

    else:
        raise ValueError(
            "Unsupported file type. Upload an .xlsx, .xls, .csv, or .json file."
        )

    df.columns = make_unique_columns(df.columns)
    df = normalize_nested_values(df)
    df = df.dropna(how="all").reset_index(drop=True)

    return df


def get_field_types(df):
    field_types = {}

    for column in df.columns:
        series = df[column]

        if pd.api.types.is_numeric_dtype(series):
            field_types[column] = "numeric"
            continue

        if pd.api.types.is_datetime64_any_dtype(series):
            field_types[column] = "date"
            continue

        nonempty = series.dropna()

        if nonempty.empty:
            field_types[column] = "category"
            continue

        numeric = pd.to_numeric(nonempty, errors="coerce")

        if numeric.notna().sum() >= max(1, int(len(nonempty) * 0.8)):
            field_types[column] = "numeric"
            continue

        dates = pd.to_datetime(nonempty, errors="coerce")

        if dates.notna().sum() >= max(1, int(len(nonempty) * 0.8)):
            field_types[column] = "date"
        else:
            field_types[column] = "category"

    return field_types


def get_filter_control_data(df, filter_fields):
    controls = {}

    for field in filter_fields:
        if field not in df.columns:
            continue

        unique_values = []

        for value in df[field].dropna().tolist():
            if isinstance(value, (pd.Timestamp, datetime, date)):
                display = value.strftime("%Y-%m-%d")
            else:
                display = str(value)

            if display not in unique_values:
                unique_values.append(display)

        controls[field] = sorted(unique_values)[:500]

    return controls


def apply_filters(df, filter_fields, selected_filters):
    filtered = df.copy()

    for field in filter_fields:
        if field not in filtered.columns:
            continue

        selected = selected_filters.get(field, [])

        if not selected:
            continue

        values = filtered[field].apply(
            lambda value: value.strftime("%Y-%m-%d")
            if isinstance(value, (pd.Timestamp, datetime, date))
            else str(value)
        )

        filtered = filtered[values.isin(selected)]

    return filtered


def perform_pivot(df, row_fields, column_fields, value_fields, aggregation):
    if not value_fields:
        return pd.DataFrame()

    valid_rows = [field for field in row_fields if field in df.columns]
    valid_columns = [field for field in column_fields if field in df.columns]
    valid_values = [field for field in value_fields if field in df.columns]

    if not valid_values:
        return pd.DataFrame()

    working = normalize_nested_values(df.copy())

    agg_map = {
        "count": "count",
        "average": "mean",
        "mean": "mean",
        "sum": "sum",
        "min": "min",
        "max": "max",
    }

    aggfunc = agg_map.get(aggregation, "count")

    if aggregation != "count":
        for field in valid_values:
            working[field] = pd.to_numeric(working[field], errors="coerce")

    if not valid_rows and not valid_columns:
        result = {}

        for field in valid_values:
            series = working[field]

            if aggregation == "count":
                value = series.count()
            elif aggregation in ("average", "mean"):
                value = series.mean()
            elif aggregation == "sum":
                value = series.sum()
            elif aggregation == "min":
                value = series.min()
            else:
                value = series.max()

            result[field] = [value]

        return pd.DataFrame(result)

    try:
        pivot = pd.pivot_table(
            working,
            index=valid_rows or None,
            columns=valid_columns or None,
            values=valid_values,
            aggfunc=aggfunc,
            dropna=False,
            fill_value=0,
        )
    except Exception:
        return pd.DataFrame()

    if isinstance(pivot, pd.Series):
        pivot = pivot.to_frame()

    pivot = pivot.reset_index()

    if isinstance(pivot.columns, pd.MultiIndex):
        flattened = []

        for column in pivot.columns:
            parts = [
                str(part).strip()
                for part in column
                if str(part).strip()
                and str(part).strip().lower() != "nan"
            ]

            flattened.append(" - ".join(parts))

        pivot.columns = make_unique_columns(flattened)
    else:
        pivot.columns = make_unique_columns(pivot.columns)

    return pivot.reset_index(drop=True)


def numeric_result_columns(pivot_df, grouping_columns):
    columns = []

    for column in pivot_df.columns:
        if column in grouping_columns:
            continue

        values = pd.to_numeric(pivot_df[column], errors="coerce")

        if values.notna().any():
            columns.append(column)

    return columns


def performance_field(field_name):
    name = str(field_name).lower()

    keywords = [
        "grade",
        "score",
        "gpa",
        "mark",
        "rating",
        "attendance",
        "performance",
        "result",
        "percentage",
        "percent",
    ]

    return any(keyword in name for keyword in keywords)


def generate_pivot_insights(
    pivot_df,
    row_fields,
    column_fields,
    value_fields,
    aggregation,
):
    insights = []

    if pivot_df is None or pivot_df.empty:
        return insights

    grouping_columns = [
        column for column in row_fields if column in pivot_df.columns
    ]

    numeric_columns = numeric_result_columns(pivot_df, grouping_columns)

    if not numeric_columns:
        return insights

    if grouping_columns:
        group_column = grouping_columns[0]

        for numeric_column in numeric_columns[:10]:
            working = pivot_df[[group_column, numeric_column]].copy()
            working[numeric_column] = pd.to_numeric(
                working[numeric_column], errors="coerce"
            )
            working = working.dropna(subset=[numeric_column])

            if working.empty:
                continue

            working = working.sort_values(numeric_column, ascending=False)

            highest = working.iloc[0]
            lowest = working.iloc[-1]

            high_group = str(highest[group_column])
            low_group = str(lowest[group_column])
            high_value = float(highest[numeric_column])
            low_value = float(lowest[numeric_column])

            if len(working) > 1:
                metric_names = {
                    "average": f"average {numeric_column}",
                    "mean": f"average {numeric_column}",
                    "count": f"count of {numeric_column}",
                    "sum": f"total {numeric_column}",
                    "min": f"minimum {numeric_column}",
                    "max": f"maximum {numeric_column}",
                }

                metric = metric_names.get(aggregation, numeric_column)

                insights.append(
                    f"{high_group} has the highest {metric}, at "
                    f"{high_value:,.2f}, while {low_group} has the lowest, "
                    f"at {low_value:,.2f}. The difference is "
                    f"{high_value - low_value:,.2f}."
                )

                value_field_name = (
                    value_fields[0] if value_fields else numeric_column
                )

                if performance_field(value_field_name):
                    insights.append(
                        f"Based on {value_field_name}, {high_group} is the "
                        f"highest-performing group, while {low_group} is "
                        "the lowest-performing group."
                    )

                if aggregation in ("average", "mean"):
                    overall = float(working[numeric_column].mean())

                    insights.append(
                        f"The overall group average for {numeric_column} "
                        f"is {overall:,.2f}."
                    )

                if len(working) >= 3:
                    second = working.iloc[1]

                    insights.append(
                        f"The second-highest group is {second[group_column]}, "
                        f"with {float(second[numeric_column]):,.2f}."
                    )
            else:
                insights.append(
                    f"{high_group} has the available result for "
                    f"{numeric_column}, at {high_value:,.2f}."
                )
    else:
        for numeric_column in numeric_columns[:10]:
            values = pd.to_numeric(
                pivot_df[numeric_column], errors="coerce"
            ).dropna()

            if values.empty:
                continue

            value = float(values.iloc[0])

            if aggregation in ("average", "mean"):
                message = (
                    f"The overall average {numeric_column} is {value:,.2f}."
                )
            elif aggregation == "count":
                message = (
                    f"The total count of {numeric_column} is {value:,.0f}."
                )
            elif aggregation == "sum":
                message = f"The total {numeric_column} is {value:,.2f}."
            elif aggregation == "min":
                message = f"The minimum {numeric_column} is {value:,.2f}."
            else:
                message = f"The maximum {numeric_column} is {value:,.2f}."

            insights.append(message)

    if column_fields and len(insights) < 10:
        insights.append(
            "The results are separated by "
            f"{', '.join(column_fields)}, allowing the categories to be "
            "compared across the row groups."
        )

    return insights


def generate_dataset_insights(df):
    insights = []

    if df is None or df.empty:
        return insights

    insights.append(
        f"The dataset contains {len(df):,} rows and {len(df.columns):,} columns."
    )

    duplicate_count = int(df.duplicated().sum())

    if duplicate_count:
        insights.append(
            f"The dataset contains {duplicate_count:,} duplicate rows."
        )
    else:
        insights.append("No duplicate rows were detected.")

    missing_counts = df.isna().sum()
    missing_counts = missing_counts[missing_counts > 0]

    if not missing_counts.empty:
        for column, count in missing_counts.head(10).items():
            insights.append(
                f"{column} contains {int(count):,} missing values."
            )
    else:
        insights.append("No missing values were detected.")

    field_types = get_field_types(df)

    numeric_fields = [
        field for field, field_type in field_types.items()
        if field_type == "numeric"
    ]

    category_fields = [
        field for field, field_type in field_types.items()
        if field_type == "category"
    ]

    for field in numeric_fields[:10]:
        values = pd.to_numeric(df[field], errors="coerce").dropna()

        if values.empty:
            continue

        insights.append(
            f"{field}: average {values.mean():,.2f}, "
            f"minimum {values.min():,.2f}, "
            f"maximum {values.max():,.2f}."
        )

    for category in category_fields[:5]:
        counts = df[category].dropna().astype(str).value_counts()

        if counts.empty:
            continue

        top_category = str(counts.index[0])

        insights.append(
            f"{category} has {counts.nunique():,} distinct frequency "
            f"levels among its non-empty values; the most frequent value "
            f"is {top_category} ({int(counts.iloc[0]):,} records)."
        )

    performance_columns = [
        field for field in numeric_fields if performance_field(field)
    ]

    for metric in performance_columns[:5]:
        for category in category_fields[:5]:
            grouped_data = pd.DataFrame({
                "category": df[category].astype(str),
                "metric": pd.to_numeric(df[metric], errors="coerce"),
            }).dropna(subset=["metric"])

            if grouped_data.empty:
                continue

            grouped = grouped_data.groupby(
                "category", dropna=False
            )["metric"].mean()

            if len(grouped) < 2:
                continue

            highest = grouped.idxmax()
            lowest = grouped.idxmin()

            insights.append(
                f"Average {metric} is highest for {category} "
                f"'{highest}' ({grouped.max():,.2f}) and lowest for "
                f"'{lowest}' ({grouped.min():,.2f})."
            )

    return insights[:30]


def create_chart_dataframe(pivot_df, row_fields):
    if pivot_df is None or pivot_df.empty:
        return pd.DataFrame()

    grouping_columns = [
        column for column in row_fields if column in pivot_df.columns
    ]

    index_column = (
        grouping_columns[0] if grouping_columns else pivot_df.columns[0]
    )

    chart_df = pivot_df.copy().set_index(index_column)
    numeric_columns = []

    for column in chart_df.columns:
        converted = pd.to_numeric(chart_df[column], errors="coerce")

        if converted.notna().any():
            chart_df[column] = converted
            numeric_columns.append(column)

    if not numeric_columns:
        return pd.DataFrame()

    return chart_df[numeric_columns]


def create_chart_figure(chart_df, chart_type, chart_title):
    if chart_df is None or chart_df.empty:
        return None

    figure, axis = plt.subplots(figsize=(11, 6.5), dpi=120)

    if chart_type == "horizontal_bar":
        chart_df.plot(kind="barh", ax=axis)

    elif chart_type == "line":
        chart_df.plot(kind="line", marker="o", ax=axis)

    elif chart_type == "area":
        chart_df.plot(kind="area", alpha=0.65, ax=axis)

    elif chart_type == "pie":
        values = pd.to_numeric(
            chart_df.iloc[:, 0], errors="coerce"
        ).fillna(0)

        labels = [str(value) for value in chart_df.index]

        if values.sum() <= 0:
            plt.close(figure)
            return None

        axis.pie(values, labels=labels, autopct="%1.1f%%")
        axis.set_ylabel("")

    elif chart_type == "scatter":
        numeric_data = chart_df.apply(pd.to_numeric, errors="coerce")

        if numeric_data.shape[1] >= 2:
            x = numeric_data.iloc[:, 0]
            y = numeric_data.iloc[:, 1]
            valid = x.notna() & y.notna()

            axis.scatter(x[valid], y[valid])
            axis.set_xlabel(str(numeric_data.columns[0]))
            axis.set_ylabel(str(numeric_data.columns[1]))
        else:
            x = np.arange(len(numeric_data))
            y = numeric_data.iloc[:, 0]
            valid = y.notna()

            axis.scatter(x[valid], y[valid])
            axis.set_xlabel(str(chart_df.index.name or "Category"))
            axis.set_ylabel(str(numeric_data.columns[0]))

    else:
        chart_df.plot(kind="bar", ax=axis)

    axis.set_title(chart_title or "Pivot Chart")

    if chart_type != "pie":
        axis.set_xlabel(str(chart_df.index.name or "Category"))
        axis.grid(axis="y", alpha=0.2)

    figure.tight_layout()

    return figure


def create_chart_image(figure):
    if figure is None:
        return ""

    buffer = io.BytesIO()
    figure.savefig(buffer, format="png", bbox_inches="tight")
    plt.close(figure)
    buffer.seek(0)

    return base64.b64encode(buffer.read()).decode("ascii")


def normalize_financial_text(value):
    text = str(value).strip().lower()

    for old, new in [
        ("₱", ""),
        ("$", ""),
        ("€", ""),
        ("£", ""),
        (",", ""),
        ("(", "-"),
        (")", ""),
    ]:
        text = text.replace(old, new)

    return " ".join(text.split())


def financial_number(value):
    if value is None:
        return None

    if isinstance(value, (int, float, np.integer, np.floating)):
        number = float(value)
        return number if math.isfinite(number) else None

    text = normalize_financial_text(value)

    if not text:
        return None

    try:
        number = float(text)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError):
        return None


def calculate_financial_ratios(df):
    if df is None or df.empty:
        return []

    accounts = {}
    account_column = None

    account_names = {
        "account",
        "accounts",
        "account name",
        "account names",
        "particulars",
        "description",
        "item",
        "financial account",
    }

    for column in df.columns:
        if normalize_financial_text(column) in account_names:
            account_column = column
            break

    amount_column = None

    if account_column is not None:
        amount_candidates = [
            "amount", "value", "balance", "total", "2025", "2024", "2023"
        ]

        for candidate in amount_candidates:
            for column in df.columns:
                if (
                    column != account_column
                    and normalize_financial_text(column) == candidate
                ):
                    amount_column = column
                    break

            if amount_column is not None:
                break

        if amount_column is None:
            numeric_candidates = []

            for column in df.columns:
                if column == account_column:
                    continue

                count = df[column].apply(financial_number).notna().sum()

                if count:
                    numeric_candidates.append((column, count))

            if numeric_candidates:
                amount_column = max(
                    numeric_candidates, key=lambda item: item[1]
                )[0]

        if amount_column is not None:
            for _, row in df.iterrows():
                account = normalize_financial_text(row[account_column])
                amount = financial_number(row[amount_column])

                if account and amount is not None:
                    accounts[account] = amount

    if not accounts:
        for column in df.columns:
            values = df[column].apply(financial_number).dropna()

            if not values.empty:
                accounts[normalize_financial_text(column)] = float(
                    values.sum()
                )

    def find_account(names):
        normalized_names = [
            normalize_financial_text(name) for name in names
        ]

        for name in normalized_names:
            if name in accounts:
                return accounts[name]

        for account_name, value in accounts.items():
            for name in normalized_names:
                if account_name.startswith(name + " "):
                    return value

        return None

    revenue = find_account([
        "revenue", "sales revenue", "total revenue", "net sales", "sales"
    ])
    gross_profit = find_account(["gross profit"])
    operating_income = find_account([
        "operating income", "operating profit"
    ])
    net_income = find_account([
        "net income", "net profit", "profit after tax", "profit after taxes"
    ])
    current_assets = find_account(["current assets", "total current assets"])
    current_liabilities = find_account([
        "current liabilities", "total current liabilities"
    ])
    total_assets = find_account(["total assets"])
    total_liabilities = find_account(["total liabilities"])
    equity = find_account([
        "equity",
        "total equity",
        "shareholders equity",
        "shareholders' equity",
        "stockholders equity",
        "stockholders' equity",
        "owners equity",
        "owner's equity",
    ])
    inventory = find_account([
        "inventory", "inventories", "merchandise inventory"
    ])
    cash = find_account([
        "cash and cash equivalents", "cash equivalents", "cash"
    ])
    receivables = find_account([
        "accounts receivable",
        "account receivable",
        "trade receivables",
        "receivables",
    ])

    ratios = []

    def add(name, value, interpretation, suffix=""):
        if value is not None and math.isfinite(float(value)):
            ratios.append({
                "name": name,
                "value": round(float(value), 4),
                "display_value": f"{float(value):,.2f}{suffix}",
                "interpretation": interpretation,
            })

    if current_assets is not None and current_liabilities not in (None, 0):
        add(
            "Current Ratio",
            current_assets / current_liabilities,
            "Measures the ability to cover current liabilities using current assets.",
        )

    if cash is not None and current_liabilities not in (None, 0):
        add(
            "Cash Ratio",
            cash / current_liabilities,
            "Measures the ability to cover current liabilities using cash.",
        )

    if (
        current_assets is not None
        and inventory is not None
        and current_liabilities not in (None, 0)
    ):
        quick_assets = current_assets - inventory

        if cash is not None and receivables is not None:
            quick_assets = cash + receivables

        add(
            "Quick Ratio",
            quick_assets / current_liabilities,
            "Measures the ability to cover current liabilities using liquid current assets.",
        )

    if inventory is not None and current_liabilities not in (None, 0):
        add(
            "Inventory to Current Liabilities",
            inventory / current_liabilities,
            "Compares inventory with current liabilities.",
        )

    if gross_profit is not None and revenue not in (None, 0):
        add(
            "Gross Profit Margin",
            gross_profit / revenue * 100,
            "Shows the percentage of revenue remaining after cost of goods sold.",
            "%",
        )

    if operating_income is not None and revenue not in (None, 0):
        add(
            "Operating Profit Margin",
            operating_income / revenue * 100,
            "Shows operating income generated from revenue.",
            "%",
        )

    if net_income is not None and revenue not in (None, 0):
        add(
            "Net Profit Margin",
            net_income / revenue * 100,
            "Shows the percentage of revenue remaining as net income.",
            "%",
        )

    if net_income is not None and total_assets not in (None, 0):
        add(
            "Return on Assets",
            net_income / total_assets * 100,
            "Measures profit generated from total assets.",
            "%",
        )

    if net_income is not None and equity not in (None, 0):
        add(
            "Return on Equity",
            net_income / equity * 100,
            "Measures profit generated from shareholders' equity.",
            "%",
        )

    if total_liabilities is not None and total_assets not in (None, 0):
        add(
            "Debt Ratio",
            total_liabilities / total_assets * 100,
            "Shows the percentage of assets financed by liabilities.",
            "%",
        )

    if total_liabilities is not None and equity not in (None, 0):
        add(
            "Debt-to-Equity Ratio",
            total_liabilities / equity,
            "Compares liabilities with shareholders' equity.",
        )

    if equity is not None and total_assets not in (None, 0):
        add(
            "Equity Ratio",
            equity / total_assets * 100,
            "Shows the percentage of assets financed by equity.",
            "%",
        )

    return ratios


def parse_list(request, key):
    values = request.POST.getlist(key)
    result = []

    for value in values:
        if "," in value:
            result.extend(
                part.strip() for part in value.split(",") if part.strip()
            )
        elif value.strip():
            result.append(value.strip())

    return result


def parse_selected_filters(request, filter_fields):
    return {
        field: request.POST.getlist(f"filter_value_{field}")
        for field in filter_fields
    }


def read_raw_dataframe(request, sheet_name=None):
    data = get_uploaded_bytes(request)

    if data is None:
        return pd.DataFrame()

    extension = get_file_extension(request)
    buffer = io.BytesIO(data)

    if extension in (".xlsx", ".xls"):
        sheets = get_sheet_names(request)

        if not sheets:
            raise ValueError(
                "The Excel workbook does not contain a readable worksheet."
            )

        if sheet_name not in sheets:
            sheet_name = sheets[0]

        return pd.read_excel(
            buffer, sheet_name=sheet_name, header=None, dtype=object
        )

    if extension == ".csv":
        try:
            return pd.read_csv(
                buffer,
                encoding="utf-8-sig",
                header=None,
                dtype=object,
                keep_default_na=True,
            )
        except UnicodeDecodeError:
            buffer.seek(0)
            return pd.read_csv(
                buffer,
                encoding="latin-1",
                header=None,
                dtype=object,
                keep_default_na=True,
            )

    if extension == ".json":
        return read_json_dataframe(data)

    raise ValueError("Unsupported file type.")


def _cell_is_blank(value):
    if value is None:
        return True

    if isinstance(value, (dict, list, tuple, set)):
        return False

    try:
        missing = pd.isna(value)

        if isinstance(missing, (bool, np.bool_)) and missing:
            return True
    except (TypeError, ValueError):
        pass

    return isinstance(value, str) and not value.strip()


def detect_header_row(raw_df):
    if raw_df is None or raw_df.empty:
        return 0

    limit = min(len(raw_df), 10)

    for row_index in range(1, limit):
        previous = raw_df.iloc[:row_index]

        preceding_nonblank = sum(
            not _cell_is_blank(value)
            for value in previous.to_numpy().ravel()
        )

        preceding_total = max(1, previous.size)

        if preceding_nonblank / preceding_total > 0.10:
            continue

        row = raw_df.iloc[row_index]
        nonblank = [
            value for value in row.tolist()
            if not _cell_is_blank(value)
        ]

        if len(nonblank) < 2:
            continue

        unique_ratio = (
            len({str(value).strip().lower() for value in nonblank})
            / max(1, len(nonblank))
        )

        if unique_ratio >= 0.7:
            return row_index

    return 0


def dataframe_from_raw(raw_df, header_row=0):
    if raw_df is None or raw_df.empty:
        return pd.DataFrame()

    header_row = max(0, min(int(header_row), len(raw_df) - 1))
    headers = make_unique_columns(raw_df.iloc[header_row].tolist())

    data = raw_df.iloc[header_row + 1:].copy()
    data.columns = headers
    data = data.dropna(how="all")

    for column in data.columns:
        data[column] = data[column].map(
            lambda value: value.strip() if isinstance(value, str) else value
        )

    return data.reset_index(drop=True)


def infer_cleaner_types(df):
    inferred = {}
    mismatches = {}

    for column in df.columns:
        series = df[column]
        nonempty = series[~series.map(_cell_is_blank)]

        if nonempty.empty:
            inferred[column] = "text"
            mismatches[column] = []
            continue

        numeric_text = (
            nonempty.astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("₱", "", regex=False)
            .str.replace("$", "", regex=False)
            .str.strip()
        )

        numeric = pd.to_numeric(numeric_text, errors="coerce")
        numeric_ratio = float(numeric.notna().mean())

        if numeric_ratio >= 0.8:
            inferred[column] = "numeric"
            bad = []

            for index, value in series.items():
                if _cell_is_blank(value):
                    continue

                text = (
                    str(value)
                    .replace(",", "")
                    .replace("₱", "")
                    .replace("$", "")
                    .strip()
                )

                if pd.isna(pd.to_numeric(text, errors="coerce")):
                    bad.append(index)

            mismatches[column] = bad
            continue

        dates = pd.to_datetime(nonempty, errors="coerce")
        date_ratio = float(dates.notna().mean())

        if date_ratio >= 0.8:
            inferred[column] = "date"
            bad = []

            for index, value in series.items():
                if (
                    not _cell_is_blank(value)
                    and pd.isna(pd.to_datetime(value, errors="coerce"))
                ):
                    bad.append(index)

            mismatches[column] = bad
        else:
            inferred[column] = "text"
            mismatches[column] = []

    return inferred, mismatches


def scan_data_quality(df, header_row=0):
    missing_cells = []

    for index, row in df.iterrows():
        for column in df.columns:
            if _cell_is_blank(row[column]):
                missing_cells.append({
                    "row": int(index) + 1,
                    "column": str(column),
                })

    inferred, mismatch_map = infer_cleaner_types(df)
    mismatches = []

    for column, indices in mismatch_map.items():
        for index in indices:
            mismatches.append({
                "row": int(index) + 1,
                "column": str(column),
                "value": str(df.at[index, column])[:120],
            })

    return {
        "header_row": int(header_row),
        "header_misplaced": int(header_row) > 0,
        "row_count": int(len(df)),
        "column_count": int(len(df.columns)),
        "missing_count": len(missing_cells),
        "missing_cells": missing_cells[:100],
        "mismatch_count": len(mismatches),
        "mismatches": mismatches[:100],
        "inferred_types": inferred,
        "empty_rows": int(sum(
            all(_cell_is_blank(value) for value in row)
            for row in df.to_numpy().tolist()
        )),
        "duplicate_rows": int(df.duplicated().sum()),
    }


def context_missing_value(column, series):
    name = str(column).lower()

    zero_terms = (
        "quantity",
        "qty",
        "units",
        "stock",
        "count",
        "sold",
        "returned",
        "adjusted",
        "attendance count",
    )

    if any(term in name for term in zero_terms):
        return 0

    return ""


def clean_dataframe(df, report, options):
    cleaned = df.copy(deep=True)

    if options.get("fill_missing") == "context":
        for column in cleaned.columns:
            replacement = context_missing_value(column, cleaned[column])
            cleaned[column] = cleaned[column].map(
                lambda value: (
                    replacement if _cell_is_blank(value) else value
                )
            )

    elif options.get("fill_missing") == "blank":
        cleaned = cleaned.map(
            lambda value: "" if _cell_is_blank(value) else value
        )

    elif options.get("fill_missing") == "zero":
        cleaned = cleaned.map(
            lambda value: 0 if _cell_is_blank(value) else value
        )

    inferred, mismatch_map = infer_cleaner_types(df)

    missing_row_mask = df.apply(
        lambda row: any(_cell_is_blank(value) for value in row),
        axis=1,
    )

    mismatch_row_mask = pd.Series(False, index=df.index)

    for indices in mismatch_map.values():
        if indices:
            mismatch_row_mask.loc[indices] = True

    if options.get("remove_empty_rows", "yes") == "yes":
        nonempty_mask = ~df.apply(
            lambda row: all(_cell_is_blank(value) for value in row),
            axis=1,
        )

        cleaned = cleaned.loc[nonempty_mask].copy()
        missing_row_mask = missing_row_mask.loc[nonempty_mask]
        mismatch_row_mask = mismatch_row_mask.loc[nonempty_mask]

    delete_mask = pd.Series(False, index=cleaned.index)

    if options.get("missing_rows") == "delete":
        delete_mask |= missing_row_mask.reindex(
            cleaned.index, fill_value=False
        )

    if options.get("mismatch_rows") == "delete":
        delete_mask |= mismatch_row_mask.reindex(
            cleaned.index, fill_value=False
        )

    cleaned = cleaned.loc[~delete_mask].copy()

    missing_highlight = set()
    mismatch_highlight = set()

    if options.get("missing_rows", "highlight") == "highlight":
        missing_highlight = set(
            missing_row_mask.reindex(cleaned.index, fill_value=False)
            .loc[lambda series: series]
            .index
        )

    if options.get("mismatch_rows", "highlight") == "highlight":
        mismatch_highlight = set(
            mismatch_row_mask.reindex(cleaned.index, fill_value=False)
            .loc[lambda series: series]
            .index
        )

    return cleaned, missing_highlight, mismatch_highlight


def json_response_download(payload, filename):
    response = HttpResponse(
        json.dumps(
            payload,
            ensure_ascii=False,
            indent=2,
            allow_nan=False,
            default=str,
        ),
        content_type="application/json; charset=utf-8",
    )

    response["Content-Disposition"] = (
        f'attachment; filename="{filename}"'
    )

    return response


def cleaned_file_response(
    df,
    filename,
    extension,
    highlight_rows=None,
    missing_cells_mode="context",
):
    highlight_rows = highlight_rows or set()
    base = os.path.splitext(os.path.basename(filename))[0] or "data"

    if extension == ".csv" and not highlight_rows:
        response = HttpResponse(
            df.to_csv(index=False),
            content_type="text/csv; charset=utf-8",
        )
        response["Content-Disposition"] = (
            f'attachment; filename="{base}_cleaned.csv"'
        )
        return response

    if extension == ".json" and not highlight_rows:
        return json_response_download(
            dataframe_to_records(df),
            f"{base}_cleaned.json",
        )

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        worksheet = writer.sheets["Cleaned Data"]

        if highlight_rows:
            from openpyxl.styles import PatternFill

            red_fill = PatternFill(
                fill_type="solid",
                fgColor="FFFFC7CE",
            )

            for row_number, original_index in enumerate(df.index, start=2):
                if original_index in highlight_rows:
                    for cell in worksheet[row_number]:
                        cell.fill = red_fill

    buffer.seek(0)

    response = HttpResponse(
        buffer.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    response["Content-Disposition"] = (
        f'attachment; filename="{base}_cleaned.xlsx"'
    )

    return response


def build_context(
    request,
    df,
    selected_sheet,
    row_fields=None,
    column_fields=None,
    value_fields=None,
    filter_fields=None,
    selected_filters=None,
    aggregation="count",
    chart_type="bar",
    chart_title="",
):
    row_fields = row_fields or []
    column_fields = column_fields or []
    value_fields = value_fields or []
    filter_fields = filter_fields or []
    selected_filters = selected_filters or {}

    filtered_df = apply_filters(df, filter_fields, selected_filters)

    pivot_df = perform_pivot(
        filtered_df,
        row_fields,
        column_fields,
        value_fields,
        aggregation,
    )

    pivot_chart = ""
    pivot_insights = []

    if not pivot_df.empty:
        chart_df = create_chart_dataframe(pivot_df, row_fields)
        figure = create_chart_figure(
            chart_df, chart_type, chart_title
        )
        pivot_chart = create_chart_image(figure)

        pivot_insights = generate_pivot_insights(
            pivot_df,
            row_fields,
            column_fields,
            value_fields,
            aggregation,
        )

    extension = get_file_extension(request)
    sheets = get_sheet_names(request)

    return {
        "uploaded": True,
        "file_name": request.session.get(SESSION_NAME_KEY, ""),
        "file_type": extension.lstrip(".").upper(),
        "headers": list(df.columns),
        "table_data": dataframe_to_records(df.head(100)),
        "sheets": sheets,
        "selected_sheet": selected_sheet or "",
        "field_types": get_field_types(df),
        "dataset_insights": generate_dataset_insights(df),
        "row_fields": row_fields,
        "column_fields": column_fields,
        "value_fields": value_fields,
        "filter_fields": filter_fields,
        "selected_filters": selected_filters,
        "filter_control_data": get_filter_control_data(
            df, filter_fields
        ),
        "aggregation": aggregation,
        "chart_type": chart_type,
        "chart_title": chart_title,
        "pivot_table": dataframe_to_records(pivot_df),
        "pivot_headers": list(pivot_df.columns),
        "pivot_chart": pivot_chart,
        "pivot_insights": pivot_insights,
        "financial_ratios": [],
        "financial_analysis": [],
        "cleaner_report": None,
        "cleaner_preview": [],
        "cleaner_headers": [],
        "error": "",
    }


def base_context():
    return {
        "uploaded": False,
        "headers": [],
        "table_data": [],
        "sheets": [],
        "selected_sheet": "",
        "field_types": {},
        "dataset_insights": [],
        "row_fields": [],
        "column_fields": [],
        "value_fields": [],
        "filter_fields": [],
        "selected_filters": {},
        "filter_control_data": {},
        "aggregation": "count",
        "chart_type": "bar",
        "chart_title": "",
        "pivot_table": [],
        "pivot_headers": [],
        "pivot_chart": "",
        "pivot_insights": [],
        "financial_ratios": [],
        "financial_analysis": [],
        "file_name": "",
        "file_type": "",
        "cleaner_report": None,
        "cleaner_preview": [],
        "cleaner_headers": [],
        "error": "",
    }


def get_request_analysis(request):
    selected_sheet = request.POST.get("sheet_name", "")
    row_fields = parse_list(request, "row_fields")
    column_fields = parse_list(request, "column_fields")
    value_fields = parse_list(request, "value_fields")
    filter_fields = parse_list(request, "filter_fields")
    aggregation = request.POST.get("aggregation", "count")
    chart_type = request.POST.get("chart_type", "bar")
    chart_title = request.POST.get("chart_title", "")
    selected_filters = parse_selected_filters(
        request, filter_fields
    )

    df = load_dataframe(request, selected_sheet)

    return (
        df,
        selected_sheet,
        row_fields,
        column_fields,
        value_fields,
        filter_fields,
        selected_filters,
        aggregation,
        chart_type,
        chart_title,
    )


def excel_response_download(dataframes, filename):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, dataframe in dataframes.items():
            safe_name = str(sheet_name)[:31] or "Sheet"
            dataframe.to_excel(
                writer,
                sheet_name=safe_name,
                index=False,
            )

    buffer.seek(0)

    response = HttpResponse(
        buffer.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    response["Content-Disposition"] = (
        f'attachment; filename="{filename}"'
    )

    return response


def export_analysis_payload(
    request,
    df,
    selected_sheet,
    row_fields,
    column_fields,
    value_fields,
    filter_fields,
    selected_filters,
    aggregation,
    chart_type,
    chart_title,
):
    filtered_df = apply_filters(
        df, filter_fields, selected_filters
    )

    pivot_df = perform_pivot(
        filtered_df,
        row_fields,
        column_fields,
        value_fields,
        aggregation,
    )

    insights = generate_pivot_insights(
        pivot_df,
        row_fields,
        column_fields,
        value_fields,
        aggregation,
    )

    chart_df = create_chart_dataframe(pivot_df, row_fields)

    chart_payload = {
        "title": chart_title or "Pivot Chart",
        "type": chart_type,
        "category_label": (
            str(chart_df.index.name or "Category")
            if not chart_df.empty else ""
        ),
        "series": [],
        "categories": [],
    }

    if not chart_df.empty:
        for column in chart_df.columns:
            chart_payload["series"].append({
                "name": str(column),
                "values": [
                    json_safe_value(value)
                    for value in chart_df[column].tolist()
                ],
            })

        chart_payload["categories"] = [
            json_safe_value(value)
            for value in chart_df.index.tolist()
        ]

    chart_figure = create_chart_figure(
        chart_df, chart_type, chart_title
    )

    chart_payload["image_base64"] = create_chart_image(chart_figure)

    return {
        "source_file": request.session.get(SESSION_NAME_KEY, ""),
        "source_type": get_file_extension(request).lstrip(".").upper(),
        "worksheet": selected_sheet or None,
        "exported_at": datetime.now().astimezone().isoformat(),
        "source_data": dataframe_to_records(df),
        "filtered_data": dataframe_to_records(filtered_df),
        "pivot_configuration": {
            "rows": row_fields,
            "columns": column_fields,
            "values": value_fields,
            "filters": filter_fields,
            "selected_filters": selected_filters,
            "aggregation": aggregation,
        },
        "pivot_table": dataframe_to_records(pivot_df),
        "insights": insights,
        "dataset_insights": generate_dataset_insights(df),
        "chart": chart_payload,
        "financial_ratios": calculate_financial_ratios(df),
    }


@login_required
def index(request):
    context = base_context()

    if request.method == "POST":
        action = request.POST.get("action", "")

        if action == "upload":
            uploaded_file = request.FILES.get("excel_file")

            if uploaded_file is None:
                context["error"] = (
                    "Please select an Excel, CSV, or JSON file."
                )

            elif not uploaded_file.name.lower().endswith(
                SUPPORTED_EXTENSIONS
            ):
                context["error"] = (
                    "Unsupported file type. Please upload "
                    ".xlsx, .xls, .csv, or .json."
                )

            else:
                try:
                    save_dataframe_to_session(request, uploaded_file)

                    sheets = get_sheet_names(request)
                    selected_sheet = sheets[0] if sheets else ""

                    df = load_dataframe(request, selected_sheet)

                    if df.empty:
                        raise ValueError(
                            "The uploaded file is empty or contains "
                            "no readable records."
                        )

                    context = build_context(
                        request, df, selected_sheet
                    )

                except Exception as exc:
                    context["error"] = (
                        f"Could not read the uploaded file: {exc}"
                    )

        elif action == "scan_cleaner":
            if not request.session.get(SESSION_DATA_KEY):
                context["error"] = (
                    "Please upload a file before scanning it."
                )

            else:
                try:
                    selected_sheet = request.POST.get(
                        "sheet_name", ""
                    )

                    raw_df = read_raw_dataframe(
                        request, selected_sheet
                    )

                    extension = get_file_extension(request)

                    if extension == ".json":
                        df = raw_df.copy()
                        header_row = 0
                    else:
                        header_row = detect_header_row(raw_df)
                        df = dataframe_from_raw(
                            raw_df, header_row
                        )

                    if df.empty:
                        raise ValueError(
                            "The selected dataset has no readable records."
                        )

                    report = scan_data_quality(df, header_row)

                    request.session["cleaner_report"] = report
                    request.session["cleaner_header_row"] = header_row
                    request.session["cleaner_sheet"] = selected_sheet

                    source_df = load_dataframe(
                        request, selected_sheet
                    )

                    context = build_context(
                        request, source_df, selected_sheet
                    )

                    context["cleaner_report"] = report
                    context["cleaner_headers"] = list(df.columns)
                    context["cleaner_preview"] = dataframe_to_records(
                        df.head(20)
                    )

                except Exception as exc:
                    context["error"] = (
                        f"Could not scan data quality: {exc}"
                    )

        elif action == "apply_cleaner":
            if not request.session.get(SESSION_DATA_KEY):
                context["error"] = (
                    "Please upload a file before cleaning it."
                )

            elif request.POST.get("consent_cleaning") != "yes":
                context["error"] = (
                    "Please confirm that you approve the selected "
                    "cleaning operations. No changes were made."
                )

                try:
                    selected_sheet = request.session.get(
                        "cleaner_sheet", ""
                    )

                    df = load_dataframe(request, selected_sheet)

                    context = build_context(
                        request, df, selected_sheet
                    )

                    context["error"] = (
                        "Please confirm that you approve the selected "
                        "cleaning operations. No changes were made."
                    )

                    context["cleaner_report"] = request.session.get(
                        "cleaner_report"
                    )

                except Exception:
                    pass

            else:
                try:
                    selected_sheet = request.session.get(
                        "cleaner_sheet", ""
                    )

                    extension = get_file_extension(request)

                    raw_df = read_raw_dataframe(
                        request, selected_sheet
                    )

                    header_row = int(
                        request.session.get(
                            "cleaner_header_row", 0
                        )
                    )

                    if extension == ".json":
                        source_df = raw_df.copy()
                    else:
                        selected_header = (
                            header_row
                            if request.POST.get("fix_headers") == "yes"
                            else 0
                        )

                        source_df = dataframe_from_raw(
                            raw_df, selected_header
                        )

                    if source_df.empty:
                        raise ValueError(
                            "The dataset is empty after reading "
                            "the selected header row."
                        )

                    report = scan_data_quality(
                        source_df, header_row
                    )

                    options = {
                        "fill_missing": request.POST.get(
                            "fill_missing", "context"
                        ),
                        "remove_empty_rows": request.POST.get(
                            "remove_empty_rows", "yes"
                        ),
                        "missing_rows": request.POST.get(
                            "missing_rows", "highlight"
                        ),
                        "mismatch_rows": request.POST.get(
                            "mismatch_rows", "highlight"
                        ),
                    }

                    cleaned, missing_highlight, mismatch_highlight = (
                        clean_dataframe(
                            source_df, report, options
                        )
                    )

                    highlight_rows = set()

                    if options["missing_rows"] == "highlight":
                        highlight_rows.update(missing_highlight)

                    if options["mismatch_rows"] == "highlight":
                        highlight_rows.update(mismatch_highlight)

                    return cleaned_file_response(
                        cleaned,
                        request.session.get(
                            SESSION_NAME_KEY, "data"
                        ),
                        extension,
                        highlight_rows,
                    )

                except Exception as exc:
                    context["error"] = (
                        f"Could not clean the dataset: {exc}"
                    )

        elif action in (
            "change_sheet",
            "build_pivot",
            "download_chart_pdf",
            "financial_analysis",
            "export_json",
            "export_excel",
        ):
            if not request.session.get(SESSION_DATA_KEY):
                context["error"] = (
                    "Please upload a file before using this feature."
                )

            else:
                try:
                    values = get_request_analysis(request)

                    (
                        df,
                        selected_sheet,
                        row_fields,
                        column_fields,
                        value_fields,
                        filter_fields,
                        selected_filters,
                        aggregation,
                        chart_type,
                        chart_title,
                    ) = values

                    if df.empty:
                        raise ValueError(
                            "The selected dataset is empty."
                        )

                    if action == "export_json":
                        payload = export_analysis_payload(
                            request,
                            df,
                            selected_sheet,
                            row_fields,
                            column_fields,
                            value_fields,
                            filter_fields,
                            selected_filters,
                            aggregation,
                            chart_type,
                            chart_title,
                        )

                        return json_response_download(
                            payload, "analysis_results.json"
                        )

                    if action == "export_excel":
                        filtered_df = apply_filters(
                            df, filter_fields, selected_filters
                        )

                        pivot_df = perform_pivot(
                            filtered_df,
                            row_fields,
                            column_fields,
                            value_fields,
                            aggregation,
                        )

                        ratios = calculate_financial_ratios(df)

                        pivot_insights = generate_pivot_insights(
                            pivot_df,
                            row_fields,
                            column_fields,
                            value_fields,
                            aggregation,
                        )

                        return excel_response_download({
                            "Source Data": df,
                            "Filtered Data": filtered_df,
                            "Pivot Table": pivot_df,
                            "Dataset Insights": pd.DataFrame({
                                "Insight": generate_dataset_insights(df)
                            }),
                            "Pivot Insights": pd.DataFrame({
                                "Insight": pivot_insights
                            }),
                            "Financial Ratios": pd.DataFrame(ratios),
                        }, "analysis_results.xlsx")

                    if action == "download_chart_pdf":
                        filtered_df = apply_filters(
                            df, filter_fields, selected_filters
                        )

                        pivot_df = perform_pivot(
                            filtered_df,
                            row_fields,
                            column_fields,
                            value_fields,
                            aggregation,
                        )

                        chart_df = create_chart_dataframe(
                            pivot_df, row_fields
                        )

                        figure = create_chart_figure(
                            chart_df,
                            chart_type,
                            chart_title or "Pivot Chart",
                        )

                        if figure is None:
                            raise ValueError(
                                "There is not enough numeric data "
                                "to create the chart."
                            )

                        response_buffer = io.BytesIO()

                        figure.savefig(
                            response_buffer,
                            format="pdf",
                            bbox_inches="tight",
                        )

                        plt.close(figure)
                        response_buffer.seek(0)

                        response = HttpResponse(
                            response_buffer.read(),
                            content_type="application/pdf",
                        )

                        response["Content-Disposition"] = (
                            'attachment; filename="pivot_chart.pdf"'
                        )

                        return response

                    if action == "financial_analysis":
                        context = build_context(
                            request, df, selected_sheet
                        )

                        ratios = calculate_financial_ratios(df)

                        context["financial_ratios"] = ratios
                        context["financial_analysis"] = ratios

                        if not ratios:
                            context["error"] = (
                                "No financial ratios could be calculated. "
                                "Make sure the dataset contains financial "
                                "account names and numerical amounts."
                            )

                    elif action == "change_sheet":
                        context = build_context(
                            request, df, selected_sheet
                        )

                    elif action == "build_pivot":
                        if not value_fields:
                            context = build_context(
                                request, df, selected_sheet
                            )

                            context["error"] = (
                                "Add at least one field to the Values "
                                "area before building the PivotTable."
                            )

                        else:
                            context = build_context(
                                request,
                                df,
                                selected_sheet,
                                row_fields,
                                column_fields,
                                value_fields,
                                filter_fields,
                                selected_filters,
                                aggregation,
                                chart_type,
                                chart_title,
                            )

                except Exception as exc:
                    context["error"] = str(exc)

    return render(request, "viewer/index.html", context)
