from __future__ import annotations

from pathlib import Path
import re
from typing import Any

import numpy as np
import pandas as pd

EXPECTED_COLUMNS = [
    "Opportunity_ID",
    "Account_ID",
    "Account_Name",
    "Opportunity_Owner",
    "Opportunity_Name",
    "Type",
    "Stage",
    "Amount",
    "Amount_Currency",
    "Total_Product_Amount",
    "Total_Product_Amount_Currency",
    "Product_Name",
    "Close_Date",
    "Created_Date",
    "Lead_Source",
    "Lead_Office",
    "Stage_Duration",
    "Delivery_Length_Months",
]

DATE_COLUMNS = ["Close_Date", "Created_Date"]
NUMERIC_COLUMNS = ["Amount", "Total_Product_Amount", "Stage_Duration", "Delivery_Length_Months"]
CURRENCY_COLUMNS = ["Amount_Currency", "Total_Product_Amount_Currency"]
ID_COLUMN = "Opportunity_ID"
UNKNOWN = "Unknown"


def is_blank(value: Any) -> bool:
    if pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def load_opportunities(path: Path) -> pd.DataFrame:
    workbook = pd.ExcelFile(path)
    lower_name_map = {name.lower(): name for name in workbook.sheet_names}
    sheet_name = lower_name_map.get("opportunities")

    if sheet_name:
        return pd.read_excel(path, sheet_name=sheet_name)
    return pd.read_excel(path)


def ensure_expected_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in EXPECTED_COLUMNS:
        if col not in out.columns:
            out[col] = np.nan
    return out[EXPECTED_COLUMNS]


def normalize_id_column(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out[ID_COLUMN] = out[ID_COLUMN].apply(
        lambda x: np.nan if is_blank(x) else str(x).strip()
    )
    return out


def parse_date_series(series: pd.Series) -> tuple[pd.Series, int]:
    original_not_blank = series.apply(lambda x: not is_blank(x))
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    numeric_values = pd.to_numeric(series, errors="coerce")
    plausible_serial = numeric_values.between(20000, 80000, inclusive="both")
    if plausible_serial.any():
        parsed.loc[plausible_serial] = pd.to_datetime(
            numeric_values.loc[plausible_serial],
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )

    unresolved = parsed.isna() & original_not_blank
    if unresolved.any():
        text = series.loc[unresolved].apply(lambda x: str(x).strip())
        first = pd.to_datetime(text, errors="coerce", dayfirst=False)
        second_mask = first.isna() & text.notna()
        if second_mask.any():
            first.loc[second_mask] = pd.to_datetime(
                text.loc[second_mask],
                errors="coerce",
                dayfirst=True,
            )
        parsed.loc[unresolved] = first

    failures = int((parsed.isna() & original_not_blank).sum())
    return parsed, failures


def normalize_numeric_text(raw: str) -> str:
    s = raw.strip()
    s = s.replace("\u00A0", " ").replace(" ", "")

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]

    s = re.sub(r"[^0-9,.-]", "", s)
    if not s:
        return ""

    if s.count("-") > 1:
        s = s.replace("-", "")
    if "-" in s and not s.startswith("-"):
        s = s.replace("-", "")

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2:
            if len(parts[1]) == 3 and len(parts[0]) >= 1:
                s = "".join(parts)
            else:
                s = parts[0] + "." + parts[1]
        else:
            if all(len(chunk) == 3 for chunk in parts[1:]):
                s = "".join(parts)
            else:
                s = "".join(parts[:-1]) + "." + parts[-1]
    elif "." in s:
        parts = s.split(".")
        if len(parts) == 2:
            if len(parts[1]) == 3 and len(parts[0]) >= 1:
                s = "".join(parts)
        else:
            if all(len(chunk) == 3 for chunk in parts[1:]):
                s = "".join(parts)
            else:
                s = "".join(parts[:-1]) + "." + parts[-1]

    if negative and not s.startswith("-"):
        s = "-" + s
    return s


def parse_numeric_value(value: Any) -> float:
    if is_blank(value):
        return np.nan

    if isinstance(value, (int, float, np.number)) and not isinstance(value, bool):
        return float(value)

    text = normalize_numeric_text(str(value))
    if text in {"", "-", ".", "-."}:
        return np.nan

    try:
        return float(text)
    except ValueError:
        return np.nan


def parse_numeric_series(series: pd.Series) -> tuple[pd.Series, int]:
    original_not_blank = series.apply(lambda x: not is_blank(x))
    parsed = series.apply(parse_numeric_value)
    failures = int((parsed.isna() & original_not_blank).sum())
    return parsed, failures


def missing_mask(series: pd.Series) -> pd.Series:
    return series.apply(is_blank)


def row_completeness_score(df: pd.DataFrame) -> pd.Series:
    score = pd.Series(0, index=df.index, dtype="int64")
    for col in df.columns:
        score += (~missing_mask(df[col])).astype("int64")
    return score


def resolve_duplicate_groups(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    if df.empty:
        return df, 0

    selected_rows: list[pd.Series] = []
    duplicates_removed = 0

    for _, group in df.groupby(ID_COLUMN, sort=False, dropna=False):
        if len(group) == 1:
            selected_rows.append(group.iloc[0])
            continue

        unique_rows = group.drop_duplicates()
        duplicates_removed += len(group) - 1

        if len(unique_rows) == 1:
            selected_rows.append(unique_rows.iloc[0])
            continue

        ranked = group.copy()
        ranked["__score"] = row_completeness_score(ranked)
        ranked["__created_rank"] = ranked["Created_Date"].fillna(pd.Timestamp.max)
        ranked = ranked.sort_values(
            by=["__score", "__created_rank"],
            ascending=[False, True],
            kind="stable",
        )
        selected_rows.append(ranked.iloc[0].drop(labels=["__score", "__created_rank"]))

    cleaned = pd.DataFrame(selected_rows).reset_index(drop=True)
    cleaned = cleaned[EXPECTED_COLUMNS]
    return cleaned, duplicates_removed


def normalize_currency_value(value: Any) -> str | None:
    if is_blank(value):
        return None

    raw = str(value).strip().upper()
    cleaned = re.sub(r"[^A-Z]", "", raw)
    if len(cleaned) == 3:
        return cleaned
    return None


def harmonize_currencies(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    out = df.copy()
    corrected = 0

    amount_curr = out["Amount_Currency"].copy()
    product_curr = out["Total_Product_Amount_Currency"].copy()

    final_currency = []
    for idx in out.index:
        ac = normalize_currency_value(amount_curr.loc[idx])
        pc = normalize_currency_value(product_curr.loc[idx])

        if ac:
            chosen = ac
        elif pc:
            chosen = pc
        else:
            chosen = UNKNOWN

        old_ac = UNKNOWN if is_blank(amount_curr.loc[idx]) else str(amount_curr.loc[idx]).strip()
        old_pc = UNKNOWN if is_blank(product_curr.loc[idx]) else str(product_curr.loc[idx]).strip()

        if old_ac != chosen or old_pc != chosen:
            corrected += 1

        final_currency.append(chosen)

    out["Amount_Currency"] = final_currency
    out["Total_Product_Amount_Currency"] = final_currency
    return out, corrected


def fill_missing_values(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    categorical_cols = [
        col for col in EXPECTED_COLUMNS if col not in NUMERIC_COLUMNS and col not in DATE_COLUMNS
    ]

    for col in categorical_cols:
        out[col] = out[col].apply(lambda x: UNKNOWN if is_blank(x) else str(x).strip())

    for col in NUMERIC_COLUMNS:
        out[col] = out[col].fillna(0.0)

    return out


def clean_opportunities(input_path: Path, output_path: Path) -> None:
    df = load_opportunities(input_path)
    df = ensure_expected_columns(df)
    initial_rows = len(df)

    df = normalize_id_column(df)
    null_id_mask = df[ID_COLUMN].isna()
    removed_null_id = int(null_id_mask.sum())
    df = df.loc[~null_id_mask].copy()

    date_failures_total = 0
    for col in DATE_COLUMNS:
        df[col], failures = parse_date_series(df[col])
        date_failures_total += failures

    numeric_failures_total = 0
    for col in NUMERIC_COLUMNS:
        df[col], failures = parse_numeric_series(df[col])
        numeric_failures_total += failures

    df, duplicates_removed = resolve_duplicate_groups(df)
    df, currency_conflicts_corrected = harmonize_currencies(df)
    df = fill_missing_values(df)

    final_rows = len(df)

    df.to_excel(output_path, index=False)

    print(f"Linhas iniciais: {initial_rows}")
    print(f"Linhas removidas por Opportunity_ID nulo/blank: {removed_null_id}")
    print(f"Duplicatas removidas: {duplicates_removed}")
    print(f"Falhas de parse de datas: {date_failures_total}")
    print(f"Falhas de parse de números: {numeric_failures_total}")
    print(f"Conflitos cambiais corrigidos: {currency_conflicts_corrected}")
    print(f"Linhas finais: {final_rows}")


def main() -> None:
    base_dir = Path(__file__).resolve().parent.parent
    input_path = base_dir / "opps_corrupted.xlsx"
    output_path = base_dir / "opps_corrigido.xlsx"

    clean_opportunities(input_path, output_path)


if __name__ == "__main__":
    main()
