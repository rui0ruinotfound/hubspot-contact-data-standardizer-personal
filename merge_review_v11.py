#!/usr/bin/env python3
"""
Merge manual review decisions into final HubSpot upload files.

Run:
  python3 merge_review_v11.py
"""

from pathlib import Path
import argparse
import pandas as pd

try:
    import openpyxl  # noqa
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


def read_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    if path.suffix.lower() in [".xlsx", ".xls"]:
        return pd.read_excel(path, dtype=str, keep_default_na=False)
    raise ValueError(f"Unsupported file type: {path}")


def write_table(df: pd.DataFrame, csv_path: Path, xlsx_path: Path) -> None:
    df.to_csv(csv_path, index=False)
    if HAS_OPENPYXL:
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Final")


def norm(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def find_col(df, candidates):
    lower = {c.lower().strip(): c for c in df.columns}
    for c in candidates:
        if c.lower().strip() in lower:
            return lower[c.lower().strip()]
    return None


def apply_review_to_file(cleaned_df: pd.DataFrame, review_df: pd.DataFrame, file_label: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    record_col = find_col(cleaned_df, ["Record ID", "RecordID", "Contact ID", "Object ID", "ID", "record_id"])
    country_col = find_col(cleaned_df, ["Country/Region", "Country", "Country Region", "Country or Region"])
    state_col = find_col(cleaned_df, ["State/Region", "State", "State Region", "Region", "Province"])

    if record_col is None:
        raise ValueError(f"Could not find Record ID column in {file_label} cleaned file.")

    cleaned = cleaned_df.copy()
    cleaned[record_col] = cleaned[record_col].astype(str)
    idx_by_record = {str(v): i for i, v in cleaned[record_col].items()}
    log_rows = []

    for _, row in review_df.iterrows():
        record_id = str(row.get("record_id", "")).strip()
        field = norm(row.get("field", ""))
        action = norm(row.get("action", "")).lower()
        final_country = norm(row.get("final_country", ""))
        final_state = norm(row.get("final_state_region", ""))

        if not action:
            action = "update" if final_country or final_state else "keep"

        if record_id not in idx_by_record:
            log_rows.append({"record_id": record_id, "status": "not_found", "detail": "Record ID not found in cleaned file."})
            continue

        target_idx = idx_by_record[record_id]

        if action == "skip":
            log_rows.append({"record_id": record_id, "status": "skipped", "detail": "Action=skip."})
            continue
        if action == "keep":
            log_rows.append({"record_id": record_id, "status": "kept", "detail": "No change."})
            continue
        if action not in ["update", "blank"]:
            log_rows.append({"record_id": record_id, "status": "invalid_action", "detail": f"Unknown action: {action}"})
            continue

        changes = []

        if country_col and (final_country or (action == "blank" and "country" in field.lower())):
            old = cleaned.at[target_idx, country_col]
            cleaned.at[target_idx, country_col] = "" if action == "blank" else final_country
            changes.append(f"{country_col}: {old} -> {cleaned.at[target_idx, country_col]}")

        if state_col and (final_state or (action == "blank" and ("state" in field.lower() or "region" in field.lower()))):
            old = cleaned.at[target_idx, state_col]
            cleaned.at[target_idx, state_col] = "" if action == "blank" else final_state
            changes.append(f"{state_col}: {old} -> {cleaned.at[target_idx, state_col]}")

        if changes:
            for col in cleaned.columns:
                if "Exception" in col or col == "Review Required":
                    cleaned.at[target_idx, col] = ""

        if not changes:
            log_rows.append({"record_id": record_id, "status": "no_change", "detail": "No final value was provided."})
        else:
            log_rows.append({"record_id": record_id, "status": "updated", "detail": "; ".join(changes)})

    return cleaned, pd.DataFrame(log_rows)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--review-file", default=None)
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    review_file = Path(args.review_file) if args.review_file else output_dir / "review_required_v11.csv"
    review = read_table(review_file)

    cleaned_files = sorted(output_dir.glob("cleaned_*_v11.csv"))
    if not cleaned_files:
        raise FileNotFoundError("No cleaned_*_v11.csv files found in output directory.")

    logs = []

    for cleaned_path in cleaned_files:
        cleaned = read_table(cleaned_path)
        final_df, log = apply_review_to_file(cleaned, review, cleaned_path.stem)
        final_name = cleaned_path.name.replace("cleaned_", "final_")
        final_csv = output_dir / final_name
        final_xlsx = output_dir / final_name.replace(".csv", ".xlsx")
        write_table(final_df, final_csv, final_xlsx)
        log["cleaned_file"] = cleaned_path.name
        logs.append(log)

    merge_log = pd.concat(logs, ignore_index=True) if logs else pd.DataFrame()
    merge_log.to_csv(output_dir / "merge_log_v11.csv", index=False)
    if HAS_OPENPYXL:
        with pd.ExcelWriter(output_dir / "merge_log_v11.xlsx", engine="openpyxl") as writer:
            merge_log.to_excel(writer, index=False, sheet_name="Merge Log")

    print("==================================")
    print("Manual Review Merge v11")
    print("==================================")
    print(f"Review file: {review_file}")
    print(f"Rows reviewed: {len(review)}")
    print("Generated final_*_v11.csv/xlsx files in output/")
    print("==================================")


if __name__ == "__main__":
    main()
