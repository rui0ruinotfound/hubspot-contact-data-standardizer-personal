#!/usr/bin/env python3
"""
Contact Data Standardizer v11 Dashboard Edition

Enterprise local-only HubSpot CRM cleanup tool.

v11 focuses on reporting and governance:
- Preserves all original columns.
- Only updates Country/Region and State/Region.
- Reads HubSpot Country Key and State Key from config/.
- Keeps alias config separate from HubSpot key files.
- Adds exception columns at the end of cleaned outputs.
- Generates:
    summary_v11.csv
    review_required_v11.csv
    qa_report_v11.xlsx
    mapping_report_v11.xlsx
    data_quality_dashboard_v11.xlsx
- Supports manual review + merge workflow.

Run:
  python3 standardize_contact_data_v11.py

After manual review:
  python3 merge_review_v11.py
"""

from __future__ import annotations

import argparse
import json
import re
from dataclasses import dataclass
from difflib import SequenceMatcher, get_close_matches
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

try:
    import openpyxl  # noqa
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


DEFAULT_COLUMNS = {
    "record_id": ["Record ID", "RecordID", "Contact ID", "Object ID", "ID", "record_id"],
    "country": ["Country/Region", "Country", "Country Region", "Country or Region", "country"],
    "state": ["State/Region", "State", "State Region", "Region", "Province", "state"],
}

DEFAULT_RULES = {
    "enable_fuzzy_matching": True,
    "fuzzy_auto_fix_threshold": 0.94,
    "fuzzy_suggest_threshold": 0.84,
    "preserve_all_columns": True,
    "strict_country_key_validation": True,
    "strict_state_key_validation": False
}

NULL_LIKE_VALUES = {
    "", "nan", "none", "null", "n/a", "na", "select country", "select state",
    "select region", "unknown", "not provided", "not set", "-", "--"
}

GARBAGE_VALUES = {"##", "ncnc", "ai", "txt", "tv", "f", "sate"}

CANADA_PROVINCES = {
    "AB": "Alberta", "BC": "British Columbia", "MB": "Manitoba", "NB": "New Brunswick",
    "NL": "Newfoundland and Labrador", "NS": "Nova Scotia", "NT": "Northwest Territories",
    "NU": "Nunavut", "ON": "Ontario", "PE": "Prince Edward Island", "QC": "Quebec",
    "SK": "Saskatchewan", "YT": "Yukon",
}

AUSTRALIA_STATES = {
    "ACT": "Australian Capital Territory", "NSW": "New South Wales", "NT": "Northern Territory",
    "QLD": "Queensland", "SA": "South Australia", "TAS": "Tasmania", "VIC": "Victoria",
    "WA": "Western Australia",
}

US_SPECIAL = {
    "PR": ("Puerto Rico - U.S. Territory", "Puerto Rico - U.S. Territory"),
    "VI": ("Virgin Islands", "Virgin Islands"),
    "GU": ("Guam - U.S. Territory", "Guam - U.S. Territory"),
    "AA": ("Armed Forces Americas", "United States"),
    "AE": ("Armed Forces Europe", "United States"),
    "AP": ("Armed Forces Pacific", "United States"),
}


@dataclass
class CleanResult:
    country: str
    state: str
    country_status: str
    state_status: str
    country_exception: str
    state_exception: str
    country_issue: str
    state_issue: str
    review_required: str
    events: List[dict]


def normalize(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip().replace("\u00a0", " "))


def key(value) -> str:
    return normalize(value).lower()


def is_null_like(value) -> bool:
    return key(value) in NULL_LIKE_VALUES


def is_garbage(value) -> bool:
    raw = normalize(value)
    return key(raw) in GARBAGE_VALUES or bool(re.fullmatch(r"[#@$%^&*]+", raw))


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


def find_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    lower = {c.lower().strip(): c for c in df.columns}
    for c in candidates:
        if c.lower().strip() in lower:
            return lower[c.lower().strip()]
    return None


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
            df.to_excel(writer, index=False, sheet_name="Cleaned")


def extract_country_key_values(country_key_path: Path) -> set[str]:
    df = pd.read_csv(country_key_path, dtype=str, keep_default_na=False)
    values = set()
    for col in df.columns:
        for value in df[col].tolist():
            v = normalize(value)
            if v:
                values.add(v)
    return values


def extract_state_key_values(state_key_path: Path) -> tuple[set[str], dict[str, str], dict[str, str]]:
    raw = pd.read_csv(state_key_path, dtype=str, keep_default_na=False, header=None)
    allowed = set()
    code_to_name = {}
    name_to_name = {}

    for _, row in raw.iterrows():
        if len(row) < 1:
            continue
        name = normalize(row.iloc[0])
        code = normalize(row.iloc[1]) if len(row) > 1 else ""

        if key(name) in {"state", "state name", "name"}:
            continue

        if name:
            allowed.add(name)
            name_to_name[key(name)] = name
        if code:
            allowed.add(code)
            code_to_name[key(code)] = name

    return allowed, code_to_name, name_to_name


def load_alias_mapping(path: Path) -> Dict[str, str]:
    if not path.exists():
        return {}
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    raw_col = "raw" if "raw" in df.columns else df.columns[0]
    std_col = "standardized" if "standardized" in df.columns else df.columns[1]
    return {
        key(row[raw_col]): normalize(row[std_col])
        for _, row in df.iterrows()
        if key(row[raw_col]) and normalize(row[std_col])
    }


def load_region_aliases(path: Path) -> List[Tuple[str, str, str]]:
    if not path.exists():
        return []
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    result = []
    for _, row in df.iterrows():
        raw = normalize(row.get("raw", ""))
        state = normalize(row.get("standardized_state", ""))
        country = normalize(row.get("country", ""))
        if raw and state:
            result.append((raw, state, country))
    return result


def fuzzy_match(value: str, choices: list[str], auto_threshold: float, suggest_threshold: float) -> Tuple[str, float, str]:
    raw = key(value)
    if not raw or len(raw) <= 2 or not choices:
        return "", 0.0, "none"
    candidates = get_close_matches(raw, choices, n=1, cutoff=suggest_threshold)
    if not candidates:
        return "", 0.0, "none"
    candidate = candidates[0]
    score = SequenceMatcher(None, raw, candidate).ratio()
    if score >= auto_threshold:
        return candidate, score, "auto"
    if score >= suggest_threshold:
        return candidate, score, "suggest"
    return "", 0.0, "none"


def build_country_mapping(config_dir: Path, allowed_countries: set[str]) -> Dict[str, str]:
    mapping = {key(v): v for v in allowed_countries}

    built_in_aliases = {
        "us": "United States",
        "u.s.": "United States",
        "u.s.a.": "United States",
        "usa": "United States",
        "united states of america": "United States",
        "america": "United States",
        "uk": "United Kingdom",
        "u.k.": "United Kingdom",
        "gb": "United Kingdom",
        "gbr": "United Kingdom",
        "great britain": "United Kingdom",
        "england": "United Kingdom",
        "can": "Canada",
        "au": "Australia",
        "aus": "Australia",
        "mx": "Mexico",
        "it": "Italy",
        "nz": "New Zealand",
        "ng": "Nigeria",
        "in": "India",
        "sg": "Singapore",
        "hu": "Hungary",
        "cn": "China",
        "de": "Germany",
        "be": "Belgium",
        "br": "Brazil",
        "fr": "France",
        "es": "Spain",
        "jp": "Japan",
        "nl": "Netherlands",
        "the netherlands": "Netherlands",
        "hk": "Hong Kong",
        "columbia": "Colombia",
        "trinidad": "Trinidad and Tobago",
        "trinidad & tobago": "Trinidad and Tobago",
        "tto": "Trinidad and Tobago",
        "pn": "Pitcairn Island - British Territory",
        "pr": "Puerto Rico - U.S. Territory",
        "puerto rico": "Puerto Rico - U.S. Territory",
        "guam": "Guam - U.S. Territory",
        "gu": "Guam - U.S. Territory",
        "curacao": "Curacao",
        "curaçao": "Curacao",
        "singpore": "Singapore",
    }
    mapping.update({key(k): v for k, v in built_in_aliases.items()})
    mapping.update(load_alias_mapping(config_dir / "country_aliases.csv"))

    return mapping


def build_state_lookup(config_dir: Path, state_code_to_name: dict[str, str], state_name_to_name: dict[str, str]) -> Dict[str, List[Tuple[str, str]]]:
    lookup: Dict[str, List[Tuple[str, str]]] = {}

    def add(alias: str, state: str, country: str):
        k = key(alias)
        pair = (state, country)
        lookup.setdefault(k, [])
        if pair not in lookup[k]:
            lookup[k].append(pair)

    for code_key, state_name in state_code_to_name.items():
        add(code_key, state_name, "United States")
    for name_key, state_name in state_name_to_name.items():
        add(name_key, state_name, "United States")

    for abbr, state in CANADA_PROVINCES.items():
        add(abbr, state, "Canada")
        add(state, state, "Canada")
    for abbr, state in AUSTRALIA_STATES.items():
        add(abbr, state, "Australia")
        add(state, state, "Australia")
    for abbr, (state, country) in US_SPECIAL.items():
        add(abbr, state, country)
        add(state, state, country)

    for raw, state, country in load_region_aliases(config_dir / "state_region_aliases.csv"):
        add(raw, state, country)

    return lookup


def standardize_country(value, country_mapping: Dict[str, str], rules: dict) -> Tuple[str, str, str, float]:
    if is_null_like(value):
        return "", "blank", "", 0.0
    if is_garbage(value):
        return "", "garbage_blank", "", 0.0

    raw = normalize(value)
    k = key(raw)

    if k in country_mapping:
        return country_mapping[k], "standardized", "", 0.0

    if re.fullmatch(r"\d+", raw):
        return "", "unknown_numeric_blank", "", 0.0

    if rules.get("enable_fuzzy_matching", True):
        candidate_key, score, mode = fuzzy_match(
            raw,
            list(country_mapping.keys()),
            float(rules.get("fuzzy_auto_fix_threshold", 0.94)),
            float(rules.get("fuzzy_suggest_threshold", 0.84)),
        )
        if mode == "auto":
            return country_mapping[candidate_key], "fuzzy_country_auto_fixed", country_mapping[candidate_key], score
        if mode == "suggest":
            return raw, "fuzzy_country_suggestion_review", country_mapping[candidate_key], score

    return raw, "country_not_in_hubspot_key_review", "", 0.0


def standardize_state(value, std_country: str, state_lookup: Dict[str, List[Tuple[str, str]]], rules: dict) -> Tuple[str, str, str, str, float]:
    if is_null_like(value):
        return "", "", "blank", "", 0.0
    if is_garbage(value):
        return "", "", "garbage_blank", "", 0.0

    raw = normalize(value)
    matches = state_lookup.get(key(raw), [])

    if matches:
        if std_country:
            filtered = [m for m in matches if key(m[1]) == key(std_country)]
            if len(filtered) == 1:
                return filtered[0][0], filtered[0][1], "standardized_country_matched", "", 0.0
        unique = list(dict.fromkeys(matches))
        if len(unique) == 1:
            return unique[0][0], unique[0][1], "standardized", "", 0.0
        return raw, "", "ambiguous_review", "", 0.0

    if re.fullmatch(r"\d+", raw):
        return "", "", "unknown_numeric_blank", "", 0.0

    if rules.get("enable_fuzzy_matching", True):
        candidate_key, score, mode = fuzzy_match(
            raw,
            list(state_lookup.keys()),
            float(rules.get("fuzzy_auto_fix_threshold", 0.94)),
            float(rules.get("fuzzy_suggest_threshold", 0.84)),
        )
        if candidate_key:
            candidate_matches = state_lookup[candidate_key]
            if std_country:
                filtered = [m for m in candidate_matches if key(m[1]) == key(std_country)]
                if len(filtered) == 1:
                    if mode == "auto":
                        return filtered[0][0], filtered[0][1], "fuzzy_state_auto_fixed", filtered[0][0], score
                    return raw, "", "fuzzy_state_suggestion_review", filtered[0][0], score
            unique = list(dict.fromkeys(candidate_matches))
            if len(unique) == 1:
                if mode == "auto":
                    return unique[0][0], unique[0][1], "fuzzy_state_auto_fixed", unique[0][0], score
                return raw, "", "fuzzy_state_suggestion_review", unique[0][0], score

    return raw, "", "state_unknown_review", "", 0.0


def parse_combined_location(value: str, country_mapping, state_lookup, rules) -> Tuple[str, str, str]:
    raw = normalize(value)
    if not raw:
        return "", "", "blank"

    parts = [p.strip() for p in re.split(r"[/|;,]", raw) if p.strip()]
    if len(parts) < 2:
        return "", "", "not_combined"

    for country_part, state_part in [(parts[0], parts[1]), (parts[1], parts[0])]:
        country, c_status, _, _ = standardize_country(country_part, country_mapping, rules)
        if country and c_status not in ["country_not_in_hubspot_key_review"]:
            state, _, s_status, _, _ = standardize_state(state_part, country, state_lookup, rules)
            if state or s_status in ["blank", "garbage_blank", "unknown_numeric_blank"]:
                return country, state, "parsed_combined"

    return "", "", "not_combined"


def validate_country(country: str, allowed_countries: set[str]) -> tuple[str, str]:
    if not country:
        return "BLANK", ""
    if country in allowed_countries:
        return "PASS", ""
    return "REVIEW", f"Country not found in HubSpot Country Key: {country}"


def add_event(events, category, file_name, idx, record_id, field, original, standardized, issue, recommendation, final_country="", final_state_region="", action="", suggestion="", confidence=""):
    events.append({
        "category": category,
        "file": file_name,
        "row_number_excel": idx + 2,
        "record_id": record_id,
        "field": field,
        "original_value": original,
        "standardized_value": standardized,
        "issue": issue,
        "recommendation": recommendation,
        "suggestion": suggestion,
        "confidence": confidence,
        "final_country": final_country,
        "final_state_region": final_state_region,
        "action": action,
        "review_notes": "",
    })


class RuleEngine:
    def __init__(self, country_mapping, state_lookup, allowed_countries, rules):
        self.country_mapping = country_mapping
        self.state_lookup = state_lookup
        self.allowed_countries = allowed_countries
        self.rules = rules

    def clean_pair(self, raw_country, raw_state, file_name, idx, record_id, country_col, state_col) -> CleanResult:
        events: List[dict] = []

        country = ""
        state = ""
        country_status = ""
        state_status = ""

        pc1, ps1, status1 = parse_combined_location(raw_country, self.country_mapping, self.state_lookup, self.rules)
        if status1 == "parsed_combined":
            country, state = pc1, ps1
            country_status, state_status = "parsed_combined", "parsed_combined"
            add_event(events, "Auto Fixed", file_name, idx, record_id, country_col, raw_country, country, "parsed_combined_country", "Parsed combined Country/State value.")
            if normalize(raw_state) != state:
                add_event(events, "Auto Fixed", file_name, idx, record_id, state_col, raw_state, state, "parsed_combined_state", "Parsed combined Country/State value.")
        else:
            pc2, ps2, status2 = parse_combined_location(raw_state, self.country_mapping, self.state_lookup, self.rules)
            if status2 == "parsed_combined":
                country, state = pc2, ps2
                country_status, state_status = "parsed_combined", "parsed_combined"
                add_event(events, "Auto Fixed", file_name, idx, record_id, country_col, raw_country, country, "parsed_combined_country_from_state", "Parsed combined value from State/Region.")
                add_event(events, "Auto Fixed", file_name, idx, record_id, state_col, raw_state, state, "parsed_combined_state", "Parsed combined value from State/Region.")
            else:
                country, country_status, country_suggestion, country_conf = standardize_country(raw_country, self.country_mapping, self.rules)
                state, inferred_country, state_status, state_suggestion, state_conf = standardize_state(raw_state, country, self.state_lookup, self.rules)

                if state_status in ["state_unknown_review", "fuzzy_state_suggestion_review"]:
                    state_as_country, state_country_status, _, _ = standardize_country(raw_state, self.country_mapping, self.rules)
                    if state_as_country and state_country_status in ["standardized", "fuzzy_country_auto_fixed"]:
                        if not country:
                            country = state_as_country
                            country_status = "country_moved_from_state"
                            state = ""
                            state_status = "field_corrected_country_in_state"
                            add_event(events, "Auto Fixed", file_name, idx, record_id, country_col, raw_country, country, "country_moved_from_state", "Country value was found in State/Region and moved to Country/Region.")
                            add_event(events, "Auto Fixed", file_name, idx, record_id, state_col, raw_state, state, "field_corrected_country_in_state", "State/Region contained a country and was cleared.")
                        elif key(country) == key(state_as_country):
                            state = ""
                            state_status = "field_corrected_country_in_state"
                            add_event(events, "Auto Fixed", file_name, idx, record_id, state_col, raw_state, state, "field_corrected_country_in_state", "State/Region duplicated Country/Region and was cleared.")

                if not country and inferred_country:
                    country = inferred_country
                    country_status = "blank_country_filled_from_state"
                    add_event(events, "Auto Fixed", file_name, idx, record_id, country_col, raw_country, country, "blank_country_filled_from_state", "Country inferred from State/Region.")

                if country_status in ["standardized", "fuzzy_country_auto_fixed"] and normalize(raw_country) != country and not is_null_like(raw_country):
                    add_event(events, "Auto Fixed", file_name, idx, record_id, country_col, raw_country, country, country_status, "Country standardized automatically.", confidence=round(country_conf, 4) if country_conf else "")
                elif country_status in ["garbage_blank", "unknown_numeric_blank"]:
                    add_event(events, "Auto Fixed", file_name, idx, record_id, country_col, raw_country, country, country_status, "Invalid country value cleared.")
                elif country_status == "fuzzy_country_suggestion_review":
                    add_event(events, "Review Required", file_name, idx, record_id, country_col, raw_country, country, country_status, "Possible fuzzy country match. Confirm or override.", suggestion=country_suggestion, confidence=round(country_conf, 4))
                elif country_status == "country_not_in_hubspot_key_review":
                    add_event(events, "Review Required", file_name, idx, record_id, country_col, raw_country, country, country_status, "Country is not in provided HubSpot Country Key.")

                if state_status in ["standardized", "standardized_country_matched", "fuzzy_state_auto_fixed"] and normalize(raw_state) != state:
                    add_event(events, "Auto Fixed", file_name, idx, record_id, state_col, raw_state, state, state_status, "State/Region standardized automatically.", confidence=round(state_conf, 4) if state_conf else "")
                elif state_status in ["garbage_blank", "unknown_numeric_blank"]:
                    add_event(events, "Auto Fixed", file_name, idx, record_id, state_col, raw_state, state, state_status, "Invalid State/Region value cleared.")
                elif state_status == "fuzzy_state_suggestion_review":
                    add_event(events, "Review Required", file_name, idx, record_id, state_col, raw_state, state, state_status, "Possible fuzzy state match. Confirm or override.", suggestion=state_suggestion, confidence=round(state_conf, 4))
                elif state_status == "ambiguous_review":
                    add_event(events, "Review Required", file_name, idx, record_id, state_col, raw_state, state, state_status, "Ambiguous State/Region. Do not guess.")
                elif state_status == "state_unknown_review":
                    add_event(events, "Review Required", file_name, idx, record_id, state_col, raw_state, state, state_status, "State/Region not recognized by mapping. Kept as-is for visibility.")

        country_validation, country_validation_issue = validate_country(country, self.allowed_countries)
        already_has_country_review = any(
            e["category"] == "Review Required" and e["field"] == country_col
            for e in events
        )
        if country_validation == "REVIEW" and not already_has_country_review:
            add_event(events, "Review Required", file_name, idx, record_id, country_col, raw_country, country, "country_validation_failed", country_validation_issue)

        country_exception = ""
        state_exception = ""
        country_issue = ""
        state_issue = ""

        for event in events:
            if event["category"] == "Review Required":
                if event["field"] == country_col and not country_exception:
                    country_exception = normalize(raw_country) or country
                    country_issue = event["issue"]
                if event["field"] == state_col and not state_exception:
                    state_exception = normalize(raw_state) or state
                    state_issue = event["issue"]

        review_required = "YES" if country_exception or state_exception else ""

        return CleanResult(
            country=country,
            state=state,
            country_status=country_status,
            state_status=state_status,
            country_exception=country_exception,
            state_exception=state_exception,
            country_issue=country_issue,
            state_issue=state_issue,
            review_required=review_required,
            events=events,
        )


def find_input_files(input_dir: Path, country_file_arg=None, state_file_arg=None) -> list[tuple[str, Path]]:
    if country_file_arg or state_file_arg:
        files = []
        if country_file_arg:
            files.append(("country_property", Path(country_file_arg)))
        if state_file_arg:
            files.append(("state_region_property", Path(state_file_arg)))
        return files

    all_files = sorted([p for p in input_dir.iterdir() if p.suffix.lower() in [".csv", ".xlsx", ".xls"]])
    if not all_files:
        raise FileNotFoundError(f"No CSV/XLSX files found in {input_dir}")

    country_matches = [p for p in all_files if "country" in p.name.lower()]
    state_matches = [p for p in all_files if "state" in p.name.lower() or "region" in p.name.lower()]

    files = []
    if country_matches:
        files.append(("country_property", country_matches[0]))
    if state_matches and (not country_matches or state_matches[0] != country_matches[0]):
        files.append(("state_region_property", state_matches[0]))

    if not files:
        files.append(("contact_location", all_files[0]))

    return files


def clean_location_file(path: Path, output_dir: Path, engine: RuleEngine, columns, output_stem: str, preserve_all_columns=True) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = read_table(path)
    record_col = find_column(df, columns["record_id"])
    country_col = find_column(df, columns["country"])
    state_col = find_column(df, columns["state"])

    if not record_col:
        raise ValueError(f"{path.name} must contain a Record ID column.")
    if not country_col and not state_col:
        raise ValueError(f"{path.name} must contain Country/Region and/or State/Region.")

    if preserve_all_columns:
        out = df.copy()
    else:
        out = pd.DataFrame({record_col: df[record_col].astype(str)})

    if not country_col:
        country_col = "Country/Region"
        df[country_col] = ""
        out[country_col] = ""
    if not state_col:
        state_col = "State/Region"
        df[state_col] = ""
        out[state_col] = ""

    countries, states, events = [], [], []
    country_exceptions, state_exceptions = [], []
    country_issues, state_issues = [], []
    review_flags = []

    for idx, row in df.iterrows():
        result = engine.clean_pair(row[country_col], row[state_col], path.name, idx, row[record_col], country_col, state_col)
        countries.append(result.country)
        states.append(result.state)
        events.extend(result.events)
        country_exceptions.append(result.country_exception)
        state_exceptions.append(result.state_exception)
        country_issues.append(result.country_issue)
        state_issues.append(result.state_issue)
        review_flags.append(result.review_required)

    out[country_col] = countries
    out[state_col] = states

    out["Country/Region - Exception Original Value"] = country_exceptions
    out["Country/Region - Exception Issue"] = country_issues
    out["State/Region - Exception Original Value"] = state_exceptions
    out["State/Region - Exception Issue"] = state_issues
    out["Review Required"] = review_flags

    write_table(out, output_dir / f"{output_stem}.csv", output_dir / f"{output_stem}.xlsx")
    return out, pd.DataFrame(events)


def pct(n, d):
    if not d:
        return 0
    return round((n / d) * 100, 4)


def build_quality_dashboard(output_dir: Path, events: pd.DataFrame, outputs: Dict[str, pd.DataFrame]) -> None:
    total_rows = sum(len(df) for df in outputs.values())
    flagged_rows = sum(int((df["Review Required"] == "YES").sum()) for df in outputs.values() if "Review Required" in df.columns)

    auto_count = int((events["category"] == "Auto Fixed").sum()) if not events.empty else 0
    review_count = int((events["category"] == "Review Required").sum()) if not events.empty else 0

    dashboard = pd.DataFrame([
        {"metric": "Total rows processed", "value": total_rows},
        {"metric": "Rows flagged for review", "value": flagged_rows},
        {"metric": "Row review rate (%)", "value": pct(flagged_rows, total_rows)},
        {"metric": "Auto fixed events", "value": auto_count},
        {"metric": "Review required events", "value": review_count},
        {"metric": "Auto fixed event rate per row (%)", "value": pct(auto_count, total_rows)},
        {"metric": "Review event rate per row (%)", "value": pct(review_count, total_rows)},
    ])

    issue_breakdown = (
        events[events["category"] == "Review Required"]
        .groupby("issue")
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
        if not events.empty else pd.DataFrame(columns=["issue", "count"])
    )
    if not issue_breakdown.empty:
        issue_breakdown["percentage_of_review_events"] = issue_breakdown["count"].apply(lambda x: pct(x, review_count))

    file_quality_rows = []
    for name, df in outputs.items():
        rows = len(df)
        flagged = int((df["Review Required"] == "YES").sum()) if "Review Required" in df.columns else 0
        file_quality_rows.append({
            "file": name,
            "rows_processed": rows,
            "rows_flagged_for_review": flagged,
            "review_rate_percent": pct(flagged, rows),
        })
    file_quality = pd.DataFrame(file_quality_rows)

    dashboard.to_csv(output_dir / "data_quality_dashboard_v11.csv", index=False)
    issue_breakdown.to_csv(output_dir / "data_quality_issue_breakdown_v11.csv", index=False)
    file_quality.to_csv(output_dir / "data_quality_by_file_v11.csv", index=False)

    if HAS_OPENPYXL:
        with pd.ExcelWriter(output_dir / "data_quality_dashboard_v11.xlsx", engine="openpyxl") as writer:
            dashboard.to_excel(writer, index=False, sheet_name="Dashboard")
            issue_breakdown.to_excel(writer, index=False, sheet_name="Issue Breakdown")
            file_quality.to_excel(writer, index=False, sheet_name="File Quality")


def build_mapping_report(output_dir: Path, events: pd.DataFrame, outputs: Dict[str, pd.DataFrame]) -> None:
    if events.empty:
        events = pd.DataFrame(columns=[
            "category", "file", "field", "original_value", "standardized_value", "issue",
            "suggestion", "confidence", "recommendation"
        ])

    auto = events[events["category"] == "Auto Fixed"].copy()
    review = events[events["category"] == "Review Required"].copy()

    auto_summary = (
        auto.groupby(["field", "original_value", "standardized_value", "issue"], dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
        if not auto.empty else pd.DataFrame(columns=["field", "original_value", "standardized_value", "issue", "count"])
    )

    review_summary = (
        review.groupby(["field", "original_value", "standardized_value", "issue", "suggestion"], dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
        if not review.empty else pd.DataFrame(columns=["field", "original_value", "standardized_value", "issue", "suggestion", "count"])
    )

    alias_usage = (
        auto[auto["issue"].isin(["standardized", "standardized_country_matched", "fuzzy_country_auto_fixed", "fuzzy_state_auto_fixed"])]
        .groupby(["field", "original_value", "standardized_value"], dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
        if not auto.empty else pd.DataFrame(columns=["field", "original_value", "standardized_value", "count"])
    )

    unknown_countries = (
        review[review["field"].str.contains("Country", case=False, na=False)]
        .groupby(["original_value", "issue", "suggestion"], dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
        if not review.empty else pd.DataFrame(columns=["original_value", "issue", "suggestion", "count"])
    )

    unknown_regions = (
        review[review["field"].str.contains("State|Region|Province", case=False, regex=True, na=False)]
        .groupby(["original_value", "issue", "suggestion"], dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
        if not review.empty else pd.DataFrame(columns=["original_value", "issue", "suggestion", "count"])
    )

    unresolved_rows = []
    for name, df in outputs.items():
        if "Review Required" in df.columns:
            flagged = df[df["Review Required"] == "YES"].copy()
            if not flagged.empty:
                flagged.insert(0, "output_file", name)
                unresolved_rows.append(flagged)

    unresolved = pd.concat(unresolved_rows, ignore_index=True) if unresolved_rows else pd.DataFrame()

    auto_summary.to_csv(output_dir / "mapping_report_auto_fixed_v11.csv", index=False)
    review_summary.to_csv(output_dir / "mapping_report_review_required_v11.csv", index=False)
    alias_usage.to_csv(output_dir / "mapping_report_alias_usage_v11.csv", index=False)
    unknown_countries.to_csv(output_dir / "mapping_report_unknown_countries_v11.csv", index=False)
    unknown_regions.to_csv(output_dir / "mapping_report_unknown_regions_v11.csv", index=False)

    if HAS_OPENPYXL:
        with pd.ExcelWriter(output_dir / "mapping_report_v11.xlsx", engine="openpyxl") as writer:
            auto_summary.to_excel(writer, index=False, sheet_name="Auto Mappings")
            review_summary.to_excel(writer, index=False, sheet_name="Review Mappings")
            alias_usage.to_excel(writer, index=False, sheet_name="Alias Usage")
            unknown_countries.to_excel(writer, index=False, sheet_name="Unknown Countries")
            unknown_regions.to_excel(writer, index=False, sheet_name="Unknown Regions")
            unresolved.to_excel(writer, index=False, sheet_name="Flagged Rows")


def write_reports(output_dir: Path, outputs: Dict[str, pd.DataFrame], events: pd.DataFrame) -> None:
    columns = [
        "category", "file", "row_number_excel", "record_id", "field", "original_value",
        "standardized_value", "issue", "recommendation", "suggestion", "confidence",
        "final_country", "final_state_region", "action", "review_notes"
    ]
    if events.empty:
        events = pd.DataFrame(columns=columns)
    else:
        for col in columns:
            if col not in events.columns:
                events[col] = ""
        events = events[columns].drop_duplicates()

    review = events[events["category"] == "Review Required"].copy()
    auto = events[events["category"] == "Auto Fixed"].copy()

    summary_rows = [
        {"metric": "Files processed", "value": len(outputs)},
        {"metric": "Auto Fixed events", "value": len(auto)},
        {"metric": "Review Required events", "value": len(review)},
        {"metric": "Manual merge file to edit", "value": "review_required_v11.csv"},
    ]
    for name, df in outputs.items():
        summary_rows.append({"metric": f"{name} rows processed", "value": len(df)})
        summary_rows.append({"metric": f"{name} rows flagged for review", "value": int((df["Review Required"] == "YES").sum())})

    if not events.empty:
        for (category, issue), group in events.groupby(["category", "issue"]):
            summary_rows.append({"metric": f"{category} - {issue}", "value": len(group)})

    summary = pd.DataFrame(summary_rows)
    summary.to_csv(output_dir / "summary_v11.csv", index=False)
    review.to_csv(output_dir / "review_required_v11.csv", index=False)
    auto.to_csv(output_dir / "auto_fixed_log_v11.csv", index=False)
    events.to_csv(output_dir / "qa_events_v11.csv", index=False)

    if HAS_OPENPYXL:
        with pd.ExcelWriter(output_dir / "qa_report_v11.xlsx", engine="openpyxl") as writer:
            summary.to_excel(writer, index=False, sheet_name="Summary")
            review.to_excel(writer, index=False, sheet_name="Review Required")
            auto.to_excel(writer, index=False, sheet_name="Auto Fixed Log")
            events.to_excel(writer, index=False, sheet_name="All Events")

    build_mapping_report(output_dir, events, outputs)
    build_quality_dashboard(output_dir, events, outputs)

    total_rows = sum(len(df) for df in outputs.values())
    flagged_rows = sum(int((df["Review Required"] == "YES").sum()) for df in outputs.values() if "Review Required" in df.columns)

    (output_dir / "run_summary_v11.txt").write_text(
        f"Files processed: {len(outputs)}\n"
        f"Total rows processed: {total_rows}\n"
        f"Rows flagged for review: {flagged_rows}\n"
        f"Auto Fixed events: {len(auto)}\n"
        f"Review Required events: {len(review)}\n"
        "Cleaned files preserve original columns and include exception columns at the end.\n"
        "Reports generated: mapping_report_v11.xlsx, data_quality_dashboard_v11.xlsx, qa_report_v11.xlsx.\n",
        encoding="utf-8"
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", default="input")
    parser.add_argument("--config-dir", default="config")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--country-file", default=None)
    parser.add_argument("--state-file", default=None)
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    config_dir = Path(args.config_dir)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    config = load_json(config_dir / "config.json")
    rules = {**DEFAULT_RULES, **load_json(config_dir / "rules.json")}
    columns = config.get("columns", DEFAULT_COLUMNS)
    preserve_all_columns = bool(rules.get("preserve_all_columns", True))

    country_key_file = config_dir / "hubspot_country_key.csv"
    state_key_file = config_dir / "hubspot_state_key.csv"

    if not country_key_file.exists():
        raise FileNotFoundError("Missing config/hubspot_country_key.csv")
    if not state_key_file.exists():
        raise FileNotFoundError("Missing config/hubspot_state_key.csv")

    allowed_countries = extract_country_key_values(country_key_file)
    _, state_code_to_name, state_name_to_name = extract_state_key_values(state_key_file)

    country_mapping = build_country_mapping(config_dir, allowed_countries)
    state_lookup = build_state_lookup(config_dir, state_code_to_name, state_name_to_name)
    engine = RuleEngine(country_mapping, state_lookup, allowed_countries, rules)

    files = find_input_files(input_dir, args.country_file, args.state_file)

    print("==================================")
    print("Contact Data Standardizer v11 Dashboard Edition")
    print("==================================")
    print(f"Input dir:    {input_dir}")
    print(f"Country key:  {country_key_file}")
    print(f"State key:    {state_key_file}")
    print(f"Output dir:   {output_dir}")
    print(f"Preserve all columns: {preserve_all_columns}")
    print("----------------------------------")

    outputs = {}
    all_events = []

    for label, path in files:
        output_stem = f"cleaned_{label}_v11"
        print(f"Processing {label}: {path}")
        cleaned, events = clean_location_file(path, output_dir, engine, columns, output_stem, preserve_all_columns)
        outputs[label] = cleaned
        if not events.empty:
            all_events.append(events)

    events = pd.concat(all_events, ignore_index=True) if all_events else pd.DataFrame()
    write_reports(output_dir, outputs, events)

    review_count = len(events[events["category"] == "Review Required"]) if not events.empty else 0
    auto_count = len(events[events["category"] == "Auto Fixed"]) if not events.empty else 0
    flagged_rows = sum(int((df["Review Required"] == "YES").sum()) for df in outputs.values() if "Review Required" in df.columns)

    print("----------------------------------")
    print(f"Files processed:          {len(outputs)}")
    print(f"Rows flagged for review:  {flagged_rows}")
    print(f"Auto Fixed events:        {auto_count}")
    print(f"Review Required events:   {review_count}")
    print("----------------------------------")
    print("Done.")
    print("Check output/mapping_report_v11.xlsx")
    print("Check output/data_quality_dashboard_v11.xlsx")
    print("Check output/review_required_v11.csv")
    print("==================================")


if __name__ == "__main__":
    main()
