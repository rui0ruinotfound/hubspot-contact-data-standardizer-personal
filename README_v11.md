# Contact Data Standardizer v11 Dashboard Edition

v11 is the reporting-focused enterprise version of the HubSpot Country/Region + State/Region cleanup tool.

## Main upgrades from v10

- Keeps all original columns.
- Only updates `Country/Region` and `State/Region`.
- Adds exception columns at the end of cleaned files.
- Generates a stronger `mapping_report_v11.xlsx`.
- Generates `data_quality_dashboard_v11.xlsx`.
- Keeps all logic local: no AI calls, no API calls, no uploads.

## Run

Put raw file(s) in `input/`, then:

```bash
python3 standardize_contact_data_v11.py
```

## Main output files

```text
output/cleaned_*_v11.xlsx
output/review_required_v11.csv
output/qa_report_v11.xlsx
output/mapping_report_v11.xlsx
output/data_quality_dashboard_v11.xlsx
```

## Mapping report sheets

`mapping_report_v11.xlsx` includes:

- Auto Mappings
- Review Mappings
- Alias Usage
- Unknown Countries
- Unknown Regions
- Flagged Rows

## Data quality dashboard sheets

`data_quality_dashboard_v11.xlsx` includes:

- Dashboard
- Issue Breakdown
- File Quality

## Manual review workflow

Open:

```text
output/review_required_v11.csv
```

Fill:

```text
final_country
final_state_region
action
```

Allowed actions:

```text
update
blank
keep
skip
```

Then run:

```bash
python3 merge_review_v11.py
```

Final files:

```text
output/final_*_v11.xlsx
```

## Privacy

This is a local Python script. It does not call OpenAI, Claude, Gemini, HubSpot API, or any external API. If you run it on your work computer, the data stays on that computer.
