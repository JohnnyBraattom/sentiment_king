# Sentiment King

Streamlit Cloud app for uploading standardized CIM aided brand sentiment Excel exports and one prompt segmentation file, then generating one enriched Excel workbook.

## Files

- `app.py` - Streamlit frontend
- `sentiment_backend.py` - Excel processing and calculations
- `requirements.txt` - Streamlit Cloud dependencies

## Expected inputs

### Prompt segmentation file

This file maps each generic prompt to a sentiment area.

Required columns:

- `Prompt`
- `Sentiment area`

The `Prompt` values must contain a brand placeholder where the brand name appears. Supported placeholders include:

- `[Brand]`
- `[Varumärke]`
- `[Varumarke]`

Example:

```text
What is the general cutting quality of [Brand]'s robotic lawn mowers?
```

The prompt wording should match the wording in the CIM exports, except that the CIM exports contain the real brand name instead of the placeholder.

### CIM aided brand sentiment export files

Upload one or multiple CIM exports created for aided brand sentiment purposes.

Each export must contain a `Results` sheet with:

- `question`
- `sentiment`

Each export should also contain a `Citations` sheet with:

- `question`
- `url`
- `domain`

Optional citation columns such as `title`, `platform`, `model`, `run_date`, `country`, and `source` are preserved when available.

The app detects the brand from the branded prompt text in the `question` column. It does not rely on file names or a brand column.

## Output sheets

- `Prompt Level Averages` - average sentiment per prompt and brand
- `Sentiment Area Summary` - average sentiment per sentiment area and brand, plus industry average
- `Top Domains Total` - each brand's top 3 cited domains overall
- `Top URLs Total` - each brand's top 10 cited URLs overall
- `Top Domains By Area` - each brand's top 3 cited domains within each sentiment area
- `Top URLs By Area` - each brand's top 10 cited URLs within each sentiment area
- `Matched Results`
- `Matched Citations`
- `Unmatched Results` if any
- `Unmatched Citations` if any

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud

Upload these files to a GitHub repo, then create a Streamlit Cloud app pointing to `app.py`.


## Export filename

The Streamlit UI includes a client name field. The downloaded Excel file is named like:

```text
Sentiment_King_CLIENT_YYYYMMDD_HHMMSS.xlsx
```

`CLIENT` is replaced by the value entered in the UI. Spaces and special characters are converted to underscores. The time component makes repeated exports unique.
