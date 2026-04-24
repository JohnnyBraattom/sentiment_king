# Sentiment King

Streamlit Cloud app for combining sentiment Excel exports into one enriched Excel workbook.

## Files

- `app.py` - Streamlit frontend
- `sentiment_backend.py` - processing and Excel export logic
- `requirements.txt` - Python dependencies for Streamlit Cloud

## Expected inputs

### Segmentation Excel
Must contain:

- `Prompt`
- `Sentiment area`

The prompt should include a brand placeholder, for example:

- `[Brand]`
- `[Varumärke]`
- `[Varumarke]`

### Sentiment export Excel files
Each uploaded export must contain a sheet named:

- `Results`

The `Results` sheet must contain:

- `question`
- `sentiment`

Optional columns such as `platform`, `model`, `run_date`, and `country` are preserved in the matched detail sheet if present.

## Output workbook

The generated workbook contains:

1. `Prompt Level Averages`
   - Prompt
   - Sentiment area
   - one column per extracted brand
   - average sentiment per prompt and brand

2. `Sentiment Area Summary`
   - Sentiment Area
   - one column per brand
   - Industry Average

3. `Matched Results`
   - detail rows used in the calculations

4. `Unmatched Results`
   - only added if any rows could not be matched

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud deployment

1. Create a GitHub repository.
2. Upload `app.py`, `sentiment_backend.py`, `requirements.txt`, and this README.
3. In Streamlit Cloud, create a new app from the repo.
4. Set the main file path to `app.py`.
