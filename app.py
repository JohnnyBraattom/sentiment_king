import re
from datetime import datetime

import streamlit as st

from sentiment_backend import SentimentKingError, build_output_workbook


st.set_page_config(page_title="Sentiment King", page_icon="👑", layout="wide")

st.title("Sentiment King")
st.write(
    "Upload one prompt segmentation file and one or more CIM aided brand sentiment exports. "
    "The app creates a downloadable Excel workbook with sentiment averages and citation rankings."
)


def make_download_filename(client_name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", client_name.strip())
    cleaned = re.sub(r"_+", "_", cleaned).strip("_")
    if not cleaned:
        cleaned = "CLIENT"
    date_stamp = datetime.now().strftime("%Y%m%d")
    time_stamp = datetime.now().strftime("%H%M%S")
    return f"Sentiment_King_{cleaned}_{date_stamp}_{time_stamp}.xlsx"

st.subheader("How to use the app")
st.markdown(
    """
1. Enter the **client name**. This is used in the downloaded Excel filename.
2. Upload the **prompt segmentation file**.
3. Upload one or multiple **CIM aided brand sentiment export Excel files**.
4. Click **Create enriched Excel**.
5. Review the calculated tables in the UI.
6. Download the enriched Excel workbook.
    """
)

with st.expander("Required file structure", expanded=True):
    st.markdown(
        """
### Prompt segmentation file
This file tells the app which sentiment area each prompt belongs to.

It must contain these columns:

- `Prompt`
- `Sentiment area`

The prompts must use a brand placeholder where the brand name appears. Supported placeholders include:

- `[Brand]`
- `[Varumärke]`
- `[Varumarke]`

Example:

`What is the general cutting quality of [Brand]'s robotic lawn mowers?`

The prompt wording should match the wording in the CIM exports, except that the CIM exports contain the actual brand name instead of the placeholder.

### CIM aided brand sentiment export files
Upload one or multiple CIM exports created for aided brand sentiment purposes. These files should not be altered after exporting from CIM, instead just upload them as they are formatted when exported. 

Each export must contain a `Results` sheet with these columns:

- `question`
- `sentiment`

Each export should also contain a `Citations` sheet with these columns:

- `question`
- `url`
- `domain`

The app detects the brand directly from the branded prompt text in `question`. It does not rely on file names or a brand column.

### Download filename
The downloaded Excel file uses this structure:

`Sentiment_King_CLIENT_YYYYMMDD_HHMMSS.xlsx`

`CLIENT` is replaced by the client name entered in the UI. Spaces and special characters are converted to underscores.

### What the app calculates
- Average sentiment per prompt and brand
- Average sentiment per sentiment area and brand
- Industry average per sentiment area
- Top 3 cited domains per brand
- Top 10 cited URLs per brand
- Top domains and URLs per brand within each sentiment area
        """
    )

client_name = st.text_input(
    "1. Client name for export filename",
    value="CLIENT",
    help="Used in the downloaded Excel filename, for example Sentiment_King_Plantagen_20260424_143012.xlsx.",
)

segmentation_file = st.file_uploader(
    "2. Upload prompt segmentation Excel file",
    type=["xlsx"],
    accept_multiple_files=False,
)

results_files = st.file_uploader(
    "3. Upload CIM aided brand sentiment export Excel files",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Upload one or multiple standardized CIM exports. There is no hard-coded file limit.",
)

round_output = st.checkbox("Round sentiment output scores to 3 decimals", value=False)

if st.button("Create enriched Excel", type="primary"):
    if segmentation_file is None:
        st.error("Upload the prompt segmentation file first.")
    elif not results_files:
        st.error("Upload at least one CIM sentiment export file.")
    else:
        with st.spinner("Processing files..."):
            try:
                output_bytes, stats = build_output_workbook(
                    segmentation_file=segmentation_file,
                    results_files=results_files,
                    decimals=3 if round_output else None,
                )
            except SentimentKingError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Unexpected error: {exc}")
            else:
                st.success("Enriched Excel file created.")

                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("Rows used", stats["results_rows_used"])
                c2.metric("Matched rows", stats["matched_rows"])
                c3.metric("Unmatched rows", stats["unmatched_rows"])
                c4.metric("Brands", stats["brands"])
                c5.metric("Sentiment areas", stats["sentiment_areas"])

                c6, c7, c8 = st.columns(3)
                c6.metric("Citation rows", stats["citation_rows_used"])
                c7.metric("Matched citations", stats["matched_citation_rows"])
                c8.metric("Unmatched citations", stats["unmatched_citation_rows"])

                if stats["unmatched_rows"]:
                    st.warning(
                        "Some Results rows did not match the segmentation prompts. "
                        "They are included in the `Unmatched Results` sheet."
                    )

                if stats["unmatched_citation_rows"]:
                    st.warning(
                        "Some Citations rows did not match the segmentation prompts. "
                        "They are included in the `Unmatched Citations` sheet."
                    )

                if stats.get("missing_citation_files"):
                    st.warning(
                        "These files did not have a `Citations` sheet and were skipped for citation rankings: "
                        + ", ".join(stats["missing_citation_files"])
                    )

                st.subheader("Sentiment-area summary")
                st.dataframe(stats["area_summary_preview"], use_container_width=True)

                st.subheader("Prompt-level averages")
                st.dataframe(stats["prompt_level_preview"], use_container_width=True)

                st.subheader("Top cited domains, total")
                st.dataframe(stats["top_domains_total_preview"], use_container_width=True)

                st.subheader("Top cited URLs, total")
                st.dataframe(stats["top_urls_total_preview"], use_container_width=True)

                with st.expander("Citation rankings by sentiment area", expanded=False):
                    st.markdown("**Top cited domains by sentiment area**")
                    st.dataframe(stats["top_domains_by_area_preview"], use_container_width=True)
                    st.markdown("**Top cited URLs by sentiment area**")
                    st.dataframe(stats["top_urls_by_area_preview"], use_container_width=True)

                st.download_button(
                    label="Download enriched Excel",
                    data=output_bytes,
                    file_name=make_download_filename(client_name),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
