import streamlit as st

from sentiment_backend import SentimentKingError, build_output_workbook


st.set_page_config(page_title="Sentiment King", page_icon="👑", layout="wide")

st.title("Sentiment King")
st.write(
    "Upload one segmentation file and one or more sentiment export files. "
    "The app creates a downloadable Excel file with prompt-level averages and sentiment-area summaries."
)

with st.expander("Expected file structure", expanded=False):
    st.markdown(
        """
**Segmentation file**
- Must contain columns: `Prompt`, `Sentiment area`
- Prompts must use a placeholder such as `[Brand]` or `[Varumärke]`

**Sentiment export files**
- Must contain a `Results` sheet
- `Results` must contain columns: `question`, `sentiment`
- Duplicate rows from different platforms, for example ChatGPT and Gemini, are averaged automatically
        """
    )

segmentation_file = st.file_uploader(
    "1. Upload prompt segmentation Excel file",
    type=["xlsx"],
    accept_multiple_files=False,
)

results_files = st.file_uploader(
    "2. Upload sentiment export Excel files",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Upload as many standardized exports as needed. The app has no hard-coded file limit.",
)

round_output = st.checkbox("Round output scores to 3 decimals", value=False)

if st.button("Create enriched Excel", type="primary"):
    if segmentation_file is None:
        st.error("Upload the segmentation file first.")
    elif not results_files:
        st.error("Upload at least one sentiment export file.")
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

                if stats["unmatched_rows"]:
                    st.warning(
                        "Some rows did not match the segmentation prompts. "
                        "They are included in the `Unmatched Results` sheet."
                    )

                st.subheader("Sentiment-area summary")
                st.dataframe(stats["area_summary_preview"], use_container_width=True)

                st.subheader("Prompt-level averages")
                st.dataframe(stats["prompt_level_preview"], use_container_width=True)

                st.download_button(
                    label="Download enriched Excel",
                    data=output_bytes,
                    file_name="sentiment_king_enriched.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
