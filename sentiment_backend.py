from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple, Any

import pandas as pd


RESULTS_SHEET_NAME = "Results"

PROMPT_COL_SEGMENTATION = "Prompt"
AREA_COL_SEGMENTATION = "Sentiment area"
QUESTION_COL_RESULTS = "question"
SENTIMENT_COL_RESULTS = "sentiment"

PROMPT_LEVEL_SHEET = "Prompt Level Averages"
AREA_SUMMARY_SHEET = "Sentiment Area Summary"
MATCHED_ROWS_SHEET = "Matched Results"
UNMATCHED_ROWS_SHEET = "Unmatched Results"

PLACEHOLDER_PATTERN = re.compile(
    r"\[(?:brand|varum[aä]rke|varumaerke|varumarke|mærke|merke|marca|marque)\]",
    flags=re.IGNORECASE,
)


@dataclass(frozen=True)
class PromptPattern:
    original_prompt: str
    output_prompt: str
    sentiment_area: str
    regex: re.Pattern
    sort_order: int


class SentimentKingError(Exception):
    """Raised for user-facing validation errors."""


def _canonical_col_map(columns: Iterable[object]) -> Dict[str, object]:
    return {str(c).strip().lower(): c for c in columns}


def _require_column(df: pd.DataFrame, wanted: str, source_name: str) -> object:
    col_map = _canonical_col_map(df.columns)
    key = wanted.strip().lower()
    if key not in col_map:
        available = ", ".join(map(str, df.columns))
        raise SentimentKingError(
            f"Missing required column '{wanted}' in {source_name}. Available columns: {available}"
        )
    return col_map[key]


def _normalize_text_for_matching(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("’", "'").replace("‘", "'").replace("`", "'")
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _placeholder_to_output_prompt(prompt: str) -> str:
    return PLACEHOLDER_PATTERN.sub("Brand", prompt)


def _compile_prompt_regex(prompt: str) -> re.Pattern:
    """Convert a segmented prompt with [Brand]/[Varumärke] into a regex.

    The regex captures the brand value from the uploaded Results question.
    Whitespace is flexible; punctuation/text must otherwise match.
    """
    normalized = _normalize_text_for_matching(prompt)
    matches = list(PLACEHOLDER_PATTERN.finditer(normalized))
    if not matches:
        raise SentimentKingError(
            f"Prompt is missing a brand placeholder like [Brand] or [Varumärke]: {prompt}"
        )

    parts: List[str] = []
    last = 0
    for idx, match in enumerate(matches):
        literal = normalized[last : match.start()]
        parts.append(_escape_literal_with_flexible_spaces(literal))
        # Named groups cannot be repeated with the same name, so only name the first one.
        if idx == 0:
            parts.append(r"(?P<brand>.+?)")
        else:
            parts.append(r".+?")
        last = match.end()
    parts.append(_escape_literal_with_flexible_spaces(normalized[last:]))

    pattern = "^" + "".join(parts) + "$"
    return re.compile(pattern, flags=re.IGNORECASE)


def _escape_literal_with_flexible_spaces(text: str) -> str:
    escaped = re.escape(text)
    # re.escape turns spaces into either '\ ' or literal spaces depending on Python version.
    escaped = escaped.replace(r"\ ", r"\s+").replace(" ", r"\s+")
    return escaped


def _clean_brand(raw_brand: object) -> str:
    brand = _normalize_text_for_matching(raw_brand)
    brand = brand.strip(" \t\n\r'\".,;:!?()[]{}")
    # Prevent accidental giant captures if a bad prompt pattern slips through.
    brand = re.sub(r"\s+", " ", brand)
    return brand


def read_segmentation_file(file_obj) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_obj)
    except Exception as exc:
        raise SentimentKingError(f"Could not read segmentation Excel file: {exc}") from exc

    prompt_col = _require_column(df, PROMPT_COL_SEGMENTATION, "segmentation file")
    area_col = _require_column(df, AREA_COL_SEGMENTATION, "segmentation file")

    df = df[[prompt_col, area_col]].copy()
    df.columns = [PROMPT_COL_SEGMENTATION, AREA_COL_SEGMENTATION]
    df[PROMPT_COL_SEGMENTATION] = df[PROMPT_COL_SEGMENTATION].map(_normalize_text_for_matching)
    df[AREA_COL_SEGMENTATION] = df[AREA_COL_SEGMENTATION].astype(str).str.strip()
    df = df[(df[PROMPT_COL_SEGMENTATION] != "") & (df[AREA_COL_SEGMENTATION] != "")]

    if df.empty:
        raise SentimentKingError("Segmentation file has no usable Prompt/Sentiment area rows.")

    return df.reset_index(drop=True)


def build_prompt_patterns(segmentation_df: pd.DataFrame) -> List[PromptPattern]:
    patterns: List[PromptPattern] = []
    seen: set[Tuple[str, str]] = set()

    for idx, row in segmentation_df.iterrows():
        original_prompt = _normalize_text_for_matching(row[PROMPT_COL_SEGMENTATION])
        sentiment_area = str(row[AREA_COL_SEGMENTATION]).strip()
        output_prompt = _placeholder_to_output_prompt(original_prompt)
        key = (output_prompt.lower(), sentiment_area.lower())
        if key in seen:
            continue
        seen.add(key)
        patterns.append(
            PromptPattern(
                original_prompt=original_prompt,
                output_prompt=output_prompt,
                sentiment_area=sentiment_area,
                regex=_compile_prompt_regex(original_prompt),
                sort_order=idx,
            )
        )

    if not patterns:
        raise SentimentKingError("No valid prompt patterns could be built from the segmentation file.")

    # Longer prompts first reduces bad partial-looking captures in very similar prompt families.
    return sorted(patterns, key=lambda p: (-len(p.original_prompt), p.sort_order))


def read_results_files(files: Iterable) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []
    for file_obj in files:
        file_name = getattr(file_obj, "name", "uploaded file")
        try:
            df = pd.read_excel(file_obj, sheet_name=RESULTS_SHEET_NAME)
        except ValueError as exc:
            raise SentimentKingError(f"'{file_name}' does not contain a '{RESULTS_SHEET_NAME}' sheet.") from exc
        except Exception as exc:
            raise SentimentKingError(f"Could not read '{file_name}': {exc}") from exc

        question_col = _require_column(df, QUESTION_COL_RESULTS, f"{file_name} / {RESULTS_SHEET_NAME}")
        sentiment_col = _require_column(df, SENTIMENT_COL_RESULTS, f"{file_name} / {RESULTS_SHEET_NAME}")

        keep_cols = [question_col, sentiment_col]
        optional_cols = []
        for optional in ["platform", "model", "run_date", "country"]:
            col_map = _canonical_col_map(df.columns)
            if optional in col_map:
                optional_cols.append(col_map[optional])
        keep_cols += optional_cols

        small = df[keep_cols].copy()
        rename = {question_col: QUESTION_COL_RESULTS, sentiment_col: SENTIMENT_COL_RESULTS}
        for col in optional_cols:
            rename[col] = str(col).strip().lower()
        small = small.rename(columns=rename)
        small["source_file"] = file_name
        frames.append(small)

    if not frames:
        raise SentimentKingError("Upload at least one results Excel file.")

    combined = pd.concat(frames, ignore_index=True)
    combined[QUESTION_COL_RESULTS] = combined[QUESTION_COL_RESULTS].map(_normalize_text_for_matching)
    combined[SENTIMENT_COL_RESULTS] = pd.to_numeric(combined[SENTIMENT_COL_RESULTS], errors="coerce")
    combined = combined[(combined[QUESTION_COL_RESULTS] != "") & combined[SENTIMENT_COL_RESULTS].notna()]

    if combined.empty:
        raise SentimentKingError("No usable question/sentiment rows found in the uploaded results files.")

    return combined.reset_index(drop=True)


def match_results_to_segments(results_df: pd.DataFrame, patterns: List[PromptPattern]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    matched_rows: List[dict] = []
    unmatched_rows: List[dict] = []

    for _, row in results_df.iterrows():
        question = _normalize_text_for_matching(row[QUESTION_COL_RESULTS])
        match_data: Optional[dict] = None

        for pattern in patterns:
            match = pattern.regex.match(question)
            if match:
                brand = _clean_brand(match.groupdict().get("brand", ""))
                if not brand:
                    continue
                match_data = {
                    "Prompt": pattern.output_prompt,
                    "Sentiment area": pattern.sentiment_area,
                    "Brand": brand,
                    "question": question,
                    "sentiment": row[SENTIMENT_COL_RESULTS],
                    "source_file": row.get("source_file", ""),
                }
                for optional in ["platform", "model", "run_date", "country"]:
                    if optional in row.index:
                        match_data[optional] = row.get(optional, "")
                break

        if match_data is None:
            unmatched_rows.append(row.to_dict())
        else:
            matched_rows.append(match_data)

    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)

    if matched_df.empty:
        raise SentimentKingError(
            "None of the Results questions matched the segmentation prompts. Check that the wording is identical except for the brand placeholder."
        )

    return matched_df, unmatched_df


def make_prompt_level_averages(matched_df: pd.DataFrame, segmentation_df: pd.DataFrame) -> pd.DataFrame:
    grouped = (
        matched_df.groupby(["Prompt", "Sentiment area", "Brand"], dropna=False)["sentiment"]
        .mean()
        .reset_index()
    )

    pivot = grouped.pivot_table(
        index=["Prompt", "Sentiment area"],
        columns="Brand",
        values="sentiment",
        aggfunc="mean",
    ).reset_index()

    brand_cols = sorted([c for c in pivot.columns if c not in ["Prompt", "Sentiment area"]], key=str.casefold)
    pivot = pivot[["Prompt", "Sentiment area"] + brand_cols]

    order_df = segmentation_df.copy()
    order_df["Prompt"] = order_df[PROMPT_COL_SEGMENTATION].map(_placeholder_to_output_prompt)
    order_df = order_df.rename(columns={AREA_COL_SEGMENTATION: "Sentiment area"})
    order_df["__order"] = range(len(order_df))
    order_df = order_df[["Prompt", "Sentiment area", "__order"]].drop_duplicates()

    pivot = pivot.merge(order_df, on=["Prompt", "Sentiment area"], how="left")
    pivot = pivot.sort_values(["__order", "Sentiment area", "Prompt"], na_position="last").drop(columns="__order")
    return pivot.reset_index(drop=True)


def make_area_summary(prompt_level_df: pd.DataFrame) -> pd.DataFrame:
    brand_cols = [c for c in prompt_level_df.columns if c not in ["Prompt", "Sentiment area"]]
    if not brand_cols:
        raise SentimentKingError("No brand columns were produced. Check prompt matching and brand extraction.")

    summary = prompt_level_df.groupby("Sentiment area", sort=False)[brand_cols].mean().reset_index()
    summary = summary.rename(columns={"Sentiment area": "Sentiment Area"})
    summary["Industry Average"] = summary[brand_cols].mean(axis=1, skipna=True)
    return summary


def build_output_workbook(
    segmentation_file,
    results_files: Iterable,
    decimals: Optional[int] = None,
) -> Tuple[bytes, Dict[str, Any]]:
    segmentation_df = read_segmentation_file(segmentation_file)
    patterns = build_prompt_patterns(segmentation_df)
    results_df = read_results_files(results_files)
    matched_df, unmatched_df = match_results_to_segments(results_df, patterns)

    prompt_level = make_prompt_level_averages(matched_df, segmentation_df)
    area_summary = make_area_summary(prompt_level)

    if decimals is not None:
        numeric_cols_prompt = prompt_level.select_dtypes(include="number").columns
        numeric_cols_summary = area_summary.select_dtypes(include="number").columns
        prompt_level[numeric_cols_prompt] = prompt_level[numeric_cols_prompt].round(decimals)
        area_summary[numeric_cols_summary] = area_summary[numeric_cols_summary].round(decimals)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        prompt_level.to_excel(writer, sheet_name=PROMPT_LEVEL_SHEET, index=False)
        area_summary.to_excel(writer, sheet_name=AREA_SUMMARY_SHEET, index=False)
        matched_df.to_excel(writer, sheet_name=MATCHED_ROWS_SHEET, index=False)
        if not unmatched_df.empty:
            unmatched_df.to_excel(writer, sheet_name=UNMATCHED_ROWS_SHEET, index=False)

        _format_workbook(writer.book)

    stats = {
        "results_rows_used": int(len(results_df)),
        "matched_rows": int(len(matched_df)),
        "unmatched_rows": int(len(unmatched_df)),
        "brands": int(len([c for c in prompt_level.columns if c not in ["Prompt", "Sentiment area"]])),
        "sentiment_areas": int(area_summary["Sentiment Area"].nunique()),
        "prompt_level_preview": prompt_level,
        "area_summary_preview": area_summary,
    }
    return output.getvalue(), stats


def _format_workbook(workbook) -> None:
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True)

    for ws in workbook.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill

        for column_cells in ws.columns:
            max_length = 0
            col_idx = column_cells[0].column
            col_letter = get_column_letter(col_idx)
            for cell in column_cells[:200]:
                if cell.value is None:
                    continue
                max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(max_length + 2, 12), 70)

        # Format numeric score columns.
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0.000"
