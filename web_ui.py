import tempfile
from pathlib import Path
from typing import List, Tuple

import gradio as gr
import pandas as pd

from extractor import (
    DEFAULT_COLUMNS,
    extract_annotations,
    write_csv,
    write_json,
    write_xlsx,
)


def process_pdfs(
    files: List[gr.File],
    output_format: str,
    include_color: bool,
    include_text: bool,
    sheet_per_file: bool,
    skip_empty: bool,
) -> Tuple[str, str]:
    if not files:
        return "", "No files provided."

    all_rows = []
    warnings = []
    for file_obj in files:
        pdf_path = Path(file_obj.name)
        rows, file_warnings = extract_annotations(
            pdf_path,
            include_color=include_color,
            include_text=include_text,
            skip_empty=skip_empty,
        )
        all_rows.extend(rows)
        warnings.extend(file_warnings)

    if not all_rows:
        message = "No annotations extracted."
        if warnings:
            message += "\n" + "\n".join(f"Warning: {w}" for w in warnings)
        return "", message

    df = pd.DataFrame(all_rows)
    df = df[DEFAULT_COLUMNS]
    df.sort_values(by=["source_file", "page_number", "date_created"], inplace=True)

    suffix = ".xlsx" if output_format == "xlsx" else ".csv" if output_format == "csv" else ".json"
    _, tmp_path = tempfile.mkstemp(suffix=suffix)
    output_path = Path(tmp_path)

    if output_format == "csv":
        write_csv(output_path, df, overwrite=True)
    elif output_format == "json":
        write_json(output_path, all_rows, overwrite=True)
    else:
        write_xlsx(output_path, df, overwrite=True, sheet_per_file=sheet_per_file)

    summary = f"Extracted {len(df)} annotations from {len(files)} PDFs."
    if warnings:
        summary += "\n" + "\n".join(f"Warning: {w}" for w in warnings)

    return str(output_path), summary


def build_ui() -> gr.Blocks:
    with gr.Blocks(title="PDF Comment Extractor") as demo:
        gr.Markdown("# PDF Comment & Annotation Extractor")
        gr.Markdown("Upload PDFs and download extracted annotations.")

        with gr.Row():
            file_input = gr.File(
                label="PDF files",
                file_types=[".pdf"],
                file_count="multiple",
            )
            output_format = gr.Dropdown(
                label="Output format",
                choices=["xlsx", "csv", "json"],
                value="xlsx",
            )

        with gr.Row():
            include_color = gr.Checkbox(label="Include color", value=True)
            include_text = gr.Checkbox(label="Extract highlighted text", value=True)
            sheet_per_file = gr.Checkbox(label="XLSX: sheet per file", value=False)
            skip_empty = gr.Checkbox(label="Skip PDFs with no annotations", value=False)

        run_btn = gr.Button("Extract")
        output_file = gr.File(label="Download output")
        output_text = gr.Textbox(label="Summary", lines=6)

        run_btn.click(
            fn=process_pdfs,
            inputs=[
                file_input,
                output_format,
                include_color,
                include_text,
                sheet_per_file,
                skip_empty,
            ],
            outputs=[output_file, output_text],
        )

    return demo


if __name__ == "__main__":
    build_ui().launch()
