import io
import os
from dataclasses import dataclass
from typing import List, Optional, Tuple

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle


ALLOWED_EXTENSIONS = {"xls", "xlsx"}
DEFAULT_DOWNLOAD_NAME = "labels.pdf"

# Label layout constants (letter paper, 3 columns x 10 rows)
LABEL_COLUMNS = 3
LABEL_ROWS = 10
LABEL_WIDTH = 2.625 * inch
LABEL_HEIGHT = 1.0 * inch
HORIZONTAL_GAP = 0.1 * inch
VERTICAL_GAP = 0 * inch
PAGE_MARGIN_X = 0.2 * inch
PAGE_MARGIN_Y = 0.4 * inch
TITLE_FONT = "Helvetica-Bold"
BODY_FONT = "Helvetica"
TITLE_FONT_SIZE = 12
BODY_FONT_SIZE = 10
COUNT_FONT_SIZE = 12
CONTENT_PADDING = 0.2 * inch
SUMMARY_TITLE_FONT_SIZE = 18
SUMMARY_BODY_FONT_SIZE = 11


@dataclass
class LabelCard:
    name: str
    count: Optional[int]
    doubles: Optional[int] = None
    singles: Optional[int] = None


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", "dev-secret")
    app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10 MB upload limit

    @app.route("/", methods=["GET", "POST"])
    def index():
        if request.method == "POST":
            upload = request.files.get("data_file")
            if upload is None or upload.filename == "":
                flash("Please choose an Excel file before uploading.", "error")
                return redirect(url_for("index"))

            if not allowed_file(upload.filename):
                flash("File type not supported. Upload an .xls or .xlsx file.", "error")
                return redirect(url_for("index"))

            try:
                upload.stream.seek(0)
                dataframe = pd.read_excel(upload, dtype=str).fillna("")
            except Exception as exc:  # keep broad: pandas raises different errors per engine
                flash(f"Could not read Excel file: {exc}", "error")
                return redirect(url_for("index"))

            if dataframe.empty:
                flash("The uploaded workbook does not contain any rows to print.", "error")
                return redirect(url_for("index"))

            try:
                pdf_buffer = build_labels_pdf(dataframe)
            except Exception as exc:
                flash(f"Failed to generate PDF: {exc}", "error")
                return redirect(url_for("index"))

            return send_file(
                pdf_buffer,
                mimetype="application/pdf",
                as_attachment=True,
                download_name=DEFAULT_DOWNLOAD_NAME,
            )

        return render_template("upload.html")

    return app


def allowed_file(filename: str) -> bool:
    _, dot, extension = filename.rpartition(".")
    return bool(dot) and extension.lower() in ALLOWED_EXTENSIONS


def build_labels_pdf(dataframe: pd.DataFrame) -> io.BytesIO:
    buffer = io.BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=letter)
    page_width, page_height = letter

    label_cards, dine_in_summary = _prepare_label_cards(dataframe)
    labels_per_page = LABEL_COLUMNS * LABEL_ROWS
    labels_drawn = False

    for label_index, card in enumerate(label_cards):
        position = label_index % labels_per_page
        column_index = position % LABEL_COLUMNS
        row_index = position // LABEL_COLUMNS

        if label_index and position == 0:
            pdf.showPage()

        label_origin_x = PAGE_MARGIN_X + column_index * (LABEL_WIDTH + HORIZONTAL_GAP)
        label_top = page_height - PAGE_MARGIN_Y - row_index * (LABEL_HEIGHT + VERTICAL_GAP)

        draw_label_card(pdf=pdf, label_top=label_top, label_left=label_origin_x, card=card)
        labels_drawn = True

    if labels_drawn:
        pdf.showPage()

    draw_dine_in_summary(pdf, dine_in_summary, page_width, page_height)
    pdf.save()
    buffer.seek(0)
    return buffer


def draw_label_card(
    pdf: canvas.Canvas,
    label_top: float,
    label_left: float,
    card: LabelCard,
) -> None:
    label_bottom = label_top - LABEL_HEIGHT
    #pdf.setStrokeColor(colors.black)
    #pdf.roundRect(label_left, label_bottom, LABEL_WIDTH, LABEL_HEIGHT, 6, stroke=1, fill=0)

    center_x = label_left + LABEL_WIDTH / 2

    if card.doubles is not None or card.singles is not None:
        title_baseline = label_top - CONTENT_PADDING - (TITLE_FONT_SIZE * 0.6)
        pdf.setFont(TITLE_FONT, TITLE_FONT_SIZE)
        pdf.drawCentredString(center_x, title_baseline, card.name)

        pdf.setFont(BODY_FONT, BODY_FONT_SIZE)
        doubles_y = title_baseline - BODY_FONT_SIZE - 6
        singles_y = doubles_y - BODY_FONT_SIZE - 4
        pdf.drawCentredString(center_x, doubles_y, f"Doubles: {card.doubles or 0}")
        pdf.drawCentredString(center_x, singles_y, f"Singles: {card.singles or 0}")
        return

    # Name placement
    if card.count is None:
        name_y = label_bottom + (LABEL_HEIGHT / 2) - (TITLE_FONT_SIZE / 2)
    else:
        name_y = label_top - CONTENT_PADDING - (TITLE_FONT_SIZE)

    pdf.setFont(TITLE_FONT, TITLE_FONT_SIZE)
    pdf.drawCentredString(center_x, name_y, card.name)

    if card.count is not None:
        pdf.setFont(TITLE_FONT, COUNT_FONT_SIZE)
        count_y = label_bottom + (LABEL_HEIGHT / 2) - (COUNT_FONT_SIZE * 0.5)
        pdf.drawCentredString(center_x, count_y, str(card.count))


def draw_dine_in_summary(
    pdf: canvas.Canvas,
    dine_in_summary: List[Tuple[str, int]],
    page_width: float,
    page_height: float,
) -> None:
    margin_x = PAGE_MARGIN_X
    margin_y = PAGE_MARGIN_Y
    pdf.setFont(TITLE_FONT, SUMMARY_TITLE_FONT_SIZE)
    pdf.drawString(margin_x, page_height - margin_y, "Dine-In Summary")

    start_y = page_height - margin_y - (SUMMARY_TITLE_FONT_SIZE + 12)

    if not dine_in_summary:
        pdf.setFont(BODY_FONT, SUMMARY_BODY_FONT_SIZE)
        pdf.drawString(margin_x, start_y, "No dine-in orders.")
        return

    data: List[List[str]] = [["Name", "Number"]]
    data.extend([[name, str(count)] for name, count in dine_in_summary])

    available_width = page_width - 2 * margin_x
    table = Table(
        data,
        colWidths=[available_width * 0.7, available_width * 0.3],
    )
    table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), TITLE_FONT),
                ("FONTSIZE", (0, 0), (-1, 0), SUMMARY_BODY_FONT_SIZE),
                ("FONTNAME", (0, 1), (-1, -1), BODY_FONT),
                ("FONTSIZE", (0, 1), (-1, -1), SUMMARY_BODY_FONT_SIZE),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("ALIGN", (1, 1), (1, -1), "CENTER"),
                ("ALIGN", (1, 0), (1, 0), "CENTER"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f4f6")]),
                ("BOX", (0, 0), (-1, -1), 1, colors.gray),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.gray),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )

    _, table_height = table.wrapOn(pdf, available_width, page_height)
    table.drawOn(pdf, margin_x, start_y - table_height)


def _prepare_label_cards(dataframe: pd.DataFrame) -> Tuple[List[LabelCard], List[Tuple[str, int]]]:
    normalized = _normalize_orders_dataframe(dataframe)
    normalized["name"] = normalized["name"].fillna("").astype(str).str.strip()
    normalized = normalized[normalized["name"] != ""]

    normalized["carry_out"] = (
        pd.to_numeric(normalized["carry_out"], errors="coerce").fillna(0).astype(int).clip(lower=0)
    )
    normalized["dine_in"] = (
        pd.to_numeric(normalized["dine_in"], errors="coerce").fillna(0).astype(int).clip(lower=0)
    )

    aggregated = (
        normalized.groupby("name", as_index=False)
        .agg({"carry_out": "sum", "dine_in": "sum"})
        .sort_values("name")
    )

    label_cards: List[LabelCard] = []
    total_doubles = 0
    total_singles = 0
    for _, row in aggregated.iterrows():
        carry_count = int(row["carry_out"])
        if carry_count <= 0:
            continue

        doubles = carry_count // 2
        singles = carry_count % 2
        total_doubles += doubles
        total_singles += singles

        total_labels = _required_label_count(carry_count)
        label_cards.append(LabelCard(name=row["name"], count=carry_count))
        for _ in range(total_labels - 1):
            label_cards.append(LabelCard(name=row["name"], count=None))

    if total_doubles or total_singles:
        label_cards.append(
            LabelCard(
                name="Pack Summary",
                count=None,
                doubles=total_doubles,
                singles=total_singles,
            )
        )

    dine_summary = [
        (row["name"], int(row["dine_in"]))
        for _, row in aggregated.iterrows()
        if int(row["dine_in"]) > 0
    ]

    return label_cards, dine_summary


def _normalize_orders_dataframe(dataframe: pd.DataFrame) -> pd.DataFrame:
    column_aliases = {
        "name": "name",
        "customer": "name",
        "carry out": "carry_out",
        "carryout": "carry_out",
        "carry-out": "carry_out",
        "dine in": "dine_in",
        "dine-in": "dine_in",
        "dinein": "dine_in",
    }

    lookup = {col.strip().lower(): col for col in dataframe.columns}
    rename_map = {}

    for alias, normalized_name in column_aliases.items():
        if alias in lookup and normalized_name not in rename_map:
            rename_map[normalized_name] = lookup[alias]

    missing = [field for field in ("name", "carry_out", "dine_in") if field not in rename_map]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    return dataframe.rename(columns={rename_map[key]: key for key in rename_map})


def _required_label_count(carry_out_total: int) -> int:
    if carry_out_total <= 0:
        return 0
    quotient, remainder = divmod(carry_out_total, 2)
    return quotient + remainder


app = create_app()


if __name__ == "__main__":
    debug_mode = True#os.environ.get("FLASK_DEBUG", "").lower() in {"1", "true", "yes"}
    app.run(debug=debug_mode, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
