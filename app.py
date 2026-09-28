import io
import os
import re
import base64
import html

import pandas as pd
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
    Image,
)
from reportlab.pdfgen import canvas


# =========================================================
# STREAMLIT CONFIG
# =========================================================

st.set_page_config(
    page_title="BAST Generator",
    page_icon="📄",
    layout="wide",
)


# =========================================================
# SIGNATURE COMPONENT
# =========================================================

from signature_component import signature_pad


# =========================================================
# CONFIG PDF
# =========================================================

PAGE_WIDTH, PAGE_HEIGHT = A4

MARGIN_LEFT = 5 * mm
MARGIN_RIGHT = 5 * mm
MARGIN_TOP = 5 * mm
MARGIN_BOTTOM = 8 * mm

TOTAL_CONTENT_WIDTH = (
    PAGE_WIDTH
    - MARGIN_LEFT
    - MARGIN_RIGHT
)

HALF_WIDTH = TOTAL_CONTENT_WIDTH / 2


# =========================================================
# REQUIRED COLUMNS
# =========================================================

REQUIRED_COLUMNS = [
    "NO",
    "DELIVERY ORDER",
    "AIRWAYBILL",
    "PROVIDER",
    "KOLI QTY",
]


# =========================================================
# HELPER
# =========================================================

def safe_filename(text):
    text = str(text).strip()

    if not text:
        text = "BAST"

    text = re.sub(
        r'[\\/*?:"<>|]',
        "_",
        text,
    )

    return text


def parse_paste_data(text):
    """
    Membaca data hasil copy dari Excel.
    """

    if not text or not text.strip():
        return pd.DataFrame()

    text = text.strip()

    # -----------------------------------------------------
    # TSV
    # -----------------------------------------------------

    try:

        df = pd.read_csv(
            io.StringIO(text),
            sep="\t",
            dtype=str,
            keep_default_na=False,
        )

        if len(df.columns) > 1:
            return df

    except Exception:
        pass

    # -----------------------------------------------------
    # CSV
    # -----------------------------------------------------

    try:

        df = pd.read_csv(
            io.StringIO(text),
            dtype=str,
            keep_default_na=False,
        )

        if len(df.columns) > 1:
            return df

    except Exception:
        pass

    # -----------------------------------------------------
    # AUTO
    # -----------------------------------------------------

    try:

        df = pd.read_csv(
            io.StringIO(text),
            sep=None,
            engine="python",
            dtype=str,
            keep_default_na=False,
        )

        return df

    except Exception:

        return pd.DataFrame()


def fix_broken_rows(df):

    if df.empty:
        return df

    df = df.copy()

    # Normalisasi nama kolom
    df.columns = [
        str(col).strip().upper()
        for col in df.columns
    ]

    # Buang Unnamed
    df = df.loc[
        :,
        ~df.columns.str.contains(
            "^UNNAMED",
            case=False,
            regex=True,
        )
    ]

    # Pastikan semua kolom ada
    for col in REQUIRED_COLUMNS:

        if col not in df.columns:
            df[col] = ""

    # Ambil hanya kolom yang digunakan
    df = df[
        REQUIRED_COLUMNS
    ].copy()

    # Bersihkan isi
    for col in df.columns:

        df[col] = (
            df[col]
            .astype(str)
            .str.replace(
                "\n",
                " ",
                regex=False,
            )
            .str.replace(
                "\r",
                " ",
                regex=False,
            )
            .str.strip()
        )

    # Isi NO yang kosong
    for idx in df.index:

        if not str(
            df.at[idx, "NO"]
        ).strip():

            df.at[idx, "NO"] = str(
                idx + 1
            )

    # KOLI QTY
    df["KOLI QTY"] = pd.to_numeric(
        df["KOLI QTY"],
        errors="coerce",
    ).fillna(0)

    df["KOLI QTY"] = (
        df["KOLI QTY"]
        .astype(int)
    )

    return df


def validate_file(df):

    if df.empty:
        return False, "Data kosong."

    missing = [
        col
        for col in REQUIRED_COLUMNS
        if col not in df.columns
    ]

    if missing:

        return False, (
            "Kolom berikut belum ada: "
            + ", ".join(missing)
        )

    return True, ""


# =========================================================
# SIGNATURE HELPER
# =========================================================

def signature_exists(signature):

    if not signature:
        return False

    if isinstance(signature, str):

        return (
            len(signature.strip())
            > 30
        )

    return True


def signature_to_bytes(signature):

    if not signature:
        return None

    try:

        if isinstance(signature, str):

            if "," in signature:

                signature = (
                    signature.split(
                        ",",
                        1,
                    )[1]
                )

            return base64.b64decode(
                signature
            )

    except Exception:

        return None

    return None


def make_signature_image(
    signature,
    width=34 * mm,
    height=13 * mm,
):

    data = signature_to_bytes(
        signature
    )

    if not data:
        return None

    image = Image(
        io.BytesIO(data),
        width=width,
        height=height,
    )

    image.hAlign = "CENTER"

    return image


# =========================================================
# PAGE NUMBER
# =========================================================

class NumberedCanvas(canvas.Canvas):

    def __init__(
        self,
        *args,
        **kwargs,
    ):

        canvas.Canvas.__init__(
            self,
            *args,
            **kwargs,
        )

        self._saved_page_states = []

    def showPage(self):

        self._saved_page_states.append(
            dict(self.__dict__)
        )

        self._startPage()

    def save(self):

        total_pages = len(
            self._saved_page_states
        )

        for state in (
            self._saved_page_states
        ):

            self.__dict__.update(
                state
            )

            self.draw_page_footer(
                total_pages
            )

            canvas.Canvas.showPage(
                self
            )

        canvas.Canvas.save(self)

    def draw_page_footer(
        self,
        total_pages,
    ):

        page_number = (
            self._pageNumber
        )

        self.saveState()

        # Garis footer
        self.setStrokeColor(
            colors.HexColor(
                "#D0D0D0"
            )
        )

        self.setLineWidth(0.3)

        self.line(
            MARGIN_LEFT,
            6 * mm,
            PAGE_WIDTH
            - MARGIN_RIGHT,
            6 * mm,
        )

        self.setFont(
            "Helvetica",
            5.5,
        )

        self.setFillColor(
            colors.HexColor(
                "#666666"
            )
        )

        self.drawString(
            MARGIN_LEFT,
            3.5 * mm,
            "BAST - Berita Acara Serah Terima",
        )

        self.drawRightString(
            PAGE_WIDTH
            - MARGIN_RIGHT,
            3.5 * mm,
            f"Halaman {page_number} / {total_pages}",
        )

        self.restoreState()


# =========================================================
# PDF STYLES
# =========================================================

styles = getSampleStyleSheet()


STYLE_TITLE = ParagraphStyle(
    "BASTTitle",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=10.5,
    leading=11,
    alignment=TA_CENTER,
    textColor=colors.HexColor(
        "#1F2937"
    ),
)


STYLE_SUBTITLE = ParagraphStyle(
    "BASTSubtitle",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=5.5,
    leading=6,
    alignment=TA_CENTER,
    textColor=colors.HexColor(
        "#6B7280"
    ),
)


STYLE_INFO_LABEL = ParagraphStyle(
    "InfoLabel",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=5.2,
    leading=5.8,
    alignment=TA_LEFT,
    textColor=colors.HexColor(
        "#4B5563"
    ),
)


STYLE_INFO_VALUE = ParagraphStyle(
    "InfoValue",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=5.5,
    leading=6,
    alignment=TA_LEFT,
    textColor=colors.HexColor(
        "#111827"
    ),
)


STYLE_STAT_LABEL = ParagraphStyle(
    "StatLabel",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=5,
    leading=5.5,
    alignment=TA_CENTER,
    textColor=colors.HexColor(
        "#6B7280"
    ),
)


STYLE_STAT_VALUE = ParagraphStyle(
    "StatValue",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=9,
    leading=9.5,
    alignment=TA_CENTER,
    textColor=colors.HexColor(
        "#111827"
    ),
)


STYLE_TABLE_HEADER = ParagraphStyle(
    "TableHeader",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=4.5,
    leading=4.8,
    alignment=TA_CENTER,
    textColor=colors.white,
)


STYLE_TABLE = ParagraphStyle(
    "TableBody",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=4.35,
    leading=4.55,
    alignment=TA_LEFT,
    textColor=colors.HexColor(
        "#111827"
    ),
)


STYLE_TABLE_CENTER = ParagraphStyle(
    "TableBodyCenter",
    parent=STYLE_TABLE,
    alignment=TA_CENTER,
)


STYLE_SIGNATURE_TITLE = ParagraphStyle(
    "SignatureTitle",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=5.4,
    leading=6,
    alignment=TA_CENTER,
    textColor=colors.HexColor(
        "#374151"
    ),
)


STYLE_SIGNATURE_NAME = ParagraphStyle(
    "SignatureName",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=5.7,
    leading=6.2,
    alignment=TA_CENTER,
    textColor=colors.HexColor(
        "#111827"
    ),
)


STYLE_SIGNATURE_ROLE = ParagraphStyle(
    "SignatureRole",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=4.8,
    leading=5.2,
    alignment=TA_CENTER,
    textColor=colors.HexColor(
        "#6B7280"
    ),
)


# =========================================================
# HEADER
# =========================================================

def create_header(
    tanggal,
    warehouse,
    courier,
    waktu,
    driver,
    police,
    total_resi,
    total_koli,
):

    # -----------------------------------------------------
    # TITLE
    # -----------------------------------------------------

    title = Paragraph(
        "BERITA ACARA SERAH TERIMA",
        STYLE_TITLE,
    )

    subtitle = Paragraph(
        "Dokumen Serah Terima Pengiriman",
        STYLE_SUBTITLE,
    )

    # -----------------------------------------------------
    # INFORMATION
    # -----------------------------------------------------

    def info_row(
        label,
        value,
    ):

        return [
            Paragraph(
                label,
                STYLE_INFO_LABEL,
            ),
            Paragraph(
                html.escape(
                    str(value)
                ),
                STYLE_INFO_VALUE,
            ),
        ]

    info_left = Table(
        [
            info_row(
                "Tanggal",
                tanggal,
            ),
            info_row(
                "Warehouse",
                warehouse,
            ),
            info_row(
                "Courier",
                courier,
            ),
        ],
        colWidths=[
            20 * mm,
            42 * mm,
        ],
    )

    info_middle = Table(
        [
            info_row(
                "Waktu",
                waktu,
            ),
            info_row(
                "Driver",
                driver,
            ),
            info_row(
                "Police",
                police,
            ),
        ],
        colWidths=[
            18 * mm,
            44 * mm,
        ],
    )

    # -----------------------------------------------------
    # STATISTICS
    # -----------------------------------------------------

    stat_resi = Table(
        [
            [
                Paragraph(
                    "JUMLAH RESI",
                    STYLE_STAT_LABEL,
                )
            ],
            [
                Paragraph(
                    str(total_resi),
                    STYLE_STAT_VALUE,
                )
            ],
        ],
        colWidths=[
            27 * mm
        ],
        rowHeights=[
            5 * mm,
            7 * mm,
        ],
    )

    stat_koli = Table(
        [
            [
                Paragraph(
                    "JUMLAH KOLI",
                    STYLE_STAT_LABEL,
                )
            ],
            [
                Paragraph(
                    str(total_koli),
                    STYLE_STAT_VALUE,
                )
            ],
        ],
        colWidths=[
            27 * mm
        ],
        rowHeights=[
            5 * mm,
            7 * mm,
        ],
    )

    stats = Table(
        [
            [
                stat_resi,
                stat_koli,
            ]
        ],
        colWidths=[
            30 * mm,
            30 * mm,
        ],
    )

    stats.setStyle(
        TableStyle(
            [
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    1,
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    1,
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    1,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    1,
                ),
            ]
        )
    )

    # -----------------------------------------------------
    # HEADER CONTAINER
    # -----------------------------------------------------

    header_info = Table(
        [
            [
                info_left,
                info_middle,
                stats,
            ]
        ],
        colWidths=[
            65 * mm,
            65 * mm,
            65 * mm,
        ],
    )

    header_info.setStyle(
        TableStyle(
            [
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.6,
                    colors.HexColor(
                        "#CBD5E1"
                    ),
                ),
                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.25,
                    colors.HexColor(
                        "#E5E7EB"
                    ),
                ),
                (
                    "BACKGROUND",
                    (2, 0),
                    (2, 0),
                    colors.HexColor(
                        "#F8FAFC"
                    ),
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    2,
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    2,
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    1.5,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    1.5,
                ),
            ]
        )
    )

    return [
        title,
        subtitle,
        Spacer(1, 1.5 * mm),
        header_info,
        Spacer(1, 1.8 * mm),
    ]


# =========================================================
# SINGLE DATA TABLE
# =========================================================

def create_single_data_table(
    df_part,
    start_number=1,
):

    data = []

    # -----------------------------------------------------
    # HEADER
    # -----------------------------------------------------

    data.append(
        [
            Paragraph(
                "NO",
                STYLE_TABLE_HEADER,
            ),
            Paragraph(
                "DELIVERY<br/>ORDER",
                STYLE_TABLE_HEADER,
            ),
            Paragraph(
                "AIRWAYBILL",
                STYLE_TABLE_HEADER,
            ),
            Paragraph(
                "PROVIDER",
                STYLE_TABLE_HEADER,
            ),
            Paragraph(
                "KOLI<br/>QTY",
                STYLE_TABLE_HEADER,
            ),
        ]
    )

    # -----------------------------------------------------
    # 50 ROWS
    # -----------------------------------------------------

    for i in range(50):

        if i < len(df_part):

            row = df_part.iloc[i]

            no = html.escape(
                str(row["NO"])
            )

            delivery_order = html.escape(
                str(
                    row[
                        "DELIVERY ORDER"
                    ]
                )
            )

            airwaybill = html.escape(
                str(
                    row[
                        "AIRWAYBILL"
                    ]
                )
            )

            provider = html.escape(
                str(
                    row[
                        "PROVIDER"
                    ]
                )
            )

            koli = html.escape(
                str(
                    row[
                        "KOLI QTY"
                    ]
                )
            )

        else:

            no = ""
            delivery_order = ""
            airwaybill = ""
            provider = ""
            koli = ""

        data.append(
            [
                Paragraph(
                    no,
                    STYLE_TABLE_CENTER,
                ),
                Paragraph(
                    delivery_order,
                    STYLE_TABLE,
                ),
                Paragraph(
                    airwaybill,
                    STYLE_TABLE,
                ),
                Paragraph(
                    provider,
                    STYLE_TABLE,
                ),
                Paragraph(
                    koli,
                    STYLE_TABLE_CENTER,
                ),
            ]
        )

    # -----------------------------------------------------
    # WIDTH
    # -----------------------------------------------------

    col_widths = [
        7 * mm,
        25 * mm,
        32 * mm,
        23 * mm,
        9 * mm,
    ]

    # 1 header + 50 rows
    row_heights = (
        [5.5 * mm]
        + [4.25 * mm] * 50
    )

    table = Table(
        data,
        colWidths=col_widths,
        rowHeights=row_heights,
        repeatRows=1,
    )

    # -----------------------------------------------------
    # STYLE
    # -----------------------------------------------------

    table_style = [
        # Outer border
        (
            "BOX",
            (0, 0),
            (-1, -1),
            0.6,
            colors.HexColor(
                "#64748B"
            ),
        ),

        # Grid
        (
            "GRID",
            (0, 0),
            (-1, -1),
            0.25,
            colors.HexColor(
                "#CBD5E1"
            ),
        ),

        # Header
        (
            "BACKGROUND",
            (0, 0),
            (-1, 0),
            colors.HexColor(
                "#334155"
            ),
        ),

        (
            "VALIGN",
            (0, 0),
            (-1, -1),
            "MIDDLE",
        ),

        # Alignment
        (
            "ALIGN",
            (0, 0),
            (0, -1),
            "CENTER",
        ),

        (
            "ALIGN",
            (4, 0),
            (4, -1),
            "CENTER",
        ),

        (
            "ALIGN",
            (0, 0),
            (-1, 0),
            "CENTER",
        ),

        # Padding
        (
            "LEFTPADDING",
            (0, 0),
            (-1, -1),
            0.7,
        ),

        (
            "RIGHTPADDING",
            (0, 0),
            (-1, -1),
            0.7,
        ),

        (
            "TOPPADDING",
            (0, 0),
            (-1, -1),
            0.2,
        ),

        (
            "BOTTOMPADDING",
            (0, 0),
            (-1, -1),
            0.2,
        ),
    ]

    # -----------------------------------------------------
    # ALTERNATING ROW
    # -----------------------------------------------------

    for row_index in range(
        1,
        51,
    ):

        if row_index % 2 == 0:

            table_style.append(
                (
                    "BACKGROUND",
                    (0, row_index),
                    (-1, row_index),
                    colors.HexColor(
                        "#F8FAFC"
                    ),
                )
            )

    # -----------------------------------------------------
    # BLANK ROWS
    # -----------------------------------------------------

    for row_index in range(
        len(df_part) + 1,
        51,
    ):

        table_style.append(
            (
                "TEXTCOLOR",
                (0, row_index),
                (-1, row_index),
                colors.HexColor(
                    "#CBD5E1"
                ),
            )
        )

    table.setStyle(
        TableStyle(
            table_style
        )
    )

    return table


# =========================================================
# SIDE BY SIDE TABLE
# =========================================================

def create_side_by_side_table(
    left_df,
    right_df,
):

    left_table = (
        create_single_data_table(
            left_df,
        )
    )

    right_table = (
        create_single_data_table(
            right_df,
        )
    )

    # Label kiri / kanan
    left_title = Paragraph(
        f"RESI {left_df.iloc[0]['NO'] if len(left_df) else ''}"
        f" - "
        f"{left_df.iloc[-1]['NO'] if len(left_df) else ''}",
        STYLE_TABLE_HEADER,
    )

    right_title = Paragraph(
        f"RESI {right_df.iloc[0]['NO'] if len(right_df) else ''}"
        f" - "
        f"{right_df.iloc[-1]['NO'] if len(right_df) else ''}",
        STYLE_TABLE_HEADER,
    )

    # Jangan pakai title terpisah supaya tinggi tidak bertambah.
    # Table langsung ditempatkan berdampingan.

    outer = Table(
        [
            [
                left_table,
                right_table,
            ]
        ],
        colWidths=[
            97 * mm,
            97 * mm,
        ],
    )

    outer.setStyle(
        TableStyle(
            [
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "TOP",
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),

                # Separator tengah
                (
                    "LINEAFTER",
                    (0, 0),
                    (0, 0),
                    0.7,
                    colors.HexColor(
                        "#475569"
                    ),
                ),
            ]
        )
    )

    return outer


# =========================================================
# SIGNATURE
# =========================================================

def create_signature_table(
    security_signature,
    dispatcher_signature,
    driver_signature,

    security_name,
    dispatcher_name,
    driver_name,
):

    security_image = (
        make_signature_image(
            security_signature,
            width=31 * mm,
            height=12 * mm,
        )
    )

    dispatcher_image = (
        make_signature_image(
            dispatcher_signature,
            width=31 * mm,
            height=12 * mm,
        )
    )

    driver_image = (
        make_signature_image(
            driver_signature,
            width=31 * mm,
            height=12 * mm,
        )
    )

    # -----------------------------------------------------
    # IMAGE / EMPTY
    # -----------------------------------------------------

    def image_or_blank(image):

        if image is None:

            return Paragraph(
                "<br/><br/>",
                STYLE_SIGNATURE_NAME,
            )

        return image

    # -----------------------------------------------------
    # HEADER
    # -----------------------------------------------------

    header = [
        [
            Paragraph(
                "DIPERIKSA OLEH",
                STYLE_SIGNATURE_TITLE,
            ),
            Paragraph(
                "DISERAHKAN OLEH",
                STYLE_SIGNATURE_TITLE,
            ),
            Paragraph(
                "DITERIMA OLEH",
                STYLE_SIGNATURE_TITLE,
            ),
        ],
        [
            Paragraph(
                "Security WH",
                STYLE_SIGNATURE_ROLE,
            ),
            Paragraph(
                "Dispatcher WH",
                STYLE_SIGNATURE_ROLE,
            ),
            Paragraph(
                "Driver Courier",
                STYLE_SIGNATURE_ROLE,
            ),
        ],
        [
            image_or_blank(
                security_image
            ),
            image_or_blank(
                dispatcher_image
            ),
            image_or_blank(
                driver_image
            ),
        ],
        [
            Paragraph(
                f"({html.escape(str(security_name))})",
                STYLE_SIGNATURE_NAME,
            ),
            Paragraph(
                f"({html.escape(str(dispatcher_name))})",
                STYLE_SIGNATURE_NAME,
            ),
            Paragraph(
                f"({html.escape(str(driver_name))})",
                STYLE_SIGNATURE_NAME,
            ),
        ],
    ]

    table = Table(
        header,
        colWidths=[
            63 * mm,
            63 * mm,
            63 * mm,
        ],
        rowHeights=[
            5 * mm,
            4 * mm,
            13 * mm,
            5 * mm,
        ],
    )

    table.setStyle(
        TableStyle(
            [
                (
                    "BOX",
                    (0, 0),
                    (-1, -1),
                    0.6,
                    colors.HexColor(
                        "#94A3B8"
                    ),
                ),

                (
                    "INNERGRID",
                    (0, 0),
                    (-1, -1),
                    0.25,
                    colors.HexColor(
                        "#CBD5E1"
                    ),
                ),

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 1),
                    colors.HexColor(
                        "#F8FAFC"
                    ),
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),

                (
                    "ALIGN",
                    (0, 0),
                    (-1, -1),
                    "CENTER",
                ),

                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    1,
                ),

                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    1,
                ),

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    0.5,
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0.5,
                ),
            ]
        )
    )

    return table


# =========================================================
# PDF GENERATOR
# =========================================================

def generate_pdf(
    df,

    tanggal,
    warehouse,
    courier,
    waktu,
    driver,
    police,

    security_signature,
    dispatcher_signature,
    driver_signature,

    security_name,
    dispatcher_name,
    driver_name,
):

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,

        leftMargin=MARGIN_LEFT,
        rightMargin=MARGIN_RIGHT,

        topMargin=MARGIN_TOP,
        bottomMargin=MARGIN_BOTTOM,

        title="BAST",
        author="BAST Generator",

        allowSplitting=False,
    )

    story = []

    # -----------------------------------------------------
    # TOTAL
    # -----------------------------------------------------

    total_resi = len(df)

    total_koli = int(
        pd.to_numeric(
            df["KOLI QTY"],
            errors="coerce",
        )
        .fillna(0)
        .sum()
    )

    # -----------------------------------------------------
    # CHUNK 100
    # -----------------------------------------------------

    chunks = []

    for start in range(
        0,
        len(df),
        100,
    ):

        chunk = df.iloc[
            start:start + 100
        ].copy()

        left_df = chunk.iloc[
            :50
        ].copy()

        right_df = chunk.iloc[
            50:100
        ].copy()

        chunks.append(
            (
                left_df,
                right_df,
            )
        )

    # Data kosong
    if not chunks:

        chunks.append(
            (
                pd.DataFrame(),
                pd.DataFrame(),
            )
        )

    # -----------------------------------------------------
    # EACH PAGE
    # -----------------------------------------------------

    for page_index, (
        left_df,
        right_df,
    ) in enumerate(chunks):

        # HEADER
        story.extend(
            create_header(
                tanggal=tanggal,
                warehouse=warehouse,
                courier=courier,
                waktu=waktu,
                driver=driver,
                police=police,
                total_resi=total_resi,
                total_koli=total_koli,
            )
        )

        # TABLE
        story.append(
            create_side_by_side_table(
                left_df,
                right_df,
            )
        )

        story.append(
            Spacer(
                1,
                2.2 * mm,
            )
        )

        # SIGNATURE
        story.append(
            create_signature_table(
                security_signature,
                dispatcher_signature,
                driver_signature,

                security_name,
                dispatcher_name,
                driver_name,
            )
        )

        # PAGE BREAK
        if page_index < len(chunks) - 1:

            story.append(
                PageBreak()
            )

    # -----------------------------------------------------
    # BUILD
    # -----------------------------------------------------

    doc.build(
        story,
        canvasmaker=NumberedCanvas,
    )

    buffer.seek(0)

    return buffer.getvalue()


# =========================================================
# STREAMLIT
# =========================================================

st.title("📄 BAST Generator")

st.caption(
    "BAST compact — 100 resi per halaman "
    "(50 kiri + 50 kanan)"
)


# =========================================================
# INFORMASI
# =========================================================

st.subheader(
    "1. Informasi BAST"
)

col1, col2, col3 = st.columns(3)


with col1:

    tanggal = st.date_input(
        "Tanggal"
    )

    warehouse = st.text_input(
        "Warehouse",
        placeholder="Contoh: WH Jakarta",
    )

    courier = st.text_input(
        "Courier",
        placeholder="Contoh: JNE",
    )


with col2:

    waktu = st.time_input(
        "Waktu"
    )

    driver = st.text_input(
        "Driver",
        placeholder="Nama Driver",
    )

    police = st.text_input(
        "Police",
        placeholder="Nomor kendaraan",
    )


with col3:

    st.markdown(
        """
        **Layout PDF**

        • A4 Portrait  
        • 50 resi kiri  
        • 50 resi kanan  
        • 100 resi / halaman  
        • >100 otomatis halaman baru  
        • 3 TTD di setiap halaman
        """
    )


# =========================================================
# NAMA SIGNER
# =========================================================

st.subheader(
    "2. Nama Penanda Tangan"
)

name1, name2, name3 = st.columns(3)


with name1:

    security_name = st.text_input(
        "Nama Security WH",
        placeholder="Nama lengkap",
    )


with name2:

    dispatcher_name = st.text_input(
        "Nama Dispatcher WH",
        placeholder="Nama lengkap",
    )


with name3:

    driver_name = st.text_input(
        "Nama Driver Courier",
        placeholder="Nama lengkap",
    )


# =========================================================
# DATA
# =========================================================

st.subheader(
    "3. Data Resi"
)

st.markdown(
    """
Paste langsung dari Excel dengan urutan kolom:

**NO | DELIVERY ORDER | AIRWAYBILL | PROVIDER | KOLI QTY**
"""
)

paste_data = st.text_area(
    "Paste data Excel",
    height=220,
    placeholder=(
        "NO\tDELIVERY ORDER\tAIRWAYBILL\tPROVIDER\tKOLI QTY\n"
        "1\tDO001\tAWB001\tJNE\t2\n"
        "2\tDO002\tAWB002\tJ&T\t1"
    ),
)


# =========================================================
# PREVIEW
# =========================================================

df = pd.DataFrame()

if paste_data.strip():

    df = parse_paste_data(
        paste_data
    )

    df = fix_broken_rows(
        df
    )

    valid, message = (
        validate_file(df)
    )

    if valid:

        total_koli_preview = int(
            df["KOLI QTY"].sum()
        )

        c1, c2, c3 = st.columns(3)

        with c1:

            st.metric(
                "Jumlah Resi",
                len(df),
            )

        with c2:

            st.metric(
                "Jumlah Koli",
                total_koli_preview,
            )

        with c3:

            pages = (
                (len(df) + 99) // 100
                if len(df)
                else 0
            )

            st.metric(
                "Perkiraan Halaman",
                pages,
            )

        with st.expander(
            "Preview Data",
            expanded=False,
        ):

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True,
            )

    else:

        st.error(
            message
        )


# =========================================================
# SIGNATURE
# =========================================================

st.subheader(
    "4. Tanda Tangan"
)

st.caption(
    "Gunakan mouse, trackpad, atau jari "
    "untuk membuat tanda tangan."
)

sig1, sig2, sig3 = st.columns(3)


with sig1:

    st.markdown(
        "**Diperiksa oleh — Security WH**"
    )

    security_signature = (
        signature_pad(
            key="security_signature",
            width=300,
            height=140,
        )
    )


with sig2:

    st.markdown(
        "**Diserahkan oleh — Dispatcher WH**"
    )

    dispatcher_signature = (
        signature_pad(
            key="dispatcher_signature",
            width=300,
            height=140,
        )
    )


with sig3:

    st.markdown(
        "**Diterima oleh — Driver Courier**"
    )

    driver_signature = (
        signature_pad(
            key="driver_signature",
            width=300,
            height=140,
        )
    )


# =========================================================
# VALIDATION
# =========================================================

errors = []


if not warehouse.strip():

    errors.append(
        "Warehouse belum diisi."
    )


if not courier.strip():

    errors.append(
        "Courier belum diisi."
    )


if not driver.strip():

    errors.append(
        "Driver belum diisi."
    )


if not police.strip():

    errors.append(
        "Police / nomor kendaraan belum diisi."
    )


if not security_name.strip():

    errors.append(
        "Nama Security WH belum diisi."
    )


if not dispatcher_name.strip():

    errors.append(
        "Nama Dispatcher WH belum diisi."
    )


if not driver_name.strip():

    errors.append(
        "Nama Driver Courier belum diisi."
    )


if df.empty:

    errors.append(
        "Data resi belum diisi."
    )


if not signature_exists(
    security_signature
):

    errors.append(
        "TTD Security WH belum dibuat."
    )


if not signature_exists(
    dispatcher_signature
):

    errors.append(
        "TTD Dispatcher WH belum dibuat."
    )


if not signature_exists(
    driver_signature
):

    errors.append(
        "TTD Driver Courier belum dibuat."
    )


# =========================================================
# GENERATE
# =========================================================

st.subheader(
    "5. Generate PDF"
)


if errors:

    with st.expander(
        "Data yang masih perlu dilengkapi",
        expanded=False,
    ):

        for error in errors:

            st.warning(
                error
            )


if st.button(
    "📄 GENERATE BAST PDF",
    type="primary",
    use_container_width=True,
):

    if errors:

        st.error(
            "Mohon lengkapi data terlebih dahulu."
        )

    else:

        try:

            with st.spinner(
                "Membuat PDF BAST..."
            ):

                pdf_bytes = generate_pdf(
                    df=df,

                    tanggal=tanggal,
                    warehouse=warehouse,
                    courier=courier,
                    waktu=waktu,
                    driver=driver,
                    police=police,

                    security_signature=(
                        security_signature
                    ),

                    dispatcher_signature=(
                        dispatcher_signature
                    ),

                    driver_signature=(
                        driver_signature
                    ),

                    security_name=(
                        security_name
                    ),

                    dispatcher_name=(
                        dispatcher_name
                    ),

                    driver_name=(
                        driver_name
                    ),
                )

            filename = (
                "BAST_"
                + safe_filename(
                    warehouse
                )
                + "_"
                + str(tanggal)
                + ".pdf"
            )

            st.success(
                "PDF berhasil dibuat."
            )

            st.download_button(
                label="⬇️ DOWNLOAD BAST PDF",
                data=pdf_bytes,
                file_name=filename,
                mime="application/pdf",
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                "Gagal membuat PDF."
            )

            st.exception(e)
