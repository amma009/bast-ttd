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
# CONFIG
# =========================================================

st.set_page_config(
    page_title="BAST Generator",
    page_icon="📄",
    layout="wide",
)


# =========================================================
# IMPORT SIGNATURE COMPONENT
# =========================================================

from signature_component import signature_pad


# =========================================================
# HELPER
# =========================================================

REQUIRED_COLUMNS = [
    "NO",
    "DELIVERY ORDER",
    "AIRWAYBILL",
    "PROVIDER",
    "KOLI QTY",
]


def safe_filename(text):
    text = str(text).strip()

    if not text:
        text = "BAST"

    text = re.sub(r'[\\/*?:"<>|]', "_", text)

    return text


def parse_paste_data(text):
    """
    Parse data hasil copy dari Excel.

    Prioritas:
    1. TSV / tab
    2. CSV
    3. auto python engine
    """

    if not text or not text.strip():
        return pd.DataFrame()

    text = text.strip()

    # Coba TSV
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

    # Coba CSV
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

    # Coba python engine
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
    """
    Membersihkan data hasil paste Excel.
    """

    if df.empty:
        return df

    df = df.copy()

    # Normalisasi nama kolom
    df.columns = [
        str(col).strip().upper()
        for col in df.columns
    ]

    # Buang kolom Unnamed
    df = df.loc[
        :,
        ~df.columns.str.contains(
            "^UNNAMED",
            case=False,
            regex=True,
        )
    ]

    # Pastikan kolom wajib ada
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    # Hanya ambil kolom yang diperlukan
    df = df[REQUIRED_COLUMNS].copy()

    # Bersihkan whitespace
    for col in df.columns:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("\n", " ", regex=False)
            .str.replace("\r", " ", regex=False)
            .str.strip()
        )

    # NO jika kosong → isi berdasarkan urutan
    for idx in df.index:
        if not str(df.at[idx, "NO"]).strip():
            df.at[idx, "NO"] = str(idx + 1)

    # KOLI QTY numeric
    df["KOLI QTY"] = pd.to_numeric(
        df["KOLI QTY"],
        errors="coerce",
    ).fillna(0)

    df["KOLI QTY"] = df["KOLI QTY"].astype(int)

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
# SIGNATURE
# =========================================================

def signature_exists(signature):
    """
    Mengecek apakah signature dari canvas ada.
    """

    if not signature:
        return False

    if isinstance(signature, str):
        return len(signature.strip()) > 30

    return True


def signature_to_bytes(signature):
    """
    Convert base64 signature menjadi bytes.
    """

    if not signature:
        return None

    try:
        if isinstance(signature, str):

            # format:
            # data:image/png;base64,xxxx
            if "," in signature:
                signature = signature.split(",", 1)[1]

            return base64.b64decode(signature)

    except Exception:
        return None

    return None


def make_signature_image(signature, width=35 * mm, height=15 * mm):
    """
    Membuat ReportLab Image dari signature canvas.
    """

    data = signature_to_bytes(signature)

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
# PAGE NUMBER CANVAS
# =========================================================

class NumberedCanvas(canvas.Canvas):

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)

        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))

        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)

        for state in self._saved_page_states:

            self.__dict__.update(state)

            self.draw_page_number(num_pages)

            canvas.Canvas.showPage(self)

        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):

        page_number = self._pageNumber

        self.saveState()

        self.setFont("Helvetica", 6.5)

        self.setFillColor(colors.grey)

        text = f"Page {page_number} / {page_count}"

        self.drawCentredString(
            A4[0] / 2,
            4 * mm,
            text,
        )

        self.restoreState()


# =========================================================
# REPORTLAB STYLES
# =========================================================

styles = getSampleStyleSheet()


STYLE_HEADER = ParagraphStyle(
    "Header",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=9,
    leading=10,
    alignment=TA_CENTER,
)


STYLE_HEADER_SMALL = ParagraphStyle(
    "HeaderSmall",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=6.5,
    leading=7.5,
    alignment=TA_LEFT,
)


STYLE_TABLE_HEADER = ParagraphStyle(
    "TableHeader",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=4.6,
    leading=5,
    alignment=TA_CENTER,
)


STYLE_TABLE = ParagraphStyle(
    "Table",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=4.5,
    leading=4.8,
    alignment=TA_LEFT,
)


STYLE_TABLE_CENTER = ParagraphStyle(
    "TableCenter",
    parent=STYLE_TABLE,
    alignment=TA_CENTER,
)


STYLE_SIGNATURE = ParagraphStyle(
    "Signature",
    parent=styles["Normal"],
    fontName="Helvetica",
    fontSize=5.5,
    leading=6.5,
    alignment=TA_CENTER,
)


STYLE_SIGNATURE_NAME = ParagraphStyle(
    "SignatureName",
    parent=styles["Normal"],
    fontName="Helvetica-Bold",
    fontSize=6,
    leading=7,
    alignment=TA_CENTER,
)


# =========================================================
# HEADER PDF
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
    """
    Header compact.
    """

    title = Paragraph(
        "BERITA ACARA SERAH TERIMA (BAST)",
        STYLE_HEADER,
    )

    data_left = [
        [
            Paragraph("<b>Tanggal</b>", STYLE_HEADER_SMALL),
            Paragraph(str(tanggal), STYLE_HEADER_SMALL),
        ],
        [
            Paragraph("<b>Warehouse</b>", STYLE_HEADER_SMALL),
            Paragraph(str(warehouse), STYLE_HEADER_SMALL),
        ],
        [
            Paragraph("<b>Courier</b>", STYLE_HEADER_SMALL),
            Paragraph(str(courier), STYLE_HEADER_SMALL),
        ],
    ]

    data_middle = [
        [
            Paragraph("<b>Waktu</b>", STYLE_HEADER_SMALL),
            Paragraph(str(waktu), STYLE_HEADER_SMALL),
        ],
        [
            Paragraph("<b>Driver</b>", STYLE_HEADER_SMALL),
            Paragraph(str(driver), STYLE_HEADER_SMALL),
        ],
        [
            Paragraph("<b>Police</b>", STYLE_HEADER_SMALL),
            Paragraph(str(police), STYLE_HEADER_SMALL),
        ],
    ]

    data_right = [
        [
            Paragraph("<b>Jumlah Resi</b>", STYLE_HEADER_SMALL),
            Paragraph(
                f"<b>{total_resi}</b>",
                STYLE_HEADER_SMALL,
            ),
        ],
        [
            Paragraph("<b>Jumlah Koli</b>", STYLE_HEADER_SMALL),
            Paragraph(
                f"<b>{total_koli}</b>",
                STYLE_HEADER_SMALL,
            ),
        ],
        [
            Paragraph("<b>Status</b>", STYLE_HEADER_SMALL),
            Paragraph(
                "READY",
                STYLE_HEADER_SMALL,
            ),
        ],
    ]

    def small_info_table(data):

        table = Table(
            data,
            colWidths=[
                25 * mm,
                35 * mm,
            ],
        )

        table.setStyle(
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
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.25,
                        colors.grey,
                    ),
                ]
            )
        )

        return table

    info_table = Table(
        [
            [
                small_info_table(data_left),
                small_info_table(data_middle),
                small_info_table(data_right),
            ]
        ],
        colWidths=[
            65 * mm,
            65 * mm,
            65 * mm,
        ],
    )

    info_table.setStyle(
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
                    0,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),
            ]
        )
    )

    return [
        title,
        Spacer(1, 1.5 * mm),
        info_table,
        Spacer(1, 2 * mm),
    ]


# =========================================================
# DATA TABLE
# =========================================================

def create_single_data_table(df_part):
    """
    Membuat SATU tabel 50 baris.
    """

    table_data = []

    # Header
    table_data.append(
        [
            Paragraph("NO", STYLE_TABLE_HEADER),
            Paragraph("DELIVERY<br/>ORDER", STYLE_TABLE_HEADER),
            Paragraph("AIRWAYBILL", STYLE_TABLE_HEADER),
            Paragraph("PROVIDER", STYLE_TABLE_HEADER),
            Paragraph("KOLI<br/>QTY", STYLE_TABLE_HEADER),
        ]
    )

    # 50 rows
    for i in range(50):

        if i < len(df_part):

            row = df_part.iloc[i]

            no = html.escape(
                str(row["NO"])
            )

            delivery_order = html.escape(
                str(row["DELIVERY ORDER"])
            )

            airwaybill = html.escape(
                str(row["AIRWAYBILL"])
            )

            provider = html.escape(
                str(row["PROVIDER"])
            )

            koli = html.escape(
                str(row["KOLI QTY"])
            )

        else:

            no = ""
            delivery_order = ""
            airwaybill = ""
            provider = ""
            koli = ""

        table_data.append(
            [
                Paragraph(no, STYLE_TABLE_CENTER),
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

    # Lebar total sekitar 96 mm
    col_widths = [
        7 * mm,    # NO
        25 * mm,   # DO
        32 * mm,   # AWB
        23 * mm,   # Provider
        9 * mm,    # Koli
    ]

    # Header + 50 rows
    row_heights = [
        6 * mm
    ] + [
        4.25 * mm
    ] * 50

    table = Table(
        table_data,
        colWidths=col_widths,
        rowHeights=row_heights,
    )

    table.setStyle(
        TableStyle(
            [
                # GRID
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.25,
                    colors.grey,
                ),

                # Header
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#EDEDED"),
                ),

                (
                    "ALIGN",
                    (0, 0),
                    (-1, 0),
                    "CENTER",
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),

                # Padding compact
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    0.8,
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    0.8,
                ),
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    0.3,
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0.3,
                ),

                # Center columns
                (
                    "ALIGN",
                    (0, 1),
                    (0, -1),
                    "CENTER",
                ),
                (
                    "ALIGN",
                    (4, 1),
                    (4, -1),
                    "CENTER",
                ),
            ]
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
    """
    Membuat:

    ┌───────────────┬───────────────┐
    │   1 - 50      │   51 - 100    │
    │   LEFT        │   RIGHT       │
    └───────────────┴───────────────┘
    """

    left_table = create_single_data_table(
        left_df
    )

    right_table = create_single_data_table(
        right_df
    )

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

                # garis pemisah tengah
                (
                    "LINEAFTER",
                    (0, 0),
                    (0, 0),
                    0.5,
                    colors.black,
                ),
            ]
        )
    )

    return outer


# =========================================================
# SIGNATURE TABLE
# =========================================================

def create_signature_table(
    security_signature,
    dispatcher_signature,
    driver_signature,
    security_name,
    dispatcher_name,
    driver_name,
):
    """
    3 signature berdampingan.
    """

    img_security = make_signature_image(
        security_signature,
        width=32 * mm,
        height=12 * mm,
    )

    img_dispatcher = make_signature_image(
        dispatcher_signature,
        width=32 * mm,
        height=12 * mm,
    )

    img_driver = make_signature_image(
        driver_signature,
        width=32 * mm,
        height=12 * mm,
    )

    def signature_content(
        title,
        image,
        name,
    ):

        if image is None:

            image_cell = Paragraph(
                "<br/><br/>",
                STYLE_SIGNATURE,
            )

        else:

            image_cell = image

        return [
            Paragraph(
                title,
                STYLE_SIGNATURE,
            ),
            image_cell,
            Paragraph(
                f"({html.escape(str(name))})",
                STYLE_SIGNATURE_NAME,
            ),
        ]

    security = signature_content(
        "Diperiksa oleh<br/>Security WH",
        img_security,
        security_name,
    )

    dispatcher = signature_content(
        "Diserahkan oleh<br/>Dispatcher WH",
        img_dispatcher,
        dispatcher_name,
    )

    driver = signature_content(
        "Diterima oleh<br/>Driver Courier",
        img_driver,
        driver_name,
    )

    data = [
        [
            security[0],
            dispatcher[0],
            driver[0],
        ],
        [
            security[1],
            dispatcher[1],
            driver[1],
        ],
        [
            security[2],
            dispatcher[2],
            driver[2],
        ],
    ]

    table = Table(
        data,
        colWidths=[
            63 * mm,
            63 * mm,
            63 * mm,
        ],
        rowHeights=[
            7 * mm,
            12 * mm,
            5 * mm,
        ],
    )

    table.setStyle(
        TableStyle(
            [
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
                    0,
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0,
                ),

                (
                    "LINEABOVE",
                    (0, 2),
                    (-1, 2),
                    0.5,
                    colors.black,
                ),
            ]
        )
    )

    return table


# =========================================================
# GENERATE PDF
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
    """
    Generate PDF A4 portrait.

    Setiap halaman:
    - 50 kiri
    - 50 kanan
    - total 100 resi
    """

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,

        leftMargin=6 * mm,
        rightMargin=6 * mm,

        topMargin=5 * mm,
        bottomMargin=8 * mm,

        title="BAST",
        author="BAST Generator",
    )

    story = []

    total_resi = len(df)

    total_koli = int(
        pd.to_numeric(
            df["KOLI QTY"],
            errors="coerce",
        )
        .fillna(0)
        .sum()
    )

    # =====================================================
    # CHUNK 100 RESI / PAGE
    # =====================================================

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

    # Kalau data kosong
    if not chunks:
        chunks.append(
            (
                pd.DataFrame(),
                pd.DataFrame(),
            )
        )

    # =====================================================
    # BUILD EACH PAGE
    # =====================================================

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

        # DATA TABLE
        story.append(
            create_side_by_side_table(
                left_df,
                right_df,
            )
        )

        story.append(
            Spacer(
                1,
                2.5 * mm,
            )
        )

        # SIGNATURE
        story.append(
            create_signature_table(
                security_signature=security_signature,
                dispatcher_signature=dispatcher_signature,
                driver_signature=driver_signature,

                security_name=security_name,
                dispatcher_name=dispatcher_name,
                driver_name=driver_name,
            )
        )

        # PAGE BREAK
        if page_index < len(chunks) - 1:

            story.append(
                PageBreak()
            )

    # =====================================================
    # BUILD
    # =====================================================

    doc.build(
        story,
        canvasmaker=NumberedCanvas,
    )

    buffer.seek(0)

    return buffer.getvalue()


# =========================================================
# STREAMLIT UI
# =========================================================

st.title("📄 BAST Generator")

st.caption(
    "1 halaman maksimal 100 resi — "
    "50 tabel kiri + 50 tabel kanan."
)


# =========================================================
# HEADER INPUT
# =========================================================

st.subheader("1. Informasi BAST")

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
        placeholder="Nama driver",
    )

    police = st.text_input(
        "Police",
        placeholder="Nomor kendaraan",
    )


with col3:

    st.info(
        """
        **Layout PDF**

        • A4 Portrait  
        • 50 resi kiri  
        • 50 resi kanan  
        • Maksimal 100 resi/page  
        • >100 otomatis halaman berikutnya
        """
    )


# =========================================================
# SIGNER NAME
# =========================================================

st.subheader("2. Nama Penanda Tangan")

name_col1, name_col2, name_col3 = st.columns(3)

with name_col1:

    security_name = st.text_input(
        "Nama Security WH",
        placeholder="Nama lengkap",
    )

with name_col2:

    dispatcher_name = st.text_input(
        "Nama Dispatcher WH",
        placeholder="Nama lengkap",
    )

with name_col3:

    driver_name = st.text_input(
        "Nama Driver Courier",
        placeholder="Nama lengkap",
    )


# =========================================================
# DATA EXCEL
# =========================================================

st.subheader("3. Data Resi")

st.markdown(
    """
Copy data dari Excel lalu paste di bawah.

Kolom yang digunakan:

`NO | DELIVERY ORDER | AIRWAYBILL | PROVIDER | KOLI QTY`
"""
)

paste_data = st.text_area(
    "Paste data Excel di sini",
    height=220,
    placeholder=(
        "NO\tDELIVERY ORDER\tAIRWAYBILL\tPROVIDER\tKOLI QTY\n"
        "1\tDO001\tAWB001\tJNE\t2\n"
        "2\tDO002\tAWB002\tJ&T\t1"
    ),
)


# =========================================================
# PREVIEW DATA
# =========================================================

df = pd.DataFrame()

if paste_data.strip():

    df = parse_paste_data(
        paste_data
    )

    df = fix_broken_rows(
        df
    )

    valid, message = validate_file(
        df
    )

    if valid:

        st.success(
            f"Data berhasil dibaca: {len(df)} resi"
        )

        preview_col1, preview_col2 = st.columns(2)

        with preview_col1:

            st.metric(
                "Jumlah Resi",
                len(df),
            )

        with preview_col2:

            total_koli_preview = int(
                df["KOLI QTY"].sum()
            )

            st.metric(
                "Jumlah Koli",
                total_koli_preview,
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

st.subheader("4. Tanda Tangan")

st.caption(
    "Tanda tangan langsung menggunakan mouse, "
    "trackpad, atau jari pada layar."
)

sig_col1, sig_col2, sig_col3 = st.columns(3)


with sig_col1:

    st.markdown(
        "**Diperiksa oleh — Security WH**"
    )

    security_signature = signature_pad(
        key="security_signature",
        width=300,
        height=140,
    )


with sig_col2:

    st.markdown(
        "**Diserahkan oleh — Dispatcher WH**"
    )

    dispatcher_signature = signature_pad(
        key="dispatcher_signature",
        width=300,
        height=140,
    )


with sig_col3:

    st.markdown(
        "**Diterima oleh — Driver Courier**"
    )

    driver_signature = signature_pad(
        key="driver_signature",
        width=300,
        height=140,
    )


# =========================================================
# VALIDATION
# =========================================================

st.subheader("5. Generate PDF")


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
        "Tanda tangan Security WH belum dibuat."
    )

if not signature_exists(
    dispatcher_signature
):
    errors.append(
        "Tanda tangan Dispatcher WH belum dibuat."
    )

if not signature_exists(
    driver_signature
):
    errors.append(
        "Tanda tangan Driver Courier belum dibuat."
    )


if errors:

    with st.expander(
        "Yang masih perlu dilengkapi",
        expanded=False,
    ):

        for error in errors:

            st.warning(
                error
            )


# =========================================================
# GENERATE BUTTON
# =========================================================

if st.button(
    "📄 Generate BAST PDF",
    type="primary",
    use_container_width=True,
):

    if errors:

        st.error(
            "Mohon lengkapi data terlebih dahulu."
        )

    else:

        with st.spinner(
            "Membuat PDF..."
        ):

            try:

                pdf_bytes = generate_pdf(
                    df=df,

                    tanggal=tanggal,
                    warehouse=warehouse,
                    courier=courier,
                    waktu=waktu,
                    driver=driver,
                    police=police,

                    security_signature=security_signature,
                    dispatcher_signature=dispatcher_signature,
                    driver_signature=driver_signature,

                    security_name=security_name,
                    dispatcher_name=dispatcher_name,
                    driver_name=driver_name,
                )

                filename = (
                    f"BAST_"
                    f"{safe_filename(warehouse)}_"
                    f"{tanggal}.pdf"
                )

                st.success(
                    "PDF berhasil dibuat."
                )

                st.download_button(
                    label="⬇️ Download BAST PDF",
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
