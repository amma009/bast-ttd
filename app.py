import streamlit as st
import pandas as pd
import base64
import io
import re
import math

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    KeepTogether,
)
from reportlab.pdfgen import canvas

from signature_component import signature_pad


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="BAST Generator",
    layout="wide"
)

st.title("📦 BAST Generator")
st.caption(
    "Generate Berita Acara Serah Terima dengan tanda tangan digital."
)


# =========================================================
# HELPER
# =========================================================

def safe_filename(text):
    if not text:
        return "BAST"

    text = str(text)

    text = re.sub(
        r'[\\/*?:"<>|]',
        "_",
        text
    )

    text = re.sub(
        r"\s+",
        "_",
        text
    )

    return text


def parse_paste_data(text):
    """
    Membaca data hasil copy dari Excel.
    """

    if not text or not text.strip():
        return None

    lines = [
        line
        for line in text.strip().splitlines()
        if line.strip()
    ]

    if not lines:
        return None

    first_line = lines[0]

    if "\t" in first_line:
        separator = "\t"
    elif ";" in first_line:
        separator = ";"
    elif "," in first_line:
        separator = ","
    else:
        separator = "\t"

    try:

        df = pd.read_csv(
            io.StringIO(
                "\n".join(lines)
            ),
            sep=separator,
            dtype=str
        )

        df.columns = [
            str(col).strip().upper()
            for col in df.columns
        ]

        df = df.fillna("")

        return df

    except Exception:

        return None


def fix_broken_rows(df):

    if df is None or df.empty:
        return df

    df = df.copy()

    for col in df.columns:

        df[col] = (
            df[col]
            .astype(str)
            .replace(
                "nan",
                "",
                regex=False
            )
            .str.strip()
        )

    return df


def validate_file(df):

    required_columns = [
        "NO",
        "DELIVERY ORDER",
        "AIRWAYBILL",
        "PROVIDER",
        "KOLI QTY",
    ]

    if df is None:
        return False, "Data kosong."

    missing = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing:

        return False, (
            "Kolom berikut belum ditemukan: "
            + ", ".join(missing)
        )

    return True, ""


def signature_exists(signature):

    return (
        isinstance(signature, str)
        and signature.startswith(
            "data:image/png;base64,"
        )
    )


def signature_to_bytes(signature):

    if not signature_exists(signature):
        return None

    try:

        encoded = signature.split(
            ",",
            1
        )[1]

        image_bytes = base64.b64decode(
            encoded
        )

        return io.BytesIO(
            image_bytes
        )

    except Exception:

        return None


# =========================================================
# NUMBERED CANVAS
# =========================================================

class NumberedCanvas(canvas.Canvas):

    def __init__(
        self,
        *args,
        **kwargs
    ):

        canvas.Canvas.__init__(
            self,
            *args,
            **kwargs
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

        for state in self._saved_page_states:

            self.__dict__.update(state)

            self.draw_page_number(
                total_pages
            )

            canvas.Canvas.showPage(self)

        canvas.Canvas.save(self)

    def draw_page_number(
        self,
        page_count
    ):

        page_number = self._pageNumber

        self.saveState()

        self.setFont(
            "Helvetica",
            7
        )

        self.drawRightString(
            A4[0] - 10 * mm,
            6 * mm,
            f"Page {page_number} / {page_count}"
        )

        self.restoreState()


# =========================================================
# SIGNATURE IMAGE
# =========================================================

def make_signature_image(
    image_bytes
):

    if not image_bytes:

        return Spacer(
            1,
            18 * mm
        )

    try:

        image_bytes.seek(0)

        return RLImage(
            image_bytes,
            width=35 * mm,
            height=18 * mm,
        )

    except Exception:

        return Spacer(
            1,
            18 * mm
        )


# =========================================================
# SIGNATURE BLOCK
# =========================================================

def create_signature_table(
    security_signature,
    dispatcher_signature,
    driver_signature,
    security_name,
    dispatcher_name,
    driver_name,
):

    security_bytes = signature_to_bytes(
        security_signature
    )

    dispatcher_bytes = signature_to_bytes(
        dispatcher_signature
    )

    driver_bytes = signature_to_bytes(
        driver_signature
    )

    security_img = make_signature_image(
        security_bytes
    )

    dispatcher_img = make_signature_image(
        dispatcher_bytes
    )

    driver_img = make_signature_image(
        driver_bytes
    )

    signature_role_style = ParagraphStyle(
        "SignatureRole",
        fontName="Helvetica",
        fontSize=7,
        leading=8,
        alignment=TA_CENTER,
    )

    signature_name_style = ParagraphStyle(
        "SignatureName",
        fontName="Helvetica-Bold",
        fontSize=7,
        leading=8,
        alignment=TA_CENTER,
    )

    data = [

        [
            Paragraph(
                "<b>Diperiksa oleh</b><br/>"
                "Security WH",
                signature_role_style
            ),

            Paragraph(
                "<b>Diserahkan oleh</b><br/>"
                "Dispatcher WH",
                signature_role_style
            ),

            Paragraph(
                "<b>Diterima oleh</b><br/>"
                "Driver Courier",
                signature_role_style
            ),
        ],

        [
            security_img,
            dispatcher_img,
            driver_img,
        ],

        [
            Paragraph(
                f"<u>{security_name}</u>",
                signature_name_style
            ),

            Paragraph(
                f"<u>{dispatcher_name}</u>",
                signature_name_style
            ),

            Paragraph(
                f"<u>{driver_name}</u>",
                signature_name_style
            ),
        ],
    ]

    table = Table(
        data,
        colWidths=[
            58 * mm,
            58 * mm,
            58 * mm,
        ],
        rowHeights=[
            8 * mm,
            19 * mm,
            8 * mm,
        ],
    )

    table.setStyle(
        TableStyle([

            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "CENTER"
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                2
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                2
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                1
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                1
            ),
        ])
    )

    return table


# =========================================================
# DATA TABLE
# =========================================================

def create_data_table(
    df_part,
    table_number
):

    header_style = ParagraphStyle(
        f"TableHeader{table_number}",
        fontName="Helvetica-Bold",
        fontSize=6.5,
        leading=7,
        alignment=TA_CENTER,
    )

    cell_style = ParagraphStyle(
        f"TableCell{table_number}",
        fontName="Helvetica",
        fontSize=6.2,
        leading=7,
        alignment=TA_LEFT,
    )

    center_style = ParagraphStyle(
        f"TableCenter{table_number}",
        fontName="Helvetica",
        fontSize=6.2,
        leading=7,
        alignment=TA_CENTER,
    )

    table_data = []

    # Header
    table_data.append([
        Paragraph(
            "<b>NO</b>",
            header_style
        ),

        Paragraph(
            "<b>DELIVERY ORDER</b>",
            header_style
        ),

        Paragraph(
            "<b>AIRWAYBILL</b>",
            header_style
        ),

        Paragraph(
            "<b>PROVIDER</b>",
            header_style
        ),

        Paragraph(
            "<b>KOLI QTY</b>",
            header_style
        ),
    ])

    # Data
    for _, row in df_part.iterrows():

        table_data.append([

            Paragraph(
                str(
                    row.get(
                        "NO",
                        ""
                    )
                ),
                center_style
            ),

            Paragraph(
                str(
                    row.get(
                        "DELIVERY ORDER",
                        ""
                    )
                ),
                cell_style
            ),

            Paragraph(
                str(
                    row.get(
                        "AIRWAYBILL",
                        ""
                    )
                ),
                cell_style
            ),

            Paragraph(
                str(
                    row.get(
                        "PROVIDER",
                        ""
                    )
                ),
                cell_style
            ),

            Paragraph(
                str(
                    row.get(
                        "KOLI QTY",
                        ""
                    )
                ),
                center_style
            ),
        ])

    table = Table(
        table_data,

        colWidths=[
            9 * mm,
            42 * mm,
            48 * mm,
            40 * mm,
            20 * mm,
        ],

        repeatRows=1,

        splitByRow=1,
    )

    table.setStyle(
        TableStyle([

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.35,
                colors.grey
            ),

            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor(
                    "#E8EEF3"
                )
            ),

            (
                "ALIGN",
                (0, 0),
                (-1, 0),
                "CENTER"
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                2
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                2
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                2
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                2
            ),
        ])
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

        rightMargin=8 * mm,
        leftMargin=8 * mm,

        topMargin=8 * mm,
        bottomMargin=9 * mm,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleCompact",
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=12,
        alignment=TA_CENTER,
        spaceAfter=3,
    )

    subtitle_style = ParagraphStyle(
        "SubtitleCompact",
        fontName="Helvetica",
        fontSize=7,
        leading=8,
        alignment=TA_CENTER,
        spaceAfter=4,
    )

    normal_style = ParagraphStyle(
        "NormalCompact",
        fontName="Helvetica",
        fontSize=7,
        leading=8,
    )

    small_style = ParagraphStyle(
        "SmallCompact",
        fontName="Helvetica",
        fontSize=6.5,
        leading=7.5,
    )

    elements = []

    # =====================================================
    # TOTAL RESI
    # =====================================================

    total_resi = len(df)

    try:

        total_koli = pd.to_numeric(
            df["KOLI QTY"],
            errors="coerce"
        ).fillna(0).sum()

        total_koli = int(
            total_koli
        )

    except Exception:

        total_koli = 0

    # =====================================================
    # HEADER
    # =====================================================

    elements.append(
        Paragraph(
            "BERITA ACARA SERAH TERIMA",
            title_style
        )
    )

    elements.append(
        Paragraph(
            "BAST",
            subtitle_style
        )
    )

    header_data = [

        [
            Paragraph(
                "<b>Tanggal</b>",
                small_style
            ),

            Paragraph(
                str(tanggal),
                small_style
            ),

            Paragraph(
                "<b>Warehouse</b>",
                small_style
            ),

            Paragraph(
                str(warehouse),
                small_style
            ),

            Paragraph(
                "<b>Jumlah Resi</b>",
                small_style
            ),

            Paragraph(
                f"<b>{total_resi}</b>",
                small_style
            ),
        ],

        [
            Paragraph(
                "<b>Waktu</b>",
                small_style
            ),

            Paragraph(
                str(waktu),
                small_style
            ),

            Paragraph(
                "<b>Courier</b>",
                small_style
            ),

            Paragraph(
                str(courier),
                small_style
            ),

            Paragraph(
                "<b>Jumlah Koli</b>",
                small_style
            ),

            Paragraph(
                f"<b>{total_koli}</b>",
                small_style
            ),
        ],

        [
            Paragraph(
                "<b>Driver</b>",
                small_style
            ),

            Paragraph(
                str(driver),
                small_style
            ),

            Paragraph(
                "<b>Police</b>",
                small_style
            ),

            Paragraph(
                str(police),
                small_style
            ),

            Paragraph(
                "<b>Halaman</b>",
                small_style
            ),

            Paragraph(
                "BAST",
                small_style
            ),
        ],
    ]

    header_table = Table(
        header_data,

        colWidths=[
            20 * mm,
            40 * mm,
            22 * mm,
            47 * mm,
            25 * mm,
            18 * mm,
        ],
    )

    header_table.setStyle(
        TableStyle([

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.35,
                colors.grey
            ),

            (
                "BACKGROUND",
                (0, 0),
                (0, -1),
                colors.HexColor(
                    "#F2F2F2"
                )
            ),

            (
                "BACKGROUND",
                (2, 0),
                (2, -1),
                colors.HexColor(
                    "#F2F2F2"
                )
            ),

            (
                "BACKGROUND",
                (4, 0),
                (4, -1),
                colors.HexColor(
                    "#F2F2F2"
                )
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),

            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                3
            ),

            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                3
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                3
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                3
            ),
        ])
    )

    elements.append(
        header_table
    )

    elements.append(
        Spacer(
            1,
            3 * mm
        )
    )

    # =====================================================
    # 2 TABLE PER PAGE
    # =====================================================

    # Sekitar 20 baris per tabel.
    # 2 tabel = ±40 AWB per halaman.
    # 100 AWB akan otomatis menjadi ±3 halaman.

    rows_per_table = 20

    total_rows = len(df)

    chunks = []

    for start in range(
        0,
        total_rows,
        rows_per_table
    ):

        end = min(
            start + rows_per_table,
            total_rows
        )

        chunks.append(
            df.iloc[start:end]
        )

    # Pastikan ada minimal satu chunk
    if not chunks:

        chunks = [
            df.iloc[0:0]
        ]

    # 2 tabel per halaman
    chunks_per_page = 2

    total_pages_data = math.ceil(
        len(chunks) /
        chunks_per_page
    )

    # =====================================================
    # BUILD TABLES
    # =====================================================

    for page_index in range(
        total_pages_data
    ):

        start_chunk = (
            page_index *
            chunks_per_page
        )

        end_chunk = min(
            start_chunk +
            chunks_per_page,
            len(chunks)
        )

        page_chunks = chunks[
            start_chunk:end_chunk
        ]

        for table_index, chunk in enumerate(
            page_chunks,
            start=1
        ):

            table_number = (
                page_index * 2
                + table_index
            )

            elements.append(
                create_data_table(
                    chunk,
                    table_number
                )
            )

            elements.append(
                Spacer(
                    1,
                    2.5 * mm
                )
            )

        # =================================================
        # SIGNATURE DI SETIAP HALAMAN
        # =================================================

        elements.append(
            create_signature_table(
                security_signature,
                dispatcher_signature,
                driver_signature,

                security_name,
                dispatcher_name,
                driver_name,
            )
        )

        # Page break kecuali halaman terakhir
        if page_index < (
            total_pages_data - 1
        ):

            from reportlab.platypus import PageBreak

            elements.append(
                PageBreak()
            )

    # =====================================================
    # BUILD
    # =====================================================

    doc.build(
        elements,
        canvasmaker=NumberedCanvas
    )

    buffer.seek(0)

    return buffer


# =========================================================
# INPUT
# =========================================================

st.markdown(
    "### Informasi BAST"
)

col1, col2 = st.columns(2)

with col1:

    tanggal = st.text_input(
        "Tanggal",
        placeholder="Contoh: 28 September 2026"
    )

    warehouse = st.text_input(
        "Warehouse",
        placeholder="Contoh: WH Jakarta"
    )

    courier = st.text_input(
        "Courier",
        placeholder="Contoh: JNE"
    )

with col2:

    waktu = st.text_input(
        "Waktu",
        placeholder="Contoh: 14:30"
    )

    driver = st.text_input(
        "Driver",
        placeholder="Nama driver"
    )

    police = st.text_input(
        "Police",
        placeholder="Nomor polisi"
    )


# =========================================================
# DATA
# =========================================================

st.markdown(
    "### Data AWB"
)

st.caption(
    "Copy data dari Excel lalu paste di bawah."
)

paste_data = st.text_area(
    "Paste Data Excel",
    height=180,

    placeholder=(
        "NO\tDELIVERY ORDER\tAIRWAYBILL\t"
        "PROVIDER\tKOLI QTY\n"
        "1\tDO001\tAWB001\tJNE\t2\n"
        "2\tDO002\tAWB002\tSPX\t3"
    )
)

df = None

if paste_data.strip():

    df = parse_paste_data(
        paste_data
    )

    if df is None:

        st.error(
            "Data tidak dapat dibaca."
        )

    else:

        df = fix_broken_rows(
            df
        )

        valid, error_message = (
            validate_file(df)
        )

        if not valid:

            st.error(
                error_message
            )

        else:

            st.success(
                f"{len(df)} AWB berhasil dibaca."
            )

            st.dataframe(
                df[
                    [
                        "NO",
                        "DELIVERY ORDER",
                        "AIRWAYBILL",
                        "PROVIDER",
                        "KOLI QTY",
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )


# =========================================================
# SIGNATURE NAMES
# =========================================================

st.markdown("---")

st.markdown(
    "## ✍️ Penandatangan"
)

name_col1, name_col2, name_col3 = st.columns(3)

with name_col1:

    security_name = st.text_input(
        "Nama Security WH",
        placeholder="Masukkan nama"
    )

with name_col2:

    dispatcher_name = st.text_input(
        "Nama Dispatcher WH",
        placeholder="Masukkan nama"
    )

with name_col3:

    driver_name = st.text_input(
        "Nama Driver Courier",
        placeholder="Masukkan nama"
    )


# =========================================================
# SIGNATURE PAD
# =========================================================

st.markdown(
    "### Tanda Tangan"
)

st.caption(
    "Tanda tangan menggunakan mouse, touchpad, "
    "atau jari pada touchscreen."
)

sig_col1, sig_col2, sig_col3 = st.columns(3)

with sig_col1:

    st.markdown(
        "**Diperiksa oleh**"
    )

    st.caption(
        "Security WH"
    )

    security_signature = signature_pad(
        key="security_signature",
        width=350,
        height=150
    )


with sig_col2:

    st.markdown(
        "**Diserahkan oleh**"
    )

    st.caption(
        "Dispatcher WH"
    )

    dispatcher_signature = signature_pad(
        key="dispatcher_signature",
        width=350,
        height=150
    )


with sig_col3:

    st.markdown(
        "**Diterima oleh**"
    )

    st.caption(
        "Driver Courier"
    )

    driver_signature = signature_pad(
        key="driver_signature",
        width=350,
        height=150
    )


# =========================================================
# STATUS
# =========================================================

status_col1, status_col2, status_col3 = st.columns(3)

with status_col1:

    if (
        security_name.strip()
        and signature_exists(
            security_signature
        )
    ):

        st.success(
            "✅ Security lengkap"
        )

    else:

        st.warning(
            "⚠️ Security belum lengkap"
        )


with status_col2:

    if (
        dispatcher_name.strip()
        and signature_exists(
            dispatcher_signature
        )
    ):

        st.success(
            "✅ Dispatcher lengkap"
        )

    else:

        st.warning(
            "⚠️ Dispatcher belum lengkap"
        )


with status_col3:

    if (
        driver_name.strip()
        and signature_exists(
            driver_signature
        )
    ):

        st.success(
            "✅ Driver lengkap"
        )

    else:

        st.warning(
            "⚠️ Driver belum lengkap"
        )


# =========================================================
# GENERATE
# =========================================================

st.markdown("---")

generate_button = st.button(
    "📄 Generate BAST PDF",
    type="primary",
    use_container_width=True
)


if generate_button:

    # =====================================================
    # HEADER VALIDATION
    # =====================================================

    required_header = {

        "Tanggal": tanggal,

        "Warehouse": warehouse,

        "Courier": courier,

        "Waktu": waktu,

        "Driver": driver,

        "Police": police,
    }

    missing_header = [

        name

        for name, value
        in required_header.items()

        if not str(value).strip()
    ]

    if missing_header:

        st.error(
            "Field berikut belum diisi: "
            + ", ".join(
                missing_header
            )
        )

        st.stop()

    # =====================================================
    # DATA VALIDATION
    # =====================================================

    if df is None or df.empty:

        st.error(
            "Data AWB belum dimasukkan."
        )

        st.stop()

    valid, error_message = (
        validate_file(df)
    )

    if not valid:

        st.error(
            error_message
        )

        st.stop()

    # =====================================================
    # NAME VALIDATION
    # =====================================================

    missing_names = []

    if not security_name.strip():

        missing_names.append(
            "Nama Security WH"
        )

    if not dispatcher_name.strip():

        missing_names.append(
            "Nama Dispatcher WH"
        )

    if not driver_name.strip():

        missing_names.append(
            "Nama Driver Courier"
        )

    if missing_names:

        st.error(
            "Nama berikut belum diisi: "
            + ", ".join(
                missing_names
            )
        )

        st.stop()

    # =====================================================
    # SIGNATURE VALIDATION
    # =====================================================

    missing_signature = []

    if not signature_exists(
        security_signature
    ):

        missing_signature.append(
            "Security WH"
        )

    if not signature_exists(
        dispatcher_signature
    ):

        missing_signature.append(
            "Dispatcher WH"
        )

    if not signature_exists(
        driver_signature
    ):

        missing_signature.append(
            "Driver Courier"
        )

    if missing_signature:

        st.error(
            "Tanda tangan belum lengkap: "
            + ", ".join(
                missing_signature
            )
        )

        st.stop()

    # =====================================================
    # GENERATE
    # =====================================================

    try:

        with st.spinner(
            "Sedang membuat PDF..."
        ):

            pdf_buffer = generate_pdf(

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
            + safe_filename(
                tanggal
            )
            + ".pdf"
        )

        st.success(
            "✅ PDF BAST berhasil dibuat."
        )

        st.download_button(

            label="⬇️ Download BAST PDF",

            data=pdf_buffer.getvalue(),

            file_name=filename,

            mime="application/pdf",

            use_container_width=True,
        )

    except Exception as e:

        st.error(
            f"Gagal membuat PDF: {str(e)}"
        )
