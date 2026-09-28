import streamlit as st
import pandas as pd
import base64
import io
import re

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
)
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

from signature_component import signature_pad


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="BAST Generator",
    layout="wide"
)

st.title("📦 BAST Generator")
st.caption("Generate Berita Acara Serah Terima (BAST) dengan tanda tangan langsung.")


# =========================================================
# HELPER
# =========================================================

def safe_filename(text):
    """Membersihkan text agar aman digunakan sebagai nama file."""
    if not text:
        return "BAST"

    text = str(text)
    text = re.sub(r'[\\/*?:"<>|]', "_", text)
    text = re.sub(r"\s+", "_", text)

    return text


def parse_paste_data(text):
    """
    Parse data hasil copy dari Excel.
    Mendukung TAB dan beberapa delimiter umum.
    """

    if not text or not text.strip():
        return None

    lines = [
        line.strip()
        for line in text.strip().splitlines()
        if line.strip()
    ]

    if not lines:
        return None

    # Deteksi delimiter
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
            io.StringIO("\n".join(lines)),
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


def validate_file(df):
    """Validasi kolom wajib."""

    required_columns = [
        "NO",
        "DELIVERY ORDER",
        "AIRWAYBILL",
        "STATE",
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


def fix_broken_rows(df):
    """
    Membersihkan data yang kosong / rusak ringan.
    """

    if df is None or df.empty:
        return df

    df = df.copy()

    for col in df.columns:
        df[col] = (
            df[col]
            .astype(str)
            .replace("nan", "", regex=False)
            .str.strip()
        )

    return df


def signature_exists(signature):
    """
    Mengecek apakah signature berasal dari canvas
    dan berupa Base64 PNG.
    """

    return (
        isinstance(signature, str)
        and signature.startswith("data:image/png;base64,")
    )


def signature_to_bytes(signature):
    """
    Mengubah Base64 signature menjadi BytesIO
    agar bisa digunakan ReportLab.
    """

    if not signature_exists(signature):
        return None

    try:
        encoded = signature.split(",", 1)[1]

        image_bytes = base64.b64decode(encoded)

        return io.BytesIO(image_bytes)

    except Exception:
        return None


# =========================================================
# NUMBERED CANVAS
# =========================================================

class NumberedCanvas(canvas.Canvas):

    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)

        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):

        total_pages = len(self._saved_page_states)

        for state in self._saved_page_states:

            self.__dict__.update(state)

            self.draw_page_number(total_pages)

            canvas.Canvas.showPage(self)

        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):

        page_number = self._pageNumber

        self.saveState()

        self.setFont("Helvetica", 8)

        self.drawCentredString(
            A4[0] / 2,
            10 * mm,
            f"Page {page_number} of {page_count}"
        )

        self.restoreState()


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
):
    """
    Generate PDF BAST.
    """

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=18 * mm,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "BASTTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        alignment=TA_CENTER,
        spaceAfter=8,
    )

    normal_style = ParagraphStyle(
        "NormalCustom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
    )

    small_style = ParagraphStyle(
        "Small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
    )

    signature_name_style = ParagraphStyle(
        "SignatureName",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        alignment=TA_CENTER,
    )

    signature_role_style = ParagraphStyle(
        "SignatureRole",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        alignment=TA_CENTER,
    )

    elements = []

    # =====================================================
    # TITLE
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
            ParagraphStyle(
                "BASTSubtitle",
                parent=title_style,
                fontSize=11,
                leading=13,
                spaceAfter=10,
            )
        )
    )

    # =====================================================
    # HEADER INFORMATION
    # =====================================================

    header_data = [
        [
            Paragraph("<b>Tanggal</b>", normal_style),
            Paragraph(str(tanggal), normal_style),
            Paragraph("<b>Warehouse</b>", normal_style),
            Paragraph(str(warehouse), normal_style),
        ],
        [
            Paragraph("<b>Courier</b>", normal_style),
            Paragraph(str(courier), normal_style),
            Paragraph("<b>Waktu</b>", normal_style),
            Paragraph(str(waktu), normal_style),
        ],
        [
            Paragraph("<b>Driver</b>", normal_style),
            Paragraph(str(driver), normal_style),
            Paragraph("<b>Police</b>", normal_style),
            Paragraph(str(police), normal_style),
        ],
    ]

    header_table = Table(
        header_data,
        colWidths=[
            28 * mm,
            58 * mm,
            28 * mm,
            58 * mm,
        ],
    )

    header_table.setStyle(
        TableStyle([
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.grey
            ),
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            ),
            (
                "BACKGROUND",
                (0, 0),
                (0, -1),
                colors.HexColor("#F2F2F2")
            ),
            (
                "BACKGROUND",
                (2, 0),
                (2, -1),
                colors.HexColor("#F2F2F2")
            ),
            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                5
            ),
            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                5
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                5
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                5
            ),
        ])
    )

    elements.append(header_table)
    elements.append(Spacer(1, 8))

    # =====================================================
    # TOTAL KOLI
    # =====================================================

    try:
        total_koli = pd.to_numeric(
            df["KOLI QTY"],
            errors="coerce"
        ).fillna(0).sum()

        total_koli = int(total_koli)

    except Exception:
        total_koli = 0

    total_table = Table(
        [
            [
                Paragraph(
                    "<b>TOTAL KOLI</b>",
                    normal_style
                ),
                Paragraph(
                    f"<b>{total_koli}</b>",
                    normal_style
                ),
            ]
        ],
        colWidths=[
            130 * mm,
            42 * mm
        ],
    )

    total_table.setStyle(
        TableStyle([
            (
                "BOX",
                (0, 0),
                (-1, -1),
                0.8,
                colors.black
            ),
            (
                "INNERGRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.grey
            ),
            (
                "ALIGN",
                (1, 0),
                (1, 0),
                "CENTER"
            ),
            (
                "BACKGROUND",
                (0, 0),
                (0, 0),
                colors.HexColor("#F2F2F2")
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                7
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                7
            ),
        ])
    )

    elements.append(total_table)
    elements.append(Spacer(1, 10))

    # =====================================================
    # DATA TABLE
    # =====================================================

    display_columns = [
        "NO",
        "DELIVERY ORDER",
        "AIRWAYBILL",
        "STATE",
        "PROVIDER",
        "KOLI QTY",
    ]

    table_data = []

    # Header
    table_data.append([
        Paragraph(
            f"<b>{col}</b>",
            small_style
        )
        for col in display_columns
    ])

    # Data
    for _, row in df.iterrows():

        table_data.append([
            Paragraph(
                str(row.get("NO", "")),
                small_style
            ),

            Paragraph(
                str(row.get("DELIVERY ORDER", "")),
                small_style
            ),

            Paragraph(
                str(row.get("AIRWAYBILL", "")),
                small_style
            ),

            Paragraph(
                str(row.get("STATE", "")),
                small_style
            ),

            Paragraph(
                str(row.get("PROVIDER", "")),
                small_style
            ),

            Paragraph(
                str(row.get("KOLI QTY", "")),
                small_style
            ),
        ])

    data_table = Table(
        table_data,
        colWidths=[
            10 * mm,
            37 * mm,
            43 * mm,
            22 * mm,
            35 * mm,
            20 * mm,
        ],
        repeatRows=1,
    )

    data_table.setStyle(
        TableStyle([
            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.4,
                colors.grey
            ),
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor("#D9EAF7")
            ),
            (
                "ALIGN",
                (0, 0),
                (-1, 0),
                "CENTER"
            ),
            (
                "ALIGN",
                (0, 1),
                (0, -1),
                "CENTER"
            ),
            (
                "ALIGN",
                (5, 1),
                (5, -1),
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
                4
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                4
            ),
        ])
    )

    elements.append(data_table)
    elements.append(Spacer(1, 12))

    # =====================================================
    # SIGNATURES
    # =====================================================

    security_bytes = signature_to_bytes(
        security_signature
    )

    dispatcher_bytes = signature_to_bytes(
        dispatcher_signature
    )

    driver_bytes = signature_to_bytes(
        driver_signature
    )

    def make_signature_image(image_bytes):

        if not image_bytes:
            return Spacer(1, 60)

        try:
            image_bytes.seek(0)

            image = RLImage(
                image_bytes,
                width=42 * mm,
                height=22 * mm,
            )

            return image

        except Exception:
            return Spacer(1, 60)

    security_img = make_signature_image(
        security_bytes
    )

    dispatcher_img = make_signature_image(
        dispatcher_bytes
    )

    driver_img = make_signature_image(
        driver_bytes
    )

    signature_data = [
        [
            Paragraph(
                "<b>Diperiksa oleh</b>",
                signature_role_style
            ),
            Paragraph(
                "<b>Diserahkan oleh</b>",
                signature_role_style
            ),
            Paragraph(
                "<b>Diterima oleh</b>",
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
                str("Security WH"),
                signature_name_style
            ),
            Paragraph(
                str("Dispatcher WH"),
                signature_name_style
            ),
            Paragraph(
                str("Driver Courier"),
                signature_name_style
            ),
        ],
    ]

    signature_table = Table(
        signature_data,
        colWidths=[
            58 * mm,
            58 * mm,
            58 * mm,
        ],
        rowHeights=[
            10 * mm,
            25 * mm,
            10 * mm,
        ],
    )

    signature_table.setStyle(
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

    elements.append(signature_table)
    elements.append(Spacer(1, 8))

    # =====================================================
    # NOTE
    # =====================================================

    elements.append(
        Paragraph(
            "Dokumen ini dibuat sebagai bukti serah terima barang.",
            small_style
        )
    )

    # =====================================================
    # BUILD PDF
    # =====================================================

    doc.build(
        elements,
        canvasmaker=NumberedCanvas
    )

    buffer.seek(0)

    return buffer


# =========================================================
# INPUT FORM
# =========================================================

st.markdown("### Informasi BAST")

col1, col2, col3 = st.columns(3)

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
        placeholder="Nomor polisi kendaraan"
    )

with col3:

    st.markdown("#### Data Excel")

    st.caption(
        "Copy data dari Excel lalu paste ke kotak di bawah."
    )


# =========================================================
# DATA EXCEL
# =========================================================

paste_data = st.text_area(
    "Paste Data Excel",
    height=220,
    placeholder=(
        "NO\tDELIVERY ORDER\tAIRWAYBILL\tSTATE\t"
        "PROVIDER\tKOLI QTY\n"
        "1\tDO001\tAWB001\tOK\tJNE\t2\n"
        "2\tDO002\tAWB002\tOK\tSPX\t3"
    )
)

df = None

if paste_data.strip():

    df = parse_paste_data(paste_data)

    if df is None:

        st.error(
            "Data tidak dapat dibaca. "
            "Pastikan format hasil copy dari Excel benar."
        )

    else:

        df = fix_broken_rows(df)

        valid, error_message = validate_file(df)

        if not valid:

            st.error(error_message)

        else:

            st.success(
                f"Data berhasil dibaca: {len(df)} baris"
            )

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )


# =========================================================
# SIGNATURE SECTION
# =========================================================

st.markdown("---")
st.markdown("## ✍️ Tanda Tangan")

st.info(
    "Tanda tangan langsung menggunakan mouse, touchpad, "
    "atau jari pada layar touchscreen."
)

sig_col1, sig_col2, sig_col3 = st.columns(3)


with sig_col1:

    st.markdown("### Diperiksa oleh")
    st.caption("Security WH")

    security_signature = signature_pad(
        key="security_signature",
        width=350,
        height=180
    )


with sig_col2:

    st.markdown("### Diserahkan oleh")
    st.caption("Dispatcher WH")

    dispatcher_signature = signature_pad(
        key="dispatcher_signature",
        width=350,
        height=180
    )


with sig_col3:

    st.markdown("### Diterima oleh")
    st.caption("Driver Courier")

    driver_signature = signature_pad(
        key="driver_signature",
        width=350,
        height=180
    )


# =========================================================
# SIGNATURE STATUS
# =========================================================

st.markdown("### Status Tanda Tangan")

status_col1, status_col2, status_col3 = st.columns(3)

with status_col1:

    if signature_exists(security_signature):
        st.success("✅ Security WH sudah tanda tangan")
    else:
        st.warning("⚠️ Security WH belum tanda tangan")

with status_col2:

    if signature_exists(dispatcher_signature):
        st.success("✅ Dispatcher WH sudah tanda tangan")
    else:
        st.warning("⚠️ Dispatcher WH belum tanda tangan")

with status_col3:

    if signature_exists(driver_signature):
        st.success("✅ Driver Courier sudah tanda tangan")
    else:
        st.warning("⚠️ Driver Courier belum tanda tangan")


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

    # ---------------------------------------------
    # VALIDASI HEADER
    # ---------------------------------------------

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
        for name, value in required_header.items()
        if not str(value).strip()
    ]

    if missing_header:

        st.error(
            "Field berikut belum diisi: "
            + ", ".join(missing_header)
        )

        st.stop()

    # ---------------------------------------------
    # VALIDASI DATA
    # ---------------------------------------------

    if df is None or df.empty:

        st.error(
            "Data Excel belum dimasukkan."
        )

        st.stop()

    valid, error_message = validate_file(df)

    if not valid:

        st.error(error_message)

        st.stop()

    # ---------------------------------------------
    # VALIDASI SIGNATURE
    # ---------------------------------------------

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
            + ", ".join(missing_signature)
        )

        st.stop()

    # ---------------------------------------------
    # GENERATE PDF
    # ---------------------------------------------

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
                security_signature=security_signature,
                dispatcher_signature=dispatcher_signature,
                driver_signature=driver_signature,
            )

        filename = (
            f"BAST_"
            f"{safe_filename(warehouse)}_"
            f"{safe_filename(tanggal)}.pdf"
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
