import streamlit as st
import pandas as pd

from datetime import datetime
from io import StringIO
import io
import re
import base64

from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image as RLImage
)

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle
)
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

from signature_component import signature_pad


# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="BAST Generator",
    page_icon="📦",
    layout="wide"
)


st.title("📦 BAST Generator")

st.caption(
    "Copy data dari Excel, isi data BAST, lalu buat "
    "3 tanda tangan langsung menggunakan mouse atau jari."
)


# ============================================================
# HEADER INPUT
# ============================================================

st.header("📋 Input Header")


col1, col2 = st.columns(2)


with col1:

    tanggal_only = st.date_input(
        "Tanggal",
        value=datetime.now().date()
    )


    warehouse = st.text_input(
        "Warehouse",
        placeholder="Contoh: WH Jakarta"
    )


    courier = st.text_input(
        "Courier Name",
        placeholder="Nama courier"
    )


with col2:

    waktu_only = st.time_input(
        "Waktu",
        value=datetime.now().time()
    )


    driver = st.text_input(
        "Driver Name",
        placeholder="Nama driver"
    )


    police = st.text_input(
        "Police Number",
        placeholder="Nomor kendaraan"
    )


# ============================================================
# DATETIME
# ============================================================

def make_datetime(
    date_obj,
    time_obj
):

    return datetime(
        date_obj.year,
        date_obj.month,
        date_obj.day,
        time_obj.hour,
        time_obj.minute,
        time_obj.second
    )


tanggal = make_datetime(
    tanggal_only,
    waktu_only
)


# ============================================================
# PASTE DATA
# ============================================================

st.header("📊 Paste Data")


raw_text = st.text_area(
    "Copy data dari Excel lalu paste di sini",
    height=300,
    placeholder=(
        "NO\tDELIVERY ORDER\tAIRWAYBILL\tSTATE\t"
        "PROVIDER\tKOLI QTY\n"
        "1\tDO001\tAWB001\tDELIVERED\tJNE\t2"
    )
)


# ============================================================
# SIGNATURE
# ============================================================

st.header("✍️ Tanda Tangan")


st.info(
    "Isi nama masing-masing orang, kemudian gambar "
    "tanda tangan menggunakan mouse atau jari."
)


# ============================================================
# SIGNER NAMES
# ============================================================

sig_col1, sig_col2, sig_col3 = st.columns(3)


with sig_col1:

    st.subheader("Diperiksa oleh")

    st.caption("Security WH")


    security_name = st.text_input(
        "Nama Security",
        key="security_name",
        placeholder="Nama Security"
    )


with sig_col2:

    st.subheader("Diserahkan oleh")

    st.caption("Dispatcher WH")


    dispatcher_name = st.text_input(
        "Nama Dispatcher",
        key="dispatcher_name",
        placeholder="Nama Dispatcher"
    )


with sig_col3:

    st.subheader("Diterima oleh")

    st.caption("Driver Courier")


    driver_name = st.text_input(
        "Nama Driver Courier",
        key="driver_courier_name",
        placeholder="Nama Driver Courier"
    )


# ============================================================
# SIGNATURE CANVAS
# ============================================================

st.subheader("🖊️ Gambar Tanda Tangan")


canvas_col1, canvas_col2, canvas_col3 = st.columns(3)


with canvas_col1:

    st.markdown("**Security WH**")


    security_signature = signature_pad(
        key="security_signature_pad",
        width=350,
        height=180
    )


with canvas_col2:

    st.markdown("**Dispatcher WH**")


    dispatcher_signature = signature_pad(
        key="dispatcher_signature_pad",
        width=350,
        height=180
    )


with canvas_col3:

    st.markdown("**Driver Courier**")


    driver_signature = signature_pad(
        key="driver_signature_pad",
        width=350,
        height=180
    )


# ============================================================
# HELPER
# ============================================================

def safe_filename(text):

    return re.sub(
        r"[^A-Za-z0-9_-]",
        "_",
        str(text)
    )


# ============================================================
# PARSE DATA
# ============================================================

def parse_paste_data(text):

    if not text or not text.strip():

        return None


    try:

        if "\t" in text:

            df = pd.read_csv(
                StringIO(text),
                sep="\t"
            )

        else:

            df = pd.read_csv(
                StringIO(text)
            )


        # Bersihkan nama kolom

        df.columns = [
            str(col).strip()
            for col in df.columns
        ]


        return df


    except Exception:

        return None


# ============================================================
# VALIDATE DATA
# ============================================================

def validate_file(df):

    errors = []


    if df is None:

        errors.append(
            "Data tidak dapat dibaca."
        )

        return False, errors


    if df.empty:

        errors.append(
            "Data kosong."
        )


    required_columns = [
        "NO",
        "DELIVERY ORDER",
        "AIRWAYBILL",
        "STATE",
        "PROVIDER",
        "KOLI QTY"
    ]


    for column in required_columns:

        if column not in df.columns:

            errors.append(
                f"Kolom wajib tidak ada: {column}"
            )


    return (
        len(errors) == 0,
        errors
    )


# ============================================================
# FIX BROKEN ROWS
# ============================================================

def fix_broken_rows(text):

    lines = (
        text
        .replace("\r", "")
        .split("\n")
    )


    fixed_lines = []


    for line in lines:

        line = line.strip()


        if not line:

            continue


        if line.isdigit() and fixed_lines:

            fixed_lines[-1] += (
                "\t" + line
            )

        else:

            fixed_lines.append(line)


    return "\n".join(
        fixed_lines
    )


# ============================================================
# SIGNATURE DATA URL -> BYTES
# ============================================================

def signature_to_bytes(
    signature_data
):

    if not signature_data:

        return None


    try:

        if "," not in signature_data:

            return None


        header, encoded = (
            signature_data.split(
                ",",
                1
            )
        )


        image_bytes = base64.b64decode(
            encoded
        )


        return io.BytesIO(
            image_bytes
        )


    except Exception:

        return None


# ============================================================
# CHECK SIGNATURE
# ============================================================

def signature_exists(
    signature_data
):

    return bool(
        signature_data
        and isinstance(
            signature_data,
            str
        )
        and signature_data.startswith(
            "data:image/"
        )
    )


# ============================================================
# PAGE NUMBER
# ============================================================

class NumberedCanvas(
    canvas.Canvas
):

    def __init__(
        self,
        *args,
        **kwargs
    ):

        super().__init__(
            *args,
            **kwargs
        )

        self.pages = []


    def showPage(self):

        self.pages.append(
            dict(self.__dict__)
        )

        self._startPage()


    def save(self):

        total = len(
            self.pages
        )


        for page in self.pages:

            self.__dict__.update(
                page
            )


            self.draw_page_number(
                total
            )


            super().showPage()


        super().save()


    def draw_page_number(
        self,
        total
    ):

        page = self.getPageNumber()


        self.setFont(
            "Helvetica",
            9
        )


        self.drawRightString(
            A4[0] - 40,
            20,
            f"{page}/{total}"
        )


# ============================================================
# PDF GENERATOR
# ============================================================

def generate_pdf(
    df,
    tanggal,
    warehouse,
    courier,
    driver,
    police,
    security_name,
    dispatcher_name,
    driver_name,
    security_signature,
    dispatcher_signature,
    driver_signature
):

    buffer = io.BytesIO()


    margin = (
        0.5 * inch
    )


    page_width = (
        A4[0]
        - (margin * 2)
    )


    # ========================================================
    # DOCUMENT
    # ========================================================

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin
    )


    styles = (
        getSampleStyleSheet()
    )


    elements = []


    # ========================================================
    # TITLE
    # ========================================================

    title_style = ParagraphStyle(
        "title",
        parent=styles["Title"],
        alignment=1,
        fontSize=18,
        spaceAfter=12
    )


    elements.append(
        Paragraph(
            "<b>BERITA ACARA SERAH TERIMA</b>",
            title_style
        )
    )


    # ========================================================
    # TOTAL KOLI
    # ========================================================

    total_koli = int(
        pd.to_numeric(
            df["KOLI QTY"],
            errors="coerce"
        )
        .fillna(0)
        .sum()
    )


    # ========================================================
    # HEADER TEXT
    # ========================================================

    tanggal_str = (
        tanggal.strftime(
            "%d/%m/%Y %H:%M:%S"
        )
    )


    header_text = f"""
    <b>Tanggal:</b> {tanggal_str}<br/>
    <b>Warehouse:</b> {warehouse}<br/>
    <b>Courier Name:</b> {courier}<br/>
    <b>Driver Name:</b> {driver}<br/>
    <b>Police Number:</b> {police}
    """


    label_style = ParagraphStyle(
        "label",
        parent=styles["Normal"],
        alignment=1,
        fontSize=10
    )


    big_style = ParagraphStyle(
        "big",
        parent=styles["Normal"],
        alignment=1,
        fontSize=20
    )


    total_box = Table(
        [
            [
                Paragraph(
                    "<b>TOTAL KOLI</b>",
                    label_style
                )
            ],

            [
                Paragraph(
                    f"<b>{total_koli}</b>",
                    big_style
                )
            ]
        ],

        colWidths=[
            130
        ],

        rowHeights=[
            25,
            45
        ]
    )


    total_box.setStyle(
        TableStyle([
            (
                "BOX",
                (0, 0),
                (-1, -1),
                1.5,
                colors.black
            ),

            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.lightgrey
            ),

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
            )
        ])
    )


    header_table = Table(
        [
            [
                Paragraph(
                    header_text,
                    styles["Normal"]
                ),

                total_box
            ]
        ],

        colWidths=[
            page_width - 130,
            130
        ]
    )


    header_table.setStyle(
        TableStyle([
            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "MIDDLE"
            )
        ])
    )


    elements.append(
        header_table
    )


    elements.append(
        Spacer(1, 12)
    )


    # ========================================================
    # DATA TABLE
    # ========================================================

    expected_columns = [
        "NO",
        "DELIVERY ORDER",
        "AIRWAYBILL",
        "STATE",
        "PROVIDER",
        "KOLI QTY"
    ]


    df_pdf = (
        df[
            expected_columns
        ]
        .fillna("")
        .astype(str)
    )


    data = [
        list(df_pdf.columns)
    ] + df_pdf.values.tolist()


    table = Table(
        data,
        repeatRows=1
    )


    table.setStyle(
        TableStyle([

            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.HexColor(
                    "#1F4E78"
                )
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.4,
                colors.black
            ),

            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                7
            ),

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
            )

        ])
    )


    elements.append(
        table
    )


    # ========================================================
    # SIGNATURE SECTION
    # ========================================================

    elements.append(
        Spacer(1, 25)
    )


    signature_name_style = ParagraphStyle(
        "signature_name",
        parent=styles["Normal"],
        alignment=1,
        fontSize=9
    )


    role_style = ParagraphStyle(
        "role",
        parent=styles["Normal"],
        alignment=1,
        fontSize=8
    )


    note_style = ParagraphStyle(
        "note",
        parent=styles["Normal"],
        alignment=1,
        fontSize=8
    )


    # ========================================================
    # SECURITY
    # ========================================================

    security_bytes = (
        signature_to_bytes(
            security_signature
        )
    )


    if security_bytes:

        security_img = RLImage(
            security_bytes,
            width=120,
            height=60
        )

    else:

        security_img = Spacer(
            1,
            60
        )


    security_cell = [

        security_img,

        Paragraph(
            f"<b>{security_name}</b>",
            signature_name_style
        ),

        Paragraph(
            "(Security WH)",
            role_style
        )
    ]


    # ========================================================
    # DISPATCHER
    # ========================================================

    dispatcher_bytes = (
        signature_to_bytes(
            dispatcher_signature
        )
    )


    if dispatcher_bytes:

        dispatcher_img = RLImage(
            dispatcher_bytes,
            width=120,
            height=60
        )

    else:

        dispatcher_img = Spacer(
            1,
            60
        )


    dispatcher_cell = [

        dispatcher_img,

        Paragraph(
            f"<b>{dispatcher_name}</b>",
            signature_name_style
        ),

        Paragraph(
            "(Dispatcher WH)",
            role_style
        )
    ]


    # ========================================================
    # DRIVER
    # ========================================================

    driver_bytes = (
        signature_to_bytes(
            driver_signature
        )
    )


    if driver_bytes:

        driver_img = RLImage(
            driver_bytes,
            width=120,
            height=60
        )

    else:

        driver_img = Spacer(
            1,
            60
        )


    driver_cell = [

        driver_img,

        Paragraph(
            f"<b>{driver_name}</b>",
            signature_name_style
        ),

        Paragraph(
            "(Driver Courier)",
            role_style
        )
    ]


    # ========================================================
    # SIGNATURE TABLE
    # ========================================================

    sign = Table(
        [

            [
                Paragraph(
                    "<b>Diperiksa oleh</b>",
                    styles["Normal"]
                ),

                Paragraph(
                    "<b>Diserahkan oleh</b>",
                    styles["Normal"]
                ),

                Paragraph(
                    "<b>Diterima oleh</b>",
                    styles["Normal"]
                )
            ],

            [
                security_cell,
                dispatcher_cell,
                driver_cell
            ],

            [
                "",
                "",
                ""
            ],

            [
                Paragraph(
                    "* BAST ini sebagai bukti bahwa "
                    "paket sudah diserahkan dengan kondisi "
                    "baik dan jumlah koli sesuai.",
                    note_style
                ),

                "",
                ""
            ]
        ],

        colWidths=[
            page_width / 3
        ] * 3
    )


    sign.setStyle(
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
                "SPAN",
                (0, 3),
                (2, 3)
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
            )

        ])
    )


    elements.append(
        sign
    )


    # ========================================================
    # BUILD
    # ========================================================

    doc.build(
        elements,
        canvasmaker=NumberedCanvas
    )


    buffer.seek(0)


    return buffer


# ============================================================
# PROCESS
# ============================================================

if raw_text.strip():

    cleaned = fix_broken_rows(
        raw_text
    )


    df = parse_paste_data(
        cleaned
    )


    valid, errors = validate_file(
        df
    )


    if not valid:

        for error in errors:

            st.error(error)


    else:

        st.success(
            "✅ Data berhasil dibaca."
        )


        # ====================================================
        # PREVIEW
        # ====================================================

        st.subheader(
            "Preview Data"
        )


        st.dataframe(
            df,
            use_container_width=True
        )


        # ====================================================
        # TOTAL KOLI
        # ====================================================

        total_koli = int(
            pd.to_numeric(
                df["KOLI QTY"],
                errors="coerce"
            )
            .fillna(0)
            .sum()
        )


        st.info(
            f"📦 TOTAL KOLI: {total_koli}"
        )


        # ====================================================
        # VALIDATION
        # ====================================================

        names_ready = (

            bool(
                security_name.strip()
            )

            and

            bool(
                dispatcher_name.strip()
            )

            and

            bool(
                driver_name.strip()
            )
        )


        # ====================================================
        # GENERATE BUTTON
        # ====================================================

        if st.button(
            "📄 Generate PDF",
            type="primary",
            use_container_width=True
        ):

            # -----------------------------------------------
            # NAME VALIDATION
            # -----------------------------------------------

            if not names_ready:

                st.error(
                    "Nama Security, Dispatcher, dan "
                    "Driver Courier wajib diisi."
                )

                st.stop()


            # -----------------------------------------------
            # SIGNATURE VALIDATION
            # -----------------------------------------------

            if not signature_exists(
                security_signature
            ):

                st.error(
                    "✍️ Tanda tangan Security WH "
                    "belum dibuat."
                )

                st.stop()


            if not signature_exists(
                dispatcher_signature
            ):

                st.error(
                    "✍️ Tanda tangan Dispatcher WH "
                    "belum dibuat."
                )

                st.stop()


            if not signature_exists(
                driver_signature
            ):

                st.error(
                    "✍️ Tanda tangan Driver Courier "
                    "belum dibuat."
                )

                st.stop()


            # -----------------------------------------------
            # GENERATE
            # -----------------------------------------------

            with st.spinner(
                "Sedang membuat PDF..."
            ):

                pdf = generate_pdf(

                    df=df,

                    tanggal=tanggal,

                    warehouse=warehouse,

                    courier=courier,

                    driver=driver,

                    police=police,

                    security_name=security_name,

                    dispatcher_name=dispatcher_name,

                    driver_name=driver_name,

                    security_signature=(
                        security_signature
                    ),

                    dispatcher_signature=(
                        dispatcher_signature
                    ),

                    driver_signature=(
                        driver_signature
                    )
                )


            # -----------------------------------------------
            # FILE NAME
            # -----------------------------------------------

            warehouse_safe = safe_filename(
                warehouse
            )


            fname = (
                f"BAST_"
                f"{warehouse_safe}_"
                f"{tanggal.strftime('%Y%m%d_%H%M%S')}.pdf"
            )


            # -----------------------------------------------
            # SUCCESS
            # -----------------------------------------------

            st.success(
                "✅ PDF berhasil dibuat!"
            )


            st.download_button(

                label="📥 Download PDF",

                data=pdf,

                file_name=fname,

                mime="application/pdf",

                use_container_width=True
            )
