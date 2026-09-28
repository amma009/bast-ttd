import os

import streamlit.components.v1 as components


# ============================================================
# REGISTER CUSTOM SIGNATURE COMPONENT
# ============================================================

_COMPONENT_DIR = os.path.join(
    os.path.dirname(__file__),
    "signature_component"
)


_signature_component = components.declare_component(
    "signature_pad",
    path=_COMPONENT_DIR
)


# ============================================================
# SIGNATURE PAD
# ============================================================

def signature_pad(
    key=None,
    width=350,
    height=180,
):
    """
    Menampilkan canvas tanda tangan.

    Return:
        Data URL PNG atau None jika belum ada tanda tangan.
    """

    return _signature_component(
        width=width,
        height=height,
        key=key,
        default=None
    )
