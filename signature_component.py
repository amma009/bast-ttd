import os

import streamlit.components.v1 as components


_COMPONENT_PATH = os.path.join(
    os.path.dirname(__file__),
    "signature_component"
)


_signature_component = components.declare_component(
    "signature_pad",
    path=_COMPONENT_PATH
)


def signature_pad(
    key=None,
    width=350,
    height=180
):
    """
    Signature pad untuk Streamlit.

    Return:
        PNG dalam bentuk Data URL:
        data:image/png;base64,...

        atau None jika belum ada tanda tangan.
    """

    return _signature_component(
        width=width,
        height=height,
        key=key,
        default=None
    )
