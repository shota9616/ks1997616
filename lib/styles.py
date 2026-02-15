"""Common styling utilities for the Streamlit apps."""

import streamlit as st

CUSTOM_CSS = """
<style>
    .main-header {
        font-size: 2rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 0.5rem;
    }

    .sub-header {
        font-size: 1rem;
        color: #64748B;
        margin-bottom: 2rem;
    }

    .app-card {
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        transition: box-shadow 0.2s;
    }

    .app-card:hover {
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    }

    .app-card h3 {
        margin-top: 0;
        color: #1E293B;
    }

    .app-card p {
        color: #64748B;
        font-size: 0.9rem;
    }

    .badge-ready {
        background-color: #DCFCE7;
        color: #166534;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 0.75rem;
        font-weight: 600;
    }

    .badge-coming {
        background-color: #FEF3C7;
        color: #92400E;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 0.75rem;
        font-weight: 600;
    }

    .footer {
        text-align: center;
        color: #94A3B8;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #E2E8F0;
    }
</style>
"""


def apply_styles():
    """Apply custom CSS styles to the app."""
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def page_header(title: str, description: str = ""):
    """Render a consistent page header."""
    st.markdown(f'<div class="main-header">{title}</div>', unsafe_allow_html=True)
    if description:
        st.markdown(
            f'<div class="sub-header">{description}</div>', unsafe_allow_html=True
        )


def footer():
    """Render a consistent footer."""
    st.markdown(
        '<div class="footer">Powered by Claude AI + Streamlit</div>',
        unsafe_allow_html=True,
    )
