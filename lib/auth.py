"""Simple password authentication gate for Streamlit apps."""

import streamlit as st


def check_auth() -> bool:
    """Check if user is authenticated. Shows login form if not.

    Returns:
        True if authenticated, False otherwise.
    """
    if st.session_state.get("authenticated"):
        return True

    st.markdown("### :lock: ログイン")
    st.markdown("アプリを利用するにはパスワードを入力してください。")

    password = st.text_input("パスワード", type="password", key="auth_password_input")

    if st.button("ログイン", type="primary"):
        if password == st.secrets.get("app_password", ""):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません。")

    return False


def logout():
    """Log out the current user."""
    st.session_state["authenticated"] = False
    st.rerun()
