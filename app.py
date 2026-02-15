"""AI経営管理ポータル - ホームページ"""

import streamlit as st
from lib.auth import check_auth, logout
from lib.styles import apply_styles, footer

st.set_page_config(
    page_title="AI経営管理ポータル",
    page_icon=":briefcase:",
    layout="wide",
    initial_sidebar_state="expanded",
)

apply_styles()

# --- Authentication ---
if not check_auth():
    st.stop()

# --- Sidebar ---
with st.sidebar:
    st.markdown("### AI経営管理ポータル")
    st.caption("制度とテクノロジーで中小企業の成果を作る")
    st.divider()
    if st.button("ログアウト"):
        logout()

# --- Main Content ---
st.markdown("# :briefcase: AI経営管理ポータル")
st.markdown("ワークフローをWebアプリとして利用できます。左のサイドバーからツールを選択してください。")

st.divider()

# --- App Cards ---
col1, col2 = st.columns(2)

with col1:
    st.markdown("### :clipboard: 省力化補助金 申請書類生成")
    st.markdown("ヒアリングシート + 決算書 + 登記簿から、省力化投資補助金（一般型）の申請書類11種を自動生成します。")
    st.caption(":green_circle: 稼働中")

with col2:
    st.markdown("### :pencil: その他ツール")
    st.markdown("今後、記事下書き・戦略分析など追加ツールを実装予定です。")
    st.caption(":yellow_circle: 開発中")

st.divider()

# --- Status ---
st.markdown("### :gear: システム状況")

status_col1, status_col2 = st.columns(2)

with status_col1:
    st.metric("登録アプリ", "1")

with status_col2:
    st.metric("稼働中", "1")

footer()
