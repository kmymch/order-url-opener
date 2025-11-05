import streamlit as st
import pandas as pd
from io import BytesIO

def open_urls_from_excel(file_bytes):
    """
    アップロードされたエクセルファイルのバイトデータからI列のURLを取得し、ブラウザで開く
    """
    df = pd.read_excel(BytesIO(file_bytes), engine="openpyxl")
    df = df.astype(str)
    st.write(df)
    url_column = df.iloc[:, 8]  # I列(0インデックスでは8番目)
    
    flag = 0
    for url in url_column:
        if isinstance(url, str) and url.startswith("http"):
            if flag == 0:
                flag = 1
            st.write(url)
            js = f"window.open('{url}', '_blank')"
            st.components.v1.html(f"<script>{js}</script>", height=0)
        elif flag == 1:
            break

# Streamlitアプリの設定
st.title("ExcelからURLを開くアプリ")

# セッションステートを初期化
if 'file_content' not in st.session_state:
    st.session_state.file_content = None
if 'last_uploaded_filename' not in st.session_state:
    st.session_state.last_uploaded_filename = None

# ファイルアップロード
uploaded_file = st.file_uploader("エクセルファイルをアップロードしてください", type=["xlsx"])

# 新しいファイルがアップロードされたかチェック
if uploaded_file is not None and uploaded_file.name != st.session_state.last_uploaded_filename:
    st.session_state.last_uploaded_filename = uploaded_file.name
    st.session_state.file_content = uploaded_file.read()
    st.write("ファイルがアップロードされました！")
    # 初回自動実行
    open_urls_from_excel(st.session_state.file_content)
    st.success("URLを開きました！")

# ファイルが処理された後に再試行ボタンを表示
if st.session_state.file_content is not None:
    st.info("ポップアップがブロックされた場合は、ブラウザの設定で許可した後に下のボタンを押してください。")
    if st.button("再試行"):
        open_urls_from_excel(st.session_state.file_content)
        st.success("URLを再度開きました！")
