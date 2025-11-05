import streamlit as st
import pandas as pd
from io import BytesIO

def open_urls_from_excel(file):
    """
    アップロードされたエクセルファイルからI列のURLを取得し、ブラウザで開く
    """
    df = pd.read_excel(file, engine="openpyxl")
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

# ファイルアップロード
uploaded_file = st.file_uploader("エクセルファイルをアップロードしてください", type=["xlsx"])

if uploaded_file is not None:
    st.write("ファイルがアップロードされました！")
    open_urls_from_excel(BytesIO(uploaded_file.read()))
    st.success("URLを開きました！")
