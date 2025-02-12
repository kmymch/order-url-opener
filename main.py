import os
import pandas as pd
import webbrowser

def open_urls_from_excel(file_path):
    """
    指定されたエクセルファイルからI列のURLを取得し、ブラウザで開く
    """
    # エクセルファイルを読み込む
    df = pd.read_excel(file_path, engine="openpyxl")

    # I列(9列目)を取得
    url_column = df.iloc[:, 8]  # 0インデックスでI列は8

    # print(url_column)

    # URLが見つかる行を探す
    flag = 0
    for url in url_column:
        # print(url, type(url))
        if isinstance(url, str) and url.startswith("http"):  # URLの判定
            if flag == 0:
                # 最初のURLは新しいウィンドウで開く
                webbrowser.open_new(url)
                flag = 1
            else:
                webbrowser.open_new_tab(url)  # ブラウザでURLを開く
        elif flag == 1:
            break  # URL以外の行が来たら終了

def process_excel_folder(folder_path):
    """
    フォルダ内のすべてのエクセルファイルを処理する
    """
    # フォルダ内のファイルを取得
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx"):  # エクセルファイルのみ対象
            file_path = os.path.join(folder_path, file_name)
            print(f"Processing file: {file_path}")
            open_urls_from_excel(file_path)

# フォルダパスの指定
folder_path = "excel"  # `excel`フォルダを指定
process_excel_folder(folder_path)
