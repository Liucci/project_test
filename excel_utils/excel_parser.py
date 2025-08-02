import pandas as pd

def extract_names_from_excel(file_path):
    df = pd.read_excel(file_path, sheet_name="原本", header=None)

    # 8行目（インデックス7）以降のデータを対象にする
    data_start_row = 8
    name_column = 2  # 列C：Unnamed: 2

    names = df.iloc[data_start_row:, name_column]
    names = names.dropna().unique().tolist()
    #print(f"抽出された名前: {names}")
    return names
