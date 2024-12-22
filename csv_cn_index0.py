import os
import pandas as pd
import subprocess

def main():
    # フォルダパスをプロンプト入力で指定
    folder_path = input("処理するフォルダのパスを入力してください: ").strip()

    # フォルダが存在するか確認
    if not os.path.isdir(folder_path):
        print("指定されたフォルダが存在しません。終了します。")
        return

    # スキップする行数をプロンプト入力で指定
    try:
        skip_rows = int(input("スキップする行数を入力してください: ").strip())
    except ValueError:
        print("無効な入力です。整数を入力してください。終了します。")
        return

    output_filename = os.path.join(folder_path, "output_file.xlsx")
    merged_df = None  # 最終的な結合データフレーム
    header_loaded = False

    # フォルダ内のCSVファイルを処理
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".csv"):
            file_path = os.path.join(folder_path, file_name)
            
            # 最初のファイルのみヘッダを読み込む
            if not header_loaded:
                merged_df = pd.read_csv(file_path, skiprows=skip_rows, encoding='shift-jis')
                merged_df.insert(0, "元ファイル名", file_name)  # ファイル名を格納する列を追加
                header_loaded = True
            else:
                df = pd.read_csv(file_path, skiprows=skip_rows, encoding='shift-jis')  # ヘッダを読み込まない
                df.insert(0, "元ファイル名", file_name)  # ファイル名を格納する列を追加
                df.columns = merged_df.columns  # 列名を合わせる
                merged_df = pd.concat([merged_df, df], ignore_index=True)  # ヘッダ行を除いて結合

    if merged_df is not None:
        # 結果をエクセルファイルに保存
        merged_df.to_excel(output_filename, index=False)
        print(f"サマリファイルを保存しました: {output_filename}")

        # 保存したファイルを開く
        subprocess.Popen(["start", "", output_filename], shell=True)
    else:
        print("CSVファイルが見つかりませんでした。")

if __name__ == "__main__":
    main()
