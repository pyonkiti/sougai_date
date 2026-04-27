# csvファイルをコピーする処理です。
import shutil

input_path = r"C:\Users\tanakak\Downloads\障害データのbkup"
ouput_path = r"C:\Users\tanakak\Downloads"

input_file = "障害・個別対応テーブル：障害テーブル一覧.csv"

shutil.copy(rf"{input_path}\{input_file}", f"{ouput_path}")
print("OK")
