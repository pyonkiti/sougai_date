# IDを連番で振り直す処理です

import csv

def create_date():
    
    path = r"C:\Users\tanakak\Downloads\障害データのbkup"
    input_file = f"{path}\\障害・個別対応テーブル：障害テーブル一覧_5倍ダミーデータ.csv"
    ouput_file = f"{path}\\障害・個別対応テーブル：障害テーブル一覧_5倍ダミーデータ_new.csv"

    with open(input_file, "r", encoding="utf-8", newline="") as f_in, \
        open(ouput_file, "w", encoding="utf-8", newline="") as f_out:

        reader = csv.reader(f_in)
        writer = csv.writer(f_out, quoting=csv.QUOTE_ALL)

        header = next(reader)
        writer.writerow(header)

        for idx, row in enumerate(reader, start=1):
            if row:                                         # 空行対策
                row[0] = f"{idx:06d}"                       # 先頭カラムを連番に置き換え
            writer.writerow(row)

    print("OK")

create_date()
