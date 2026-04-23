# ************************************************************************************************************
#   システム名         ：障害データ SharePoint連携システム
#   プログラムID       ：FailSys010.py
#   プログラム名       ：メイン処理
# 
# ************************************************************************************************************
#   概要
#   SharePointへの障害データのアップロード/ダウンロードを行う機能です
#   障害データは楽楽販売の障害テーブルで管理しています
# 
# ************************************************************************************************************
#   変更履歴
# 
# ************************************************************************************************************
import sys
import os
import configparser
import traceback
import inspect
import glob
import csv
import logging
import subprocess
import re
import msal                                         # Microsoft Authentication Library
from pprint    import pprint
from pathlib   import Path
from openpyxl  import Workbook

# 共通関数
import Common.ComDefine as ComDefine                # グローバル変数の定義
from Common.SysClsMsGraph import SysClsMsGraph      # SharePoint接続用

# SYSTEM.INIファイルの変数
DICT_INI = {
    "SHARE_INFO" :                                  # SharePoint
        {"CLIENT_ID"       : "",                    # アプリケーション (クライアント) ID
         "CLIENT_SEC"      : "",                    # クライアントシークレット(有効期限あり)
         "TENANT_ID"       : "",                    # テナントID
         "HOST_NM"         : "",                    # ホスト名
         "SITE_PATH"       : "",                    # サイトPath
         "EXPIRATION_DATE" : "",                    # 有効期限
         "LM_PATH"         : ""},                   # SharePontのパス
    "FILE_INFO"  :                                  # CSVファイル関連
        {"CSV_FILES"       : "",                    # CSVファイルのローカルのアップロード元パス
         "DOWNLOAD_PATH"   : "",                    # CSVファイルのローカルのダウンロード先パス
         "CSV_FILE"        : "",                    # SharePointに保存されるCSVファイル名
         "EXCEL_FILE"      : ""}                    # SharePointに保存されるEXCELファイル名
}

# ----------------------------------------------------------------------------------------
# 初期処理
# ----------------------------------------------------------------------------------------
class PROC_HEAD:

    # 引数の入力チェック
    def check_argv(argv):

        retbln = False
        ret    = None

        match len(argv):
            case 1:
                logger.error("引数に値が設定されていません。")
                return retbln, ret
            case 2:
                pass
            case _:
                logger.error("引数の値が２個以上設定されています。引数の値は１個しか設定できません。")
                return retbln, ret

        # 入力できる引数の一覧
        if argv[-1].lower()   in ["u", "up", "upload"]:
            ret = "up"
        elif argv[-1].lower() in ["d", "down", "download"]:
            ret = "down"
        elif argv[-1].lower() in ["c", "csv"]:
            ret = "csv"
        elif argv[-1].lower() in ["e", "excel"]:
            ret = "excel"
        else:
            logger.error("引数の値の指定に誤りがあります。")
            return retbln, ret

        retbln = True

        return retbln, ret

    # ロギングの開始
    def init_log():
        try:
            retbln = False

            # ログファイルのパス
            ComDefine.log_file = fr'{Path(__file__).resolve().parent}\log\production.log'

            handler   = logging.FileHandler(ComDefine.log_file, mode = "w", encoding = 'utf-8')
            formatter = logging.Formatter('%(asctime)s %(levelname)s [%(funcName)s]: %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logger.setLevel(logging.INFO)

            logger.info("----- 処理を開始しました。 -----")

            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

    # INIファイルの読み込み
    def get_ini():

        try:
            retbln = False
            ini_filepath = f"{os.path.dirname(__file__)}" + "\\Common\\" + "SYSTEM.INI"

            # INIファイルの存在チェック
            if not os.path.exists(ini_filepath):
                logger.error("SYSTEM.INIファイルが存在しません。")
                return retbln
            
            # INIファイルの読み込み
            config = configparser.ConfigParser()
            config.read(ini_filepath, encoding = "utf-8-sig")

            # dictを読み込む
            for key, value in DICT_INI.items():

                # INIファイル内のセクションの存在チェック
                if not config.has_section(key):
                    logger.error("[SHARE_INFO]セクションが存在しません。")
                    return retbln
            
                # dictの入れ子dictを読み込む
                for key_key in value:

                    # INIファイル内のセクションに属する全オプションの存在チェック
                    if not config.has_option(key, key_key):
                        logger.error(f"{key_key} オプションが存在しません。")
                        return retbln

            # INIファイルの読み込み
            config = configparser.ConfigParser()
            config.read(ini_filepath, encoding = "utf-8-sig")
            
            logger.info(f"「{ini_filepath}」ファイルを読み込みました。")

            DICT_INI["SHARE_INFO"]["CLIENT_ID"]       = config.get('SHARE_INFO', 'CLIENT_ID')          # SharePoint.アプリケーション (クライアント) ID
            DICT_INI["SHARE_INFO"]["CLIENT_SEC"]      = config.get('SHARE_INFO', 'CLIENT_SEC')         # SharePoint.クライアントシークレット(有効期限あり)
            DICT_INI["SHARE_INFO"]["TENANT_ID"]       = config.get('SHARE_INFO', 'TENANT_ID')          # SharePoint.テナントID
            DICT_INI["SHARE_INFO"]["HOST_NM"]         = config.get('SHARE_INFO', 'HOST_NM')            # ホスト名
            DICT_INI["SHARE_INFO"]["SITE_PATH"]       = config.get('SHARE_INFO', 'SITE_PATH')          # サイトPath
            DICT_INI["SHARE_INFO"]["EXPIRATION_DATE"] = config.get('SHARE_INFO', 'EXPIRATION_DATE')    # 有効期限
            DICT_INI["SHARE_INFO"]["LM_PATH"]         = config.get('SHARE_INFO', 'LM_PATH')            # SharePontのパス
            DICT_INI["FILE_INFO"]["CSV_FILES"]        = config.get('FILE_INFO', 'CSV_FILES')           # CSVファイルのローカルのアップロード元パス
            DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"]    = config.get('FILE_INFO', 'DOWNLOAD_PATH')       # CSVファイルのローカルのダウンロード先パス
            DICT_INI["FILE_INFO"]["CSV_FILE"]         = config.get('FILE_INFO', 'CSV_FILE')            # SharePointに保存される CSVファイル名
            DICT_INI["FILE_INFO"]["EXCEL_FILE"]       = config.get('FILE_INFO', 'EXCEL_FILE')          # SharePointに保存される EXCELファイル名
            
            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

    # CSVファイルの存在チェック
    def check_csv():

        try:
            retbln = False
            csv_list  = glob.glob(DICT_INI["FILE_INFO"]["CSV_FILES"])

            # CSVファイルの存在チェック
            if not csv_list:

                csv_files = os.path.basename(DICT_INI["FILE_INFO"]["CSV_FILES"])
                logger.error(f"取り込み元のCSVファイル「{csv_files}」が１件も存在しません。")
                return retbln
            
            logger.info(f"「{DICT_INI["FILE_INFO"]["CSV_FILES"]}」のファイルは{len(csv_list)}個存在しています。")

            # 直近のタイムスタンプの１ファイルを取得
            ComDefine.csv_file = max(csv_list, key = os.path.getmtime)

            logger.info(f"「{ComDefine.csv_file}」が最新のタイムスタンプです。")

            # 空ファイルチェック
            if os.path.getsize(ComDefine.csv_file) == 0:
                logger.error("障害データのCSVファイルは0KBです。")
                return retbln
            
            # CSVファイルの文字コードチェック
            if not PROC_HEAD.check_fille_utf(ComDefine.csv_file):
                logger.error("障害データのCSVファイルがUTF-8以外で作成されています。")
                return retbln

            # CSVファイルのヘッダのカラム数チェック
            if not PROC_HEAD.check_csv_row(ComDefine.csv_file):
                logger.error("CSVファイルのカラム数に相違があります。")
                return retbln

            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

    # ファイルの文字コードチェック
    def check_fille_utf(csv_file):
        try:
            retbln = False
            with open(csv_file, "r", encoding="utf-8") as file:
                file.read()
            retbln = True
        except UnicodeDecodeError as e:
            pass
        except Exception as e:
            pass
        finally:
            return retbln

    # CSVファイルのヘッダのカラム数チェック
    def check_csv_row(csv_file):
        try:
            retbln = False
            with open(csv_file, newline = "", encoding = "utf-8") as file:
                reader = csv.reader(file)
                header = next(reader)

                if not len(header) == 9:
                    return retbln
                
            retbln = True
        except Exception as e:
            pass
        finally:
            return retbln

    # ログファイルを表示する
    def disp_log(syori_flg, check_argv):

        try:
            retbln = False

            if not syori_flg:
                logger.info("----- 処理が異常終了しました。 -----")
            else:
                logger.info("----- 処理が正常に終了しました。 -----")

            if check_argv == "up":
                pass
            else:
                # 共有Windows Serverではログ表示を止める
                subprocess.Popen(["notepad.exe", ComDefine.log_file])

            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

    # パターンにマッチしたCSV/Excelファイルを削除
    def delete_files(file_flg):
        
        try:
            retbln = False

            match file_flg:
                case "csv":   full_filename = DICT_INI["FILE_INFO"]["CSV_FILE"]
                case "excel": full_filename = DICT_INI["FILE_INFO"]["EXCEL_FILE"]
                case _:       
                    logger.error("パラメータの指定に誤りがあります。")
                    return retbln

            # ファイル名と拡張子に分割
            fil_filename, ext_filename = os.path.splitext(full_filename)
            
            cnt = 0
            pattern = re.compile(rf"^{fil_filename}\d+\{ext_filename}$")
            target_dir = DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"]

            # CSV/Excelファイルをループ
            for fil in Path(target_dir).glob(f"*{ext_filename}"):

                # パターンマッチ
                if pattern.match(fil.name):
                    os.remove(fil)
                    cnt += 1

            if cnt > 0:
                logger.info(f"「{DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"]}\\{fil_filename}[001-{cnt:03}]{ext_filename}」を{cnt}件削除しました。")
            
            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

# ----------------------------------------------------------------------------------------
# CSVファイルの加工処理
# ----------------------------------------------------------------------------------------
class PROC_CSVSYORI():

    # CSVファイル名を変更
    def update_csv():
        try:
            retbln = False

            input_csv = ComDefine.csv_file
            output_csv = DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"] + "\\" + DICT_INI["FILE_INFO"]["CSV_FILE"]

            # コピー先のCSVファイルを削除
            if os.path.exists(output_csv): os.remove(output_csv)
            
            # CSVファイルをリネーム
            os.rename(input_csv, output_csv)

            logger.info(f"「{input_csv}」を「{output_csv}」にリネームしました。")
            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln
        
    # CSVファイルを行単位で分割
    def division_csv():
        
        try:
            retbln = False
            
            input_file = DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"] + "\\" + DICT_INI["FILE_INFO"]["CSV_FILE"]
            
            # 元ファイル名（拡張子除く）を取得
            base_name, ext_name = os.path.splitext(input_file)

            # 1ファイル当りの行数（ヘッダ行を含めない）
            rows_per_file = 19
                
            # 障害データ.csvを読み込む
            with open(input_file, "r", encoding = "utf-8", newline = "") as f:
                reader = csv.reader(f)

                # ヘッダ行を取得
                header = next(reader)

                file_count, row_count = 1, 0
                out_file, writer      = None, None

                for row in reader:
                    # 行数のブレイク条件
                    if row_count % rows_per_file == 0:

                        if out_file: out_file.close()
                            
                        # ファイルを新規オープン
                        out_file = open(f"{base_name}{file_count:03}.csv", "w", encoding = "utf-8-sig", newline = "")
                        writer = csv.writer(out_file)

                        # ヘッダを書き込む
                        writer.writerow(header)

                        file_count += 1

                    # データ行を書き込む
                    writer.writerow(row)
                    row_count += 1

                # 最終のファイルを閉じる
                if out_file: out_file.close()
            
            logger.info(f"「{base_name}[001-{file_count - 1:03}].csv」を{file_count - 1}件作成しました。")
            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

# ----------------------------------------------------------------------------------------
# Excelファイルの加工処理
# ----------------------------------------------------------------------------------------
class PROC_EXCELSYORI():

    # CSVファイルをExcelファイルに書き込む
    def create_excel():

        try:
            retbln = False

            csv_path = DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"] + "\\" + DICT_INI["FILE_INFO"]["CSV_FILE"]
            excel_path = DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"] + "\\" + DICT_INI["FILE_INFO"]["EXCEL_FILE"]

            # CSVファイルの存在チェック
            if not os.path.exists(csv_path):
                logger.error(f"{csv_path}ファイルが存在しません。")
                return retbln

            # Excelファイルを削除
            if os.path.exists(excel_path): os.remove(excel_path)

            # Excelブックを作成
            wb = Workbook()
            ws = wb.active
            ws.title = Path(excel_path).stem

            # CSVファイルをExcelファイルに書き込む
            with open(csv_path, newline = "", encoding = "utf-8") as f:
                for row in csv.reader(f):
                    ws.append(row)

            # Excelブックを書き込む
            wb.save(excel_path)

            logger.info(f"「{DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"]}」にある{DICT_INI["FILE_INFO"]["CSV_FILE"]}を「{DICT_INI["FILE_INFO"]["EXCEL_FILE"]}」に変換しました。")

            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

# ----------------------------------------------------------------------------------------
# SharePointへのアクセス処理
# ----------------------------------------------------------------------------------------
class PROC_SHAREPOINT():

    # コンストラクタ
    def __init__(self):

        self.clsMsGraph = SysClsMsGraph(True, "", "", False, 
                    DICT_INI["SHARE_INFO"]["CLIENT_ID"],                  # クライアントID
                    DICT_INI["SHARE_INFO"]["CLIENT_SEC"],                 # クライアントシークレット
                    DICT_INI["SHARE_INFO"]["TENANT_ID"],                  # テナントID
                    DICT_INI["SHARE_INFO"]["HOST_NM"],                    # ホスト名
                    DICT_INI["SHARE_INFO"]["SITE_PATH"])                  # サイトPath
        
    # SharePointのフォルダIDを取得
    def get_folder_id(self):
        
        try:
            retbln = False

            # SharePointの認証
            ret = self.clsMsGraph.sys_sharepoint_access()
            if not ret[0]:
                logger.error("SharePointへの接続の認証に失敗しました。")
                logger.error(ret[1])
                return retbln
                
            logger.info("SharePointへの接続の認証に成功しました。")
        
            # SharePointのfolder id を取得
            ret = self.clsMsGraph.sys_sharepoint_get_folder_id(DICT_INI["SHARE_INFO"]["LM_PATH"])
            ComDefine.folder_id = ret[0]
        
            if ComDefine.folder_id == None: 
                logger.error("SharePointのフォルダーIDの取得に失敗しました。")
                logger.error(ret[1])
                return retbln

            logger.info(f"SharePointの{DICT_INI["SHARE_INFO"]["LM_PATH"]}のフォルダIDの取得に成功しました。")

            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln
    
    # SharePontのファイルをすべて削除
    def delete_files(self):

        try:
            retbln = False

            # SharePointのファイルの一覧を取得
            ret = self.clsMsGraph.sys_sharepoint_get_filelist(ComDefine.folder_id)
            file_list = ret[0]

            if not file_list:
                logger.error("SharePointのファイルの一覧の取得に失敗しました。")
                logger.error(ret[1])
                return retbln
            
            if len(file_list['value']) == 0:
                logger.info(f"SharePointの中の{DICT_INI["SHARE_INFO"]["LM_PATH"]}フォルダはもともと空です。")
            else:
                logger.info(f"SharePointの{DICT_INI["SHARE_INFO"]["LM_PATH"]}フォルダには{len(file_list['value'])}個のファイルが存在しています。")
                
                for cnt, file_info in enumerate(file_list['value'], start = 1):
                    
                    # ファイル以外は除外
                    if file_info.get('file') == None:
                        continue
                    
                    file_nm = file_info['name']
                    file_id = file_info['id']
                    
                    # SharePointのファイルを削除
                    ret = self.clsMsGraph.sys_sharepoint_del_file(file_nm, file_id)
                    if not ret[0]:
                        logger.error("SharePointのファイルの削除に失敗しました。")
                        logger.error(ret[1])
                        return retbln
            
                logger.info(f"SharePointからファイルを{cnt}削除しました。")
                
            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

    # SharePointへのファイルのアップロード
    def upload_file(self):
            
        try:
            retbln = False

            # ファイル名と拡張子に分割
            fil_filename, ext_filename = os.path.splitext(DICT_INI["FILE_INFO"]["CSV_FILE"])
            
            cnt = 0
            pattern = re.compile(rf"^{fil_filename}\d+\{ext_filename}$")
            target_dir = DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"]

            # CSVファイルをループ
            for fil in Path(target_dir).glob(f"*{ext_filename}"):

                # パターンマッチ
                if pattern.match(fil.name):
                    cnt += 1
           
                    # SharePointへのアップロード
                    ret = self.clsMsGraph.sys_sharepoint_upload_file(ComDefine.folder_id, target_dir + "\\" + fil.name, fil.name)
                    sub_rtn = ret[0]

                    if not sub_rtn:
                        logger.error(f"{fil.name}のSharePointへのアップロードで失敗しました。")
                        logger.error(ret[1])
                        return retbln

            logger.info(f"SharePointに{fil_filename}[001-{cnt:03}]{ext_filename}を{cnt}件アップロードしました。")
            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

    # SharePontからのダウンロード（障害データ[999].csv：複数件）
    def download_files(self):

        try:
            retbln = False

            # SharePointのファイルの一覧を取得
            ret = self.clsMsGraph.sys_sharepoint_get_filelist(ComDefine.folder_id)
            file_list = ret[0]

            if not file_list:
                logger.error("SharePointのファイルの一覧の取得に失敗しました。")
                logger.error(ret[1])
                return retbln
            
            if len(file_list['value']) == 0:
                logger.info(f"SharePointの中の{DICT_INI["SHARE_INFO"]["LM_PATH"]}ダウンロードできるフォルダがありません。")
            else:
                logger.info(f"SharePointの{DICT_INI["SHARE_INFO"]["LM_PATH"]}フォルダには{len(file_list['value'])}個のファイルが存在しています。")
                
                for cnt, file_info in enumerate(file_list['value'], start = 1):
                    
                    # ファイル以外は除外
                    if file_info.get('file') == None: continue
                        
                    download_path = DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"] + "\\" + file_info['name']
                    
                    # SharePoint からダウンロード
                    ret = self.clsMsGraph.sys_sharepoint_move_file(ComDefine.folder_id, file_info['name'], download_path)
                    sub_rtn = ret[0]

                    if not sub_rtn:
                        logger.error(f"SharePointから「{file_info['name']}」のダウンロードに失敗しました。")
                        logger.error(ret[1])
                        return retbln

            logger.info(f"SharePointの「{DICT_INI["SHARE_INFO"]["LM_PATH"]}」から{cnt}件のファイルをダウンロードしました。")

            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln
# ----------------------------------------------------------------------------------------
# メイン処理
# ----------------------------------------------------------------------------------------
logger = logging.getLogger(__name__)

def main():
    try:
        retbln = False

        # ロギングの開始
        if not PROC_HEAD.init_log(): raise
    
        # 引数の入力チェック
        ret_check_argv = PROC_HEAD.check_argv(sys.argv)
        if not ret_check_argv[0]: raise
        
        # INIファイルの読み込み
        if not PROC_HEAD.get_ini(): raise

        if ret_check_argv[1] in ("csv", "excel", "up"):

            # CSVファイルの存在チェック
            if not PROC_HEAD.check_csv(): raise

            # CSVファイル名を変更
            if not PROC_CSVSYORI.update_csv(): raise
            
        if ret_check_argv[1] in ("excel"):

            # CSVファイルをExcelファイルに書き込む
            if not PROC_EXCELSYORI.create_excel(): raise

        if ret_check_argv[1] in ("up"):

            # 障害データ01～99.csvファイルを削除
            if not PROC_HEAD.delete_files("csv"): raise
            
            # CSVファイルを19行単位で分割
            if not PROC_CSVSYORI.division_csv(): raise

        if ret_check_argv[1] in ("up", "down"):
           
            proc = PROC_SHAREPOINT()

            # SharePointへのアクセス処理
            if not proc.get_folder_id(): raise
            
        match ret_check_argv[1]:
            case "up":

                # SharePointのファイル削除
                if not proc.delete_files(): raise
    
                # SharePointへのアップロード
                if not proc.upload_file(): raise
                
            case "down":

                # SharePointからのダウンロード
                if not proc.download_files(): raise

        retbln = True
        
    except Exception as e:
        pass
    finally:
        # ログファイルを表示する
        PROC_HEAD.disp_log(retbln, ret_check_argv[1])
        return retbln

# メイン処理
main()