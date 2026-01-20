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



# メモ
# まだできていないこと
# 動作テスト
# 使い方の説明文


import sys
import os
import configparser
import traceback
import inspect
import glob
import csv
import logging
import subprocess
import msal                                         # Microsoft Authentication Library
from pprint import pprint

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
         "CSV_FILE"        : ""}                    # SharePointに保存されるCSVファイル名
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
                logger.error("引数に値が２個以上設定されています。引数の値は１個しか設定できません。")
                return retbln, ret

        # 入力できる引数の一覧
        if argv[-1].lower()   in ["u", "up", "upload"]:
            ret = "up"
        elif argv[-1].lower() in ["d", "down", "download"]:
            ret = "down"
        elif argv[-1].lower() in ["c", "csv"]:
            ret = "csv"
        else:
            logger.error("引数の値の指定に誤りがあります。")
            return retbln, ret

        retbln = True

        return retbln, ret

    # ロギングの開始
    def init_log():
        try:
            retbln = False

            handler   = logging.FileHandler(r'.\log\production.log', mode="w", encoding='utf-8')
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

        # INIファイルの読み込みチェック
        def check_ini(ini_filepath):
            
            try:
                retbln = False
                
                # INIファイルの存在チェック
                if not os.path.exists(ini_filepath):
                    logger.error("SYSTEM.INIファイルが存在しません。")
                    return retbln
                    
                config = configparser.ConfigParser()
                config.read(ini_filepath, encoding="utf-8")

                # dictを読み込む
                for key, value in DICT_INI.items():

                    # INIファイル内のセクションの存在チェック
                    if not config.has_section(key):
                        logger.error("[SHARE_INFO]セクションが存在しません。")
                        return retbln
                
                    # 入れ子のdictを読み込む
                    for key_key in value:

                        # INIファイル内のセクションに属する全オプションの存在チェック
                        if not config.has_option(key, key_key):
                            logger.error(f"{key_key} オプションが存在しません。")
                            return retbln
                
                retbln = True

            except Exception as e:
                msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
                logger.exception(msg_err)
                traceback.print_exc()
            finally:
                return retbln
        
        try:
            retbln = False
            ini_filepath = f"{os.path.dirname(__file__)}" + "\\Common\\" + "SYSTEM.INI"

            # INIファイルのチェック
            if not check_ini(ini_filepath):
                return retbln
            
            # INIファイルの読み込み
            config = configparser.ConfigParser()
            config.read(ini_filepath, encoding="utf-8")
            
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
            DICT_INI["FILE_INFO"]["CSV_FILE"]         = config.get('FILE_INFO', 'CSV_FILE')            # SharePointに保存されるCSVファイル名
            
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
                logger.error("INIファイルに指定したフォルダに障害データのCSVファイルが存在しません。")
                return retbln
            
            logger.info(f"「{DICT_INI["FILE_INFO"]["CSV_FILES"]}」のファイルは{len(csv_list)}個存在しています。")

            # 直近のタイムスタンプの１ファイルを取得
            ComDefine.csv_file = max(csv_list, key=os.path.getmtime)

            logger.info(f"「{ComDefine.csv_file}」が最新のファイルです。")

            # 空ファイルチェック
            if os.path.getsize(ComDefine.csv_file) == 0:
                logger.error("障害データのCSVファイルは0KBです。")
                return retbln
            
            # CSVファイルの文字コードチェック
            if not PROC_HEAD.check_fille_utf(ComDefine.csv_file):
                logger.error("障害データのCSVファイルはUTF-8で作成してください。")
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
            with open(csv_file, newline="", encoding="utf-8") as file:
                reader = csv.reader(file)
                header = next(reader)

                if not len(header) == 11:
                    return retbln
                
            retbln = True
        except Exception as e:
            pass
        finally:
            return retbln

    # ログファイルを表示する
    def disp_log(syori_flg):

        try:
            retbln = False

            if not syori_flg:
                logger.info("----- 処理が異常終了しました。 -----")
            else:
                logger.info("----- 処理が正常に終了しました。 -----")

            subprocess.Popen(["notepad.exe", r".\log\production.log"])
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
                logger.info(f"SharePointの{DICT_INI["SHARE_INFO"]["LM_PATH"]}フォルダには{len(file_list['value'])}個のファイルが存在していました。")

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
            
                    logger.info(f"{cnt}件目 SharePointから「{file_nm}」を削除しました。")
                    
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
            logger.info(f"「{ComDefine.csv_file}」がSharePointへのアップロード元ファイルです。")
            
            # SharePointへのアップロード
            ret = self.clsMsGraph.sys_sharepoint_upload_file(ComDefine.folder_id, ComDefine.csv_file, "障害データ.csv")
            sub_rtn = ret[0]

            if not sub_rtn:
                logger.error("SharePointへのアップロードに失敗しました。")
                logger.error(ret[1])
                return retbln

            logger.info(f"SharePointに「障害データ.csv」をアップロードしました。")

            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            logger.exception(msg_err)
            traceback.print_exc()
        finally:
            return retbln

    # SharePontからのダウンロード
    def download_file(self):

        try:
            retbln = False

            # SharePoint からダウンロード
            download_path = DICT_INI["FILE_INFO"]["DOWNLOAD_PATH"] + "\\" + DICT_INI["FILE_INFO"]["CSV_FILE"]
            ret = self.clsMsGraph.sys_sharepoint_move_file(ComDefine.folder_id, f"{DICT_INI["FILE_INFO"]["CSV_FILE"]}", f"{download_path}")
            sub_rtn = ret[0]

            if not sub_rtn:
                logger.error(f"SharePointから「{DICT_INI["FILE_INFO"]["CSV_FILE"]}」のダウンロードに失敗しました。")
                logger.error(ret[1])
                return retbln
            
            logger.info(f"SharePointから「{DICT_INI["FILE_INFO"]["CSV_FILE"]}」のダウンロードができました。")

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

        # CSVファイルの存在チェック
        if not PROC_HEAD.check_csv(): raise

        if ret_check_argv[1] == "csv": return retbln
            
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
                if not proc.download_file(): raise

        retbln = True
        
    except Exception as e:
        pass
    finally:
        # ログファイルを表示する
        PROC_HEAD.disp_log(retbln)
        return retbln

# メイン処理
main()