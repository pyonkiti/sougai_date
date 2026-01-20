# ************************************************************************************************************
#   システム名         ：障害データ SharePoint連携システム
#   プログラムID       ：SysClsMsGraph.py
#   プログラム名       ：SharePoint 連携クラス
# 
# ************************************************************************************************************
#   概要
#   SharePoint との連携を行うクラスです
# 
# ************************************************************************************************************
#   変更履歴
# 
# ************************************************************************************************************

# 結構自分用にカスタマイズしているので、コードの整理が必要

import inspect
import sys
import msal
import requests 

from datetime import tzinfo
from pprint import pprint
from msal import ConfidentialClientApplication 

# ------------------------------------
# グローバル変数
# ------------------------------------
debug_flg = False                       # ログ出力時に画面メッセージの出力要否
LOG_FILENM = ""                         # ログファイル名
LOG_SP_FILENM = ""                      # ログファイル名（簡易版）
log_sp_err_write = False                # ログファイル（簡易版）エラーメッセージ出力Flg

client_id = ""                          # SharePoint.アプリケーション (クライアント) ID
client_secret = ""                      # SharePoint.クライアントシークレット
tenant_id = ""                          # SharePoint.テナントID
host_nm  = ""                           # SharePoint.ホスト名
site_path = ""                          # SharePoint.サイトPath

access_token = None
site_id = ""
drive_id = ""


class SysClsMsGraph():
    
    # コンストラクタ
    def __init__(self, arg_debug_flg, arg_LOG_FILENM, arg_LOG_SP_FILENM, arg_log_sp_err_write, arg_client_id, arg_client_sec, arg_tenant_id, arg_host_nm, arg_site_path):
        
        global debug_flg 
        debug_flg = arg_debug_flg
        
        global LOG_FILENM
        LOG_FILENM = arg_LOG_FILENM
        
        global LOG_SP_FILENM
        LOG_SP_FILENM = arg_LOG_SP_FILENM
        
        global log_sp_err_write
        log_sp_err_write = arg_log_sp_err_write

        global client_id
        global client_secret
        global tenant_id
        global host_nm
        global site_path

        client_id = arg_client_id
        client_secret = arg_client_sec
        tenant_id = arg_tenant_id
        host_nm = arg_host_nm
        site_path = arg_site_path


    # ************************************************************
    # メソッド  ：SharePoint の接続情報の取得・設定
    # ************************************************************
    def sys_sharepoint_access(self):

        rtn_value = False
        msg = None

        try:
            app = ConfidentialClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant_id}", client_credential=client_secret)
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

            if not "access_token" in result:
                msg = "アクセストークンの取得に失敗しました。有効期限が切れている可能性があります。"
                return rtn_value, msg
            
            global access_token
            access_token = result['access_token']
            
            # ************************************************************
            # メソッド  ：SharePoint のサイトIDの取得
            # ************************************************************
            def get_sharepoint_site_by_path(arg_access_token, arg_hostname, arg_server_relative_path):
                    
                try:
                    url = f"https://graph.microsoft.com/v1.0/sites/{arg_hostname}:/{arg_server_relative_path}"
                    headers = {"Authorization": f"Bearer {arg_access_token}", "Content-Type": "application/json"}
                    
                    response = requests.get(url, headers=headers)
                    
                    if response.status_code != 200:
                        return None
                    
                    site_info = response.json()
                    return site_info['id']
                            
                except Exception as e1:
                    return None

            # ------------------------------------------------------

            # サイトIDの取得
            global site_id
            site_id = get_sharepoint_site_by_path(access_token, host_nm, "sites/" + site_path)
            
            if site_id == None:
                msg = " サイトIDの取得に失敗しました。"
                return rtn_value, msg
            
            # サイトIDからドライブIDを取得
            graph_api_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
            headers = { "Authorization": "Bearer " + access_token, "Content-Type": "application/json"}
                         
            response = requests.get(graph_api_url, headers=headers)
                
            if response.status_code != 200:
                msg = "ドライブIDの取得に失敗しました（サイトへのアクセスに失敗しました）。"
                return rtn_value, msg
                    
            children = response.json()
            
            if not "value" in children:
                msg = "ドライブIDの取得に失敗しました（childrenキーが存在しません）。"
                return rtn_value, msg

            if not "id" in children['value'][0]:
                msg = "ドライブIDの取得に失敗しました（childrenキーの中にidキーが存在しません）。"
                return rtn_value, msg

            # childrenからドライブIDを取得
            global drive_id
            drive_id = children['value'][0]['id']

            rtn_value = True
        
        except Exception as e:
            msg = f"{e}"
        finally:
            return rtn_value, msg

    # ************************************************************
    # SharePoint のフォルダの URL より、フォルダID の取得
    # ************************************************************
    def sys_sharepoint_get_folder_id(self, arg_folder_nm):

        rtn_value = None
        msg = None

        try:
            # URL フォルダ
            graph_api_url_folder = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/"
        
            if arg_folder_nm.strip() == "":
                msg = "INIファイルにフォルダ名が登録されていません。"
                return rtn_value, msg
            
            folder_nm = arg_folder_nm
            
            if folder_nm[0]  == "/": folder_nm = folder_nm[1:]
            if folder_nm[-1] == "/": folder_nm = folder_nm[0:len(folder_nm)-1]
                     
            graph_api_url_folder = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{folder_nm}:/"
            headers = {"Authorization": "Bearer " + access_token, "Content-Type": "application/json"}
                        
            response = requests.get(graph_api_url_folder, headers=headers)
                   
            if response.status_code != 200:
                msg = "フォルダIDの取得に失敗しました。"
                return rtn_value, msg
                
            folder_info = response.json()
            folder_id = folder_info['id']
            
            rtn_value = folder_id
        
        except Exception as e:
            msg = f"{e}"
            rtn_value = None
        finally:
            return rtn_value, msg

    # ************************************************************
    # SharePoint へファイルをアップロード    
    # ************************************************************
    def sys_sharepoint_upload_file(self, arg_folder_id, arg_in_file_nm, arg_out_file_nm):

        rtn_value = False
        msg = None

        try:
            graph_api_url_fileupload = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{arg_folder_id}:/{arg_out_file_nm}:/content"
            headers = { "Authorization": "Bearer " + access_token, "Content-Type": "application/json"}
            
            # rb:バイナリモード
            with open(arg_in_file_nm, 'rb') as file:
                file_content = file.read()

            response = requests.put(graph_api_url_fileupload, headers=headers, data=file_content)
                   
            if response.status_code < 200 or response.status_code >= 300:
                msg = "SharePointへのファイルのアップロードに失敗しました。"
                return rtn_value, msg
            
            rtn_value = True

        except Exception as e:
            msg = f"{e}"
        finally:
            return rtn_value, msg

    # ************************************************************
    # SharePoint のファイルをローカルにダウンロード
    # ************************************************************
    def sys_sharepoint_move_file(self, arg_from_folder_id, arg_from_file_nm, arg_to_local_file_nm):
        
        rtn_value = False
        msg = None

        try:
            graph_api_url_file_from = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{arg_from_folder_id}:/{arg_from_file_nm}:/content"
            headers = {"Authorization": "Bearer " + access_token, "Content-Type": "application/json"}
                       
            # ダウンロード元ファイルの取得
            response_f = requests.get(graph_api_url_file_from, headers=headers)
                   
            if response_f.status_code != 200:
                msg = "ダウンロードファイルの取得に失敗しました。"
                return rtn_value, msg
                
            from_file_content = response_f.content

            # ローカルにコピー
            with open(arg_to_local_file_nm, "wb") as file:
                file.write(from_file_content)
            
            rtn_value = True
        
        except Exception as e:
            msg = f"{e}"
        finally:
            return rtn_value, msg

    # ************************************************************
    # SharePoint 上のファイルの削除
    # ************************************************************
    def sys_sharepoint_del_file(self, arg_file_nm, arg_file_id):

        rtn_value = False
        msg = None

        try:
            graph_api_url_file_2 = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{arg_file_id}"
            headers = {"Authorization": "Bearer " + access_token, "Content-Type": "application/json"}
                        
            # ファイルの削除
            response_d = requests.delete(graph_api_url_file_2, headers=headers)

            if response_d.status_code != 204:
                msg = "SharePointのファイルの削除に失敗しました。"
                return rtn_value, msg
                    
            rtn_value = True
        
        except Exception as e:
            msg = f"{e}"
        finally:
            return rtn_value, msg

    # ************************************************************
    # SharePoint より、ファイルの一覧を取得    
    # ************************************************************
    def sys_sharepoint_get_filelist(self, arg_folder_id):

        rtn_value = None
        msg = None

        try:
            graph_api_url_fileread = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{arg_folder_id}/children"
            headers = { "Authorization": "Bearer " + access_token, "Content-Type": "application/json"}
                         
            response = requests.get(graph_api_url_fileread, headers=headers)
            
            rtn_value = response.json()

        except Exception as e:
            msg = f"{e}"
        finally:
            return rtn_value, msg

if __name__ == "__main__":
    pass
