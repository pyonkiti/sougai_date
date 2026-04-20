# ************************************************************************************************************
#   システム名         ：障害データ SharePoint連携システム
#   プログラムID       ：FailSys020.py
#   プログラム名       ：メイン処理
# 
# ************************************************************************************************************
#   概要
#   本番環境で動作させて、production.logをSlackに送信する機能です
# 
# ************************************************************************************************************
#   変更履歴
# 
# ************************************************************************************************************
import sys
import os
import configparser
import inspect
from pathlib   import Path
from slack_sdk import WebClient

# 共通関数
import Common.ComDefine as ComDefine                # グローバル変数の定義

# SYSTEM.INIファイルの変数
DICT_INI = {
    "SLACKCH1"   :                                  # Slackの接続先情報（技術チャンネル）
        {"CHANNEL"         : "",                    # チャンネル名
         "TOKEN"           : ""},                   # トークン
    "SLACKCH2"   :                                  # Slackの接続先情報（SC）
        {"CHANNEL"         : "",                    # チャンネル名
         "TOKEN"           : ""}                    # トークン
}

# ----------------------------------------------------------------------------------------
# 初期処理
# ----------------------------------------------------------------------------------------
class PROC_HEAD:

    # 引数の入力チェック
    def check_argv(argv):

        retbln = False

        match len(argv):
            case 1:
                return retbln, "引数に値が設定されていません。"
            case 2:
                pass
            case _:
                return retbln, "引数の値が２個以上設定されています。引数の値は１個しか設定できません。"

        # 入力できる引数の一覧
        if argv[-1].lower() in ["p", "production"]:
            ret = "production"
        else:
            return retbln, "引数の値の指定に誤りがあります。"

        retbln = True
        return retbln, ret
    
    # INIファイルの読み込み
    def get_ini():

        try:
            retbln = False
            ini_filepath = f"{os.path.dirname(__file__)}" + "\\Common\\" + "SYSTEM.INI"
            
            # INIファイルの読み込み
            config = configparser.ConfigParser()
            config.read(ini_filepath, encoding = "utf-8-sig")
            
            DICT_INI["SLACKCH1"]["CHANNEL"]   =  config.get('SLACKCH1', 'CHANNEL')         # Slackの接続先情報（技術チャンネル） チャンネル名
            DICT_INI["SLACKCH1"]["TOKEN"]     =  config.get('SLACKCH1', 'TOKEN')           # Slackの接続先情報（〃）            トークン
            DICT_INI["SLACKCH2"]["CHANNEL"]   =  config.get('SLACKCH2', 'CHANNEL')         # Slackの接続先情報（SC）            チャンネル名
            DICT_INI["SLACKCH2"]["TOKEN"]     =  config.get('SLACKCH2', 'TOKEN')           # Slackの接続先情報（〃）            トークン
            
            ComDefine.log_file = fr'{Path(__file__).resolve().parent}\log\production.log'

            retbln = True
            return retbln, None

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            return retbln, msg_err
        finally:
            pass
    
    # ログファイルの存在チェック
    def check_logfile(file):
        try:
            retbln = False

            if not os.path.exists(file):
                return retbln, f"{file}が存在しません。"
            else:
                if os.path.getsize(file) == 0:
                    return retbln, f"{file}が0KBです。"
                
            retbln = True
            return retbln, None
        
        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            return retbln, msg_err
        finally:
            pass

# ----------------------------------------------------------------------------------------
# Slackへの送信処理
# ----------------------------------------------------------------------------------------
class PROC_SLACK():
    
    # Slackにメッセージを送信
    def slack_send(file, ch, linecnt):
        try:
            retbln = False

            match ch:
                case 1:                                                     # 技術チャンネル
                    slack_token = DICT_INI["SLACKCH1"]["TOKEN"]
                    slack_channel = DICT_INI["SLACKCH1"]["CHANNEL"]
                case 2:                                                     # SCloud
                    slack_token = DICT_INI["SLACKCH2"]["TOKEN"]
                    slack_channel = DICT_INI["SLACKCH2"]["CHANNEL"]
                case _:
                    slack_token = None
                    slack_channel = None

            # トークンをセット
            client = WebClient(token = f"{slack_token}")

            # ファイル読み込み
            with open(file, "r", encoding = "utf-8") as f:
                content = f.read()

                match linecnt:
                    case "short":
                        contents = content.splitlines()[0] + "\n" + f"以下、{len(content.splitlines())}件の行を省略・・・\n" + content.splitlines()[-1]
                    case "all":
                        contents = content
                    case _:
                        print(f"{linecnt}はパラメータに指定できません。")
                        contents = ""

                # 3000文字ずつ分割して送信
                for i in range(0, len(contents), 3000):
                    msg = contents[i:i+3000]
                    if msg: client.chat_postMessage( channel = f"{slack_channel}", username = "MessageBot", icon_emoji = ":interrobang:", text = msg)
            
            retbln = True

        except Exception as e:
            msg_err = f"「{__class__.__name__}.{inspect.currentframe().f_code.co_name}で" + "エラーが発生しました。 " + "エラー内容 ： " + f"{e}」"
            return retbln, msg_err
        finally:
            return retbln, None

# ----------------------------------------------------------------------------------------
# メイン処理
# ----------------------------------------------------------------------------------------

def main():
    try:
        retbln = False
    
        # 引数の入力チェック
        ret = PROC_HEAD.check_argv(sys.argv)
        if not ret[0]: raise
        
        # INIファイルの読み込み
        ret = PROC_HEAD.get_ini()
        if not ret[0]: raise

        # ログファイルの存在チェック
        ret = PROC_HEAD.check_logfile(ComDefine.log_file)
        if not ret[0]: raise
        
        # Slackにメッセージを送信
        ret = PROC_SLACK.slack_send(ComDefine.log_file, 1, "short")
        if not ret[0]: raise

        retbln = True
        
    except Exception as e:
        pass
    finally:
        if ret[1] is not None: print(ret[1])
        return retbln

# メイン処理
main()