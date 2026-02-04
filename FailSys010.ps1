$pythonPath = "C:\Users\tanakak\AppData\Local\Programs\Python\Python313"    # Pythonが入っているパス
$scriptFile = "FailSys010.py"                                               # スクリプトファイルの名前
powershell.exe -NoProfile -ExecutionPolicy Bypass "& '$pythonPath\python.exe' '$PSScriptRoot\$scriptFile' csv"
