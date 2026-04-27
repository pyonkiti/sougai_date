$pythonPath = "C:\Users\tanakak\AppData\Local\Programs\Python\Python313"    					# Pythonが入っているパス
$scriptFile1 = "FailSys010.py"                                               					# スクリプト
$scriptFile2 = "FailSys020.py"                                               					# スクリプトファイルの名前
powershell.exe -NoProfile -ExecutionPolicy Bypass "& '$pythonPath\python.exe' '$PSScriptRoot\$scriptFile' up"
powershell.exe -NoProfile -ExecutionPolicy Bypass "& '$pythonPath\python.exe' '$PSScriptRoot\$scriptFile' production"
