@SETLOCAL&SETLOCAL ENABLEDELAYEDEXPANSION&(chcp 65001>NUL)&PUSHD "%~dp0"&(SET SCRIPT_PATH=%~f0)&(SET SCRIPT_DIR=%~dp0)&(SET SCRIPT_NAME=%~nx0)&(SET SCRIPT_BASE_NAME=%~n0)&(SET SCRIPT_EXTENSION=%~x0)&(SET SCRIPT_ARGUMENTS=%*)&(POWERSHELL -NoLogo -Sta -NoProfile -ExecutionPolicy Unrestricted "&([scriptblock]::create('$OutputEncoding=[Console]::OutputEncoding;'+\"`n\"+((gc -Encoding UTF8 -Path \"!SCRIPT_PATH!\"|?{$_.readcount -gt 1})-join\"`n\")))" !SCRIPT_ARGUMENTS!)&(SET EXIT_CODE=!ERRORLEVEL!)&POPD&PAUSE&EXIT !EXIT_CODE!&ENDLOCAL&GOTO :EOF
# ポップアップ初期設定
$wsobj = new-object -comobject wscript.shell;

# Server パスをセット
$serverPath = "\file-server\pcmaint"

# OneDrive 上のパスをセット
$drivePath = [Environment]::GetEnvironmentVariable("OneDriveCommercial", "User")

# パスをセット
$path = Join-Path $drivePath $serverPath

# 存在チェック用のスクリプト名
$file_name = "getLog_script.ps1";

# 存在チェック用のスクリプトパス
$chkPath = Join-Path $path $file_name;

# パスの存在をチェック
if( -not ( Test-Path -Path $chkPath ) ) {
	# パスが探せないのでエラーを表示して終了
	$result = $wsobj.popup("SharePoint にアクセスできません。`r`n`r`nシステム担当に確認してください。",0,"OneDrive アクセス警告");
	exit;
}

# スクリプトファイルを実行
& $chkPath;
