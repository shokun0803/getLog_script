# ポップアップ初期設定
$wsobj = new-object -comobject wscript.shell;

# Server パスをセット
$serverPath = "\file-server\pcmaint"

# OneDrive 上のパスをセット
$drivePath = [Environment]::GetEnvironmentVariable("OneDriveCommercial", "User")

# パスをセット
$path = Join-Path $drivePath $serverPath

# 存在チェック用のスクリプトパス
$chkPath = Join-Path $path getLog_task.vbs;

# パスの存在をチェック
if( -not ( Test-Path -Path $chkPath ) ) {
	# パスが探せないのでエラーを表示して終了
	$result = $wsobj.popup("SharePoint にアクセスできません。`r`n`r`nシステム担当に確認してください。",0,"OneDrive アクセス警告");
	exit;
}

# ローカルのディレクトリの存在を確認
$dir = "C:\pcmaint";
if( -not ( Test-Path -Path $dir ) ) {
	# ディレクトリがなければ作成
	New-Item $dir -ItemType Directory;
}

# 実行スクリプトの存在を確認
$vbsFile = Join-Path $dir getLog_task.vbs;
if( -not ( Test-Path -Path $vbsFile ) ) {
	# なければファイルサーバーからコピー
	$fromFile = Join-Path $path getLog_task.vbs;
	Copy-Item -Path $fromFile -Destination $vbsFile
}
$batFile = Join-Path $dir getLog_task.bat;
if( -not ( Test-Path -Path $batFile ) ) {
	# なければファイルサーバーからコピー
	$fromFile = Join-Path $path getLog_task.bat;
	Copy-Item -Path $fromFile -Destination $batFile
}

# スクリプトの実行をタスクスケジューラーに登録
$action = New-ScheduledTaskAction -Execute $vbsFile;
$trigger = New-ScheduledTaskTrigger -DaysInterval 1 -Daily -At "12:30 PM" -RandomDelay "00:30:00";
$settings = New-ScheduledTaskSettingsSet -Compatibility Win8 -RestartInterval "00:01:00" -RestartCount 3 -ExecutionTimeLimit "01:00:00" -Hidden;

try{
	$tasks = (Get-ScheduledTask -TaskName pcmaint -TaskPath \ -ErrorAction stop);
} catch {
	Register-ScheduledTask -TaskPath \ -TaskName pcmaint -Action $action -Trigger $trigger -Settings $settings;
	$tasks = (Get-ScheduledTask -TaskName pcmaint -TaskPath \ -ErrorAction stop);
}

#Write-Output $tasks.Actions;

# すでにタスクが存在し、古いタスクだったら強制的に上書き
if( $tasks.Actions.Execute -ne $vbsFile ) {
	Register-ScheduledTask -TaskPath \ -TaskName pcmaint -Action $action -Trigger $trigger -Settings $settings -Force;
}

# タスク登録が完了したユーザーを登録
$username = $env:UserName;
$td = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss');
$outfile = Join-Path $path setup_user.csv;
$data = @(
	[pscustomobject]@{
		Time = $td
		User = $username
	}
);
$data | Export-Csv -NoTypeInformation $outfile -Encoding UTF8 -Append;
