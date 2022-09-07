# ポップアップ初期設定
$wsobj = new-object -comobject wscript.shell;

# Server パスをセット
$serverPath = "\file-server\timecard_log"

# OneDrive 上のパスをセット
$drivePath = [Environment]::GetEnvironmentVariable("OneDriveCommercial", "User")

# パスをセット
$path = Join-Path $drivePath $serverPath

# 存在チェック用のスクリプトパス
$chkPath = Join-Path $path path_chk.txt;

# パスの存在をチェック
if( -not ( Test-Path -Path $chkPath ) ) {
	# パスが探せないのでエラーを表示して終了
	$result = $wsobj.popup("SharePoint にアクセスできません。`r`n`r`nシステム担当に確認してください。",0,"OneDrive アクセス警告");
	exit;
}

# 端末のユーザー名を取得
$username = $env:UserName;
$uinfo = (Get-WmiObject Win32_UserAccount | Where-Object { $_.name -eq $username });
$fullname = $uinfo.FullName;

# 今日の日付
$td = (Get-Date);

# 年のディレクトリを確認、存在していなければ作成
$year = Join-Path $path $td.Year;
if( -not ( Test-Path -Path $year ) ) {
	# ディレクトリがなければ作成
	New-Item $year -ItemType Directory;
}

# 自分のファイルを確認、存在してなければ今年のファイルを新規生成
$fullname = $fullname.Replace("　", " ");
$file_name = $fullname.Replace(" ", "_") + ".csv";
$file = Join-Path $year $file_name;
if( -not ( Test-Path -Path $file ) ) {
	# 取得開始日時、終了日時設定
	$sd = Get-Date $td.ToString('yyyy-01-01 00:00:00');
	$date = Get-Date $td.ToString('yyyy-MM-dd 23:59:59');
	$ed = $date.AddDays(-1);
	# 起動終了時間の取得
	$array = GET-WinEvent -FilterHashTable @{LogName='System'; StartTime=$sd; EndTime=$ed} | Where-Object{$_.Id -eq 7001 -or $_.Id -eq 7002} | select-Object TimeCreated,Id,Message | Sort-Object -Property @{E = {Get-Date $_.TimeCreated}};

	# 日時データのみ成形
	$csvlist = @();
	$flag = 0;
	foreach ($data in $array) {
		if( $data.Id -eq "7001" ) {
			$attend = $data.TimeCreated.ToString('HH:mm:ss');
			$flag = 1;
		} elseif( $data.Id -eq "7002" -and $flag -eq 0 ) {
			$csvlist += @(
				[PSCustomObject]@{
					date = $data.TimeCreated.ToString('MM-dd')
					leav = ""
					attend = $data.TimeCreated.ToString('HH:mm:ss')
				}
			);
			$flag = 0;
		} elseif( $data.Id -eq "7002" -and $flag -eq 1 ) {
			$csvlist += @(
				[PSCustomObject]@{
					date = $data.TimeCreated.ToString('MM-dd')
					leav = $attend
					attend = $data.TimeCreated.ToString('HH:mm:ss')
				}
			);
			$flag = 0;
		}
	}

	# 新規ファイルの書き込み
	$csvlist | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | % {$_ -replace '"',''} | Set-Content $file -Encoding UTF8;

# 自分のファイルが存在している場合は、追記のみ実施
} else {
	# CSVファイルの最終日時を取得
	$last = Get-Content $file | Select-Object -Last 1;
	$date = Get-Date $td.ToString('yyyy-' + $last.split(",")[0] + " " + $last.split(",")[2]);
	# 取得開始日時、終了日時設定（既存ファイルの最終時間から前日の24時まで）
	$sd = $date.AddMinutes(1);
	$date = Get-Date $td.ToString('yyyy-MM-dd 23:59:59');
	$ed = $date.AddDays(-1);
	# 起動終了時間の取得
	$array = GET-WinEvent -FilterHashTable @{LogName='System'; StartTime=$sd; EndTime=$ed} | Where-Object{$_.Id -eq 7001 -or $_.Id -eq 7002} | select-Object TimeCreated,Id,Message | Sort-Object -Property @{E = {Get-Date $_.TimeCreated}};

	# 日時データのみ成形
	$csvlist = @();
	$flag = 0;
	foreach ($data in $array) {
		if( $data.Id -eq "7001" ) {
			$attend = $data.TimeCreated.ToString('HH:mm:ss');
			$flag = 1;
		} elseif( $data.Id -eq "7002" -and $flag -eq 0 ) {
			$csvlist += @(
				[PSCustomObject]@{
					date = $data.TimeCreated.ToString('MM-dd')
					leav = ""
					attend = $data.TimeCreated.ToString('HH:mm:ss')
				}
			);
			$flag = 0;
		} elseif( $data.Id -eq "7002" -and $flag -eq 1 ) {
			$csvlist += @(
				[PSCustomObject]@{
					date = $data.TimeCreated.ToString('MM-dd')
					leav = $attend
					attend = $data.TimeCreated.ToString('HH:mm:ss')
				}
			);
			$flag = 0;
		}
	}
	# 既存ファイルへの追記
	$csvlist | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | % {$_ -replace '"',''} | Add-Content $file -Encoding UTF8;

}

# ここからサインアウトしていない場合の警告を表示するスクリプト
# 前日のサインイン、サインアウトを調べる基準日の算出
if( $td.DayOfWeek.value__ -eq 1 ) {
	# 月曜
	$sd = Get-Date $td.AddDays(-3).ToString('yyyy-MM-dd 00:00:00');
	$ed = Get-Date $td.AddDays(-3).ToString('yyyy-MM-dd 23:59:59');
} elseif( $td.DayOfWeek.value__ -eq 0 ) {
	# 日曜
	$sd = Get-Date $td.AddDays(-2).ToString('yyyy-MM-dd 00:00:00');
	$ed = Get-Date $td.AddDays(-2).ToString('yyyy-MM-dd 23:59:59');
} else {
	# それ以外
	$sd = Get-Date $td.AddDays(-1).ToString('yyyy-MM-dd 00:00:00');
	$ed = Get-Date $td.AddDays(-1).ToString('yyyy-MM-dd 23:59:59');
}

# 前日のデータを取得
$data = GET-WinEvent -FilterHashTable @{LogName='System'; StartTime=$sd; EndTime=$ed} | Where-Object{$_.Id -eq 6005 -or $_.Id -eq 6006 -or $_.Id -eq 6008 -or $_.Id -eq 7002 -or $_.Id -eq 7001} | select-Object TimeCreated,Id,Message;

# 最後のログオフデータを取得
$lastsd = Get-Date $td.AddDays(-7).ToString('yyyy-MM-dd 00:00:00');
$lasted = Get-Date $td.AddDays(-1).ToString('yyyy-MM-dd 23:59:59');
$lastlog = GET-WinEvent -FilterHashTable @{LogName='System'; StartTime=$lastsd; EndTime=$lasted} | Where-Object{$_.Id -eq 7002} | select-Object -First 1 TimeCreated;
$lastlog = $lastlog.TimeCreated;

# 前日のデータを元に警告を表示
if( $data.Count -ne 0 ) {
	$attend = $data | select @{L = "TimeCreated"; E = {Get-Date $_.TimeCreated}},Id | sort TimeCreated -Descending | where {$_.Id -eq 7001};
	$leav = $data | select @{L = "TimeCreated"; E = {Get-Date $_.TimeCreated}},Id | sort TimeCreated | where {$_.Id -eq 7002};
	if( -not ($leav) ) {
		# 前日サインアウトなし警告表示
		$result = $wsobj.popup("前日のサインアウト記録がありません。`r`n`r`n業務終了後には必ずサインアウトを実施してください。`r`n`r`n最終ログオフ日時：$lastlog",0,"サインアウト警告");
	}
	if( -not ($attend) ) {
		# 前日サインインなし警告表示
		$result = $wsobj.popup("前日のサインイン記録がありません。`r`n`r`n業務終了後には必ずサインアウトを実施してください。",0,"サインイン警告");
	}
} else {
	# 前日の記録なし、土日祝日を判定
	$date = $sd.ToString('yyyyMMdd');

  # パスをセット
	$path = Join-Path  $PSScriptRoot check_holiday.ps1;

	# 土日祝日判定スクリプトを実行
	."$path" $date;

	if( $LastExitCode -eq 1 ) {
		# 前日の記録なし、休みだったかも
		$result = $wsobj.popup("前日のサインイン、サインアウト記録がありません。`r`n`r`n業務終了後には必ずサインアウトを実施してください。`r`n`r`n休みだった場合はこのメッセージを無視してください。`r`n`r`n最終ログオフ日時：$lastlog",0,"前日の記録なし警告");
	}
}

# システムの状況を調査
$systemChk = Join-Path $PSScriptRoot system_chk.ps1;
& $systemChk;
