#####################################################################
# システム情報
#####################################################################
function QueryServerInfo(){
	# 端末のユーザー名を取得
	$username = $env:UserName;
	# 今日の日付
	$td = (Get-Date);

	# オブジェクトを作成
	$ReturnData = New-Object PSObject | Select-Object HostName,Manufacturer,Model,SN,CPUName,PhysicalCores,Sockets,MemorySize,DiskInfos,OS,Install

	# 基本的な情報を取得
	$Win32_BIOS = Get-WmiObject Win32_BIOS
	$Win32_Processor = Get-WmiObject Win32_Processor
	$Win32_ComputerSystem = Get-WmiObject Win32_ComputerSystem
	$Win32_OperatingSystem = Get-WmiObject Win32_OperatingSystem

	# ホスト名
	$ReturnData.HostName = hostname

	# メーカー名
	$ReturnData.Manufacturer = $Win32_BIOS.Manufacturer

	# モデル名
	$ReturnData.Model = $Win32_ComputerSystem.Model

	# シリアル番号
	$ReturnData.SN = $Win32_BIOS.SerialNumber

	# CPU 名
	$ReturnData.CPUName = @($Win32_Processor.Name)[0]

	# 物理コア数
	$PhysicalCores = 0
	$Win32_Processor.NumberOfCores | % { $PhysicalCores += $_}
	$ReturnData.PhysicalCores = $PhysicalCores
	
	# ソケット数
	$ReturnData.Sockets = $Win32_ComputerSystem.NumberOfProcessors
	
	# メモリーサイズ(GB)
	$Total = 0
	Get-WmiObject -Class Win32_PhysicalMemory | % {$Total += $_.Capacity}
	$ReturnData.MemorySize = [int]($Total/1GB)
	
	# ディスク情報
	[array]$DiskDrives = Get-WmiObject Win32_DiskDrive | ? {$_.Caption -notmatch "Msft"} | sort Index
	$DiskInfos = @()
	foreach( $DiskDrive in $DiskDrives ){
		$DiskInfo = New-Object PSObject | Select-Object Index, DiskSize
		$DiskInfo.Index = $DiskDrive.Index			  # ディスク番号
		$DiskInfo.DiskSize = [int]($DiskDrive.Size/1GB) # ディスクサイズ(GB)
		$DiskInfos += $DiskInfo
	}
	$ReturnData.DiskInfos = $($DiskInfos)
	
	# OS情報
	$OS = $Win32_OperatingSystem.Caption
	$SP = $Win32_OperatingSystem.ServicePackMajorVersion
	if( $SP -ne 0 ){ $OS += "SP" + $SP }
	$build = $Win32_OperatingSystem.BuildNumber
	$ret = switch ($build) {
		22000 {"21H2"; break}
		19044 {"21H2"; break}
		19043 {"21H1"; break}
		19042 {"20H2"; break}
		19041 {"2004"; break}
		18363 {"1909"; break}
		18362 {"1903"; break}
		17763 {"1809"; break}
		17134 {"1803"; break}
		default {0}
	}
	if( $ret -ne 0 ){ $OS += " " + $ret }
	$ReturnData.OS = $OS
	
	# インストール済みソフトウェア情報
	$path1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	$path2 = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	$path3 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
	[array]$lists = Get-ChildItem -Path ($path1,$path2,$path3) | %{Get-ItemProperty $_.PsPath} | ?{$_.systemcomponent -ne 1 -and $_.parentkeyname -eq $null -and $_.DisplayName -ne $null}
	$installs = @()
	foreach( $list in $lists ){
		$install = New-Object PSObject | Select-Object DisplayName, Publisher, DisplayVersion
		$install.DisplayName = $list.DisplayName
		$install.Publisher = $list.Publisher
		$install.DisplayVersion = $list.DisplayVersion
		$installs += $install
	}
	$ReturnData.Install = $($installs)

	# Server パスをセット
	$serverPath = "\file-server\system_log"

	# OneDrive 上のパスをセット
	$drivePath = [Environment]::GetEnvironmentVariable("OneDriveCommercial", "User")

	# パスをセット
	$path = Join-Path $drivePath $serverPath

	# ファイル保存先パスを生成
	$file_name = $username + ".json"
	$file = Join-Path $path $file_name

	# JSON データの生成
	$ReturnJson = ConvertTo-Json $ReturnData
	#$ReturnJson

	# ファイルの書き込み
	$ReturnJson | Out-File $file -Encoding utf8

	return
}

QueryServerInfo
