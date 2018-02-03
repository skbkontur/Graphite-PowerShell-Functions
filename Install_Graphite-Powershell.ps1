Clear-Host
$Service = 'Graphite-PowerShell'
$Path = "$env:SystemDrive\$Service"
function DeleteGraphitePowerShell{
    param(
	[string]$Service = $Service,
	[string]$Path = $Path
	)
	
	#Removing the Graphite-PowerShell service
	if ((Get-Service -Name $Service -ErrorAction SilentlyContinue) -ne $null){
		Stop-Service -Name $Service -Force
		While ((Get-Service -Name $Service).Status -eq 'Stoped'){}
		Start-Process -FilePath "C:\Windows\System32\cmd.exe" -ArgumentList "/c sc delete $Service"
	}

	if(Test-Path -Path $Path){
		Remove-Item  -Path $Path -Recurse -Force
	}

	#Removing other Graphite-PowerShell modules
	$path_gp = (Get-Module -ListAvailable $Service).ModuleBase
	if ($path_gp -ne $null){
		Remove-Item -Path $path_gp  -Recurse -Force
	}
}

DeleteGraphitePowerShell

New-Item -Path $Path -ItemType 'directory'

#Download zip file and unzip
Add-Type -AssemblyName System.IO.Compression.FileSystem
function DownUnzip{
    param(
	[string]$url,
	[string]$zipfile,
	[string]$output
	)
	#Download
	try{
		Invoke-WebRequest -Uri $url -OutFile $zipfile
	}
	catch [System.Net.WebException] {
		$Request = $_.Exception
		Write-host "Exception caught: $Request"
		break
	}
	#Unzip
    try{
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $output)
    }
     catch [System.IO.IOException]{
        Write-Host 'The file already exists'
    }
}

$url = "https://codeload.github.com/skbkontur/Graphite-PowerShell-Functions/zip/master"
$zipfile = "$env:TEMP\Graphite-PowerShell-Functions-master.zip"
DownUnzip -url $url -zipfile $zipfile -output "$env:TEMP\"

$url = "https://nssm.cc/release/nssm-2.24.zip"
$zipfile = "$env:TEMP\nssm-2.24.zip"
DownUnzip -url $url -zipfile $zipfile -output "$env:TEMP\"

if(Test-Path -Path "$env:TEMP\$Service"){
    Remove-Item -Path "$env:TEMP\$Service" -Recurse -Force
}

Rename-Item -Path "$env:TEMP\Graphite-PowerShell-Functions-master" -NewName "$env:TEMP\$Service"
Copy-Item -Path "$env:TEMP\$Service" -Destination "$env:SystemDrive\"  -Recurse -Force

#Checking the size of registers
if(Test-Path -Path 'HKLM:\Software\Wow6432Node'){
	Copy-Item -Path "$env:TEMP\nssm-2.24\win64\nssm.exe" -Destination $Path -Recurse -Force
}
else{
	Copy-Item -Path "$env:TEMP\nssm-2.24\win64\nssm.exe" -Destination $Path -Recurse -Force
}

Remove-Item -Path "$env:TEMP\nssm*" -Recurse -Force
Remove-Item -Path "$env:TEMP\$Service*" -Recurse -Force

#Configure nssm
Start-Process -FilePath $Path\nssm.exe -ArgumentList "install $Service  powershell.exe -command & { Import-Module $Path\Graphite-PowerShell.psm1 ; Start-StatsToGraphite }"
Start-Process -FilePath "C:\Windows\System32\cmd.exe" -ArgumentList "/c sc failure $Service actions= restart/60000/restart/60000/restart/60000// reset= 240"
Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service  AppRotateFiles 1"
Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service  AppRotateOnline 1"
Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service  AppThrottle 1500"
Start-Process -FilePath $Path\nssm.exe -ArgumentList "start $Service"