Set-Location -Path $env:SystemDrive\
$Service = 'Graphite-PowerShell'
$Path = "$env:SystemDrive\$Service"

#Removing the Graphite-PowerShell service
if ((Get-Service -Name $Service -ErrorAction SilentlyContinue) -ne $null){
	cmd.exe /c "$Path\nssm.exe stop  $Service"
	cmd.exe /c "sc delete $Service"
}

if(Test-Path -Path $Path){
    Remove-Item  -Path $Path -Recurse -Force
}

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
Set-Location -Path $Path
 .\nssm install $Service  powershell.exe -command "& { Import-Module $Path\Graphite-PowerShell.psm1 ; Start-StatsToGraphite }"
cmd.exe /c "sc failure $Service actions= restart/60000/restart/60000/restart/60000// reset= 240"
 .\nssm set  $Service  AppRotateFiles 1
 .\nssm set  $Service  AppRotateOnline 1
 .\nssm set $Service  AppThrottle 1500
 .\nssm start $Service 
