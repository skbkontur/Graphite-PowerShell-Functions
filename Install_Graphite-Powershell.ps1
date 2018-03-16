Clear-Host
function DeleteGraphitePowerShell {
    $Service = 'Graphite-PowerShell'
    $Path = "$env:SystemDrive\$Service"
       
    if ($null -ne (Get-Service -Name $Service -ErrorAction SilentlyContinue)) {
        CheckingStateGP
    }  
   
    if (Test-Path -Path $Path) {
        CheckingStateNSSM
    }
}
 
function CheckingStateGP {
    $Timer = 10
    $MaxAttempts = 5
    $count = 0
    $Service = 'Graphite-PowerShell'
 
    if ((Get-Service -Name $Service).Status -eq 'Stopped') {
        Start-Process -FilePath "C:\Windows\System32\cmd.exe" -ArgumentList "/c sc delete $Service"
        return
    }
    else {
        Stop-Service -Name $Service -Force
        while ((Get-Service -Name $Service).Status -ne 'Stopped') {
            if (count -lt $MaxAttempts) {
                $count++
                Start-Sleep -seconds $timer                
            }
            else {
                $exception = New-Object -TypeName System.Exception -ArgumentList "Service $Service can not be stopped. Increase the value of the timer and try again." 
                throw $exception
            }                
        }
        Start-Process -FilePath "C:\Windows\System32\cmd.exe" -ArgumentList "/c sc delete $Service" -Wait      
    }        
}
 
function CheckingStateNSSM {
    $Timer = 10
    $MaxAttempts = 5
    $count = 0
    $Process = 'nssm'
    $Service = 'Graphite-PowerShell'
    $Path = "$env:SystemDrive\$Service"
    
    # The process nssm works
    if ($null -ne ((Get-Process -Name $Process -ErrorAction SilentlyContinue))) {
        $all_nssm = (Get-Process -Name $Process | Select-Object -ExpandProperty Path).count

        foreach ($nssm_proc in $all_nssm) {
            # nssm.exe is located on path
            if (((Get-Process -Name $Process | Select-Object -ExpandProperty Path)[$nssm_proc]) -eq "$Path\nssm.exe") { 
                $nssm_id = ((Get-Process -Name $Process | Select-Object -ExpandProperty Id)[$nssm_proc])
                Stop-Process -Id  $id -Force

                while ($null -ne $nssm_id) {                         
                    if ($count -lt $MaxAttempts) {
                        $count++
                        Start-Sleep -seconds $timer
                    }
                    else {
                        $exception = New-Object -TypeName System.Exception -ArgumentList "Process $Process can not be stopped. Increase the value of the timer and try again." 
                        throw $exception
                    }
                }
            }
        }
    }
    Remove-Item  -Path $Path -Recurse -Force
}

function DownUnzip {
    param(
        [string]$url,
        [string]$zipfile,
        [string]$output
    )
    # Download
    try {
        Invoke-WebRequest -Uri $url -OutFile $zipfile
    }
    catch [System.Net.WebException] {
        $Request = $_.Exception
        Write-Output "Exception caught: $Request"
        break
    }
    # Unzip
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $output)
    }
    catch [System.IO.IOException] {
        Write-Output 'The file already exists'
    }
}

function InstallGraphitePowerShell {
    $Service = 'Graphite-PowerShell'
    $Path = "$env:SystemDrive\$Service"
    $ver_nssm ='2.24'

    try {
        DeleteGraphitePowerShell
    }
    catch {
        $_.Exception.Message
        return
    }
    New-Item -Path $Path -ItemType 'directory'
 
    # Download zip file and unzip
    Add-Type -AssemblyName System.IO.Compression.FileSystem
 
    $url = "https://codeload.github.com/skbkontur/Graphite-PowerShell-Functions/zip/master"
    $zipfile = "$env:TEMP\Graphite-PowerShell-Functions-master.zip"
    DownUnzip -url $url -zipfile $zipfile -output "$env:TEMP\"
 
    $url = "https://nssm.cc/release/nssm-$ver_nssm.zip"
    $zipfile = "$env:TEMP\nssm-$ver_nssm.zip"
    DownUnzip -url $url -zipfile $zipfile -output "$env:TEMP\"
 
    if (Test-Path -Path "$env:TEMP\$Service") {
        Remove-Item -Path "$env:TEMP\$Service" -Recurse -Force
    }
 
    Rename-Item -Path "$env:TEMP\Graphite-PowerShell-Functions-master" -NewName "$env:TEMP\$Service"
    Copy-Item -Path "$env:TEMP\$Service" -Destination "$env:SystemDrive\"  -Recurse -Force
 
    # Checking the size of registers
    if (Test-Path -Path 'HKLM:\Software\Wow6432Node') {
        Copy-Item -Path "$env:TEMP\nssm-$ver_nssm\win64\nssm.exe" -Destination $Path -Recurse -Force
    }
    else {
        Copy-Item -Path "$env:TEMP\nssm-$ver_nssm\win64\nssm.exe" -Destination $Path -Recurse -Force
    }
 
    Remove-Item -Path "$env:TEMP\nssm*" -Recurse -Force
    Remove-Item -Path "$env:TEMP\$Service*" -Recurse -Force
 
    # Configure nssm
    Start-Process -FilePath $Path\nssm.exe -ArgumentList "install $Service  powershell.exe -command & { Import-Module $Path\Graphite-PowerShell.psm1 ; Start-StatsToGraphite }" -Wait
    Start-Process -FilePath "C:\Windows\System32\cmd.exe" -ArgumentList "/c sc failure $Service actions= restart/60000/restart/60000/restart/60000// reset= 240"
    Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service  AppRotateFiles 1"
    Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service  AppRotateOnline 1"
    Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service  AppThrottle 1500"
    Start-Process -FilePath $Path\nssm.exe -ArgumentList "start $Service"  
}
InstallGraphitePowerShell
