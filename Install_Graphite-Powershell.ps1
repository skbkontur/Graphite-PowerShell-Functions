Clear-Host

$Service = 'Graphite-PowerShell'
$Path = "$env:SystemDrive\$Service"
$Process = 'nssm'

function DeleteGraphitePowerShell {
    if ($null -ne (Get-Service -Name $Service -ErrorAction SilentlyContinue)) {
        CheckingStateGP
    }
    if ($null -ne (Get-Process -Name $Process -ErrorAction SilentlyContinue)) {
        CheckingStateNSSM
    }
    if (Test-Path -Path $Path) {
        Remove-Item  -Path $Path -Recurse -Force
    }
}
 
function CheckingStateGP {
    $Timer = 10
    $MaxAttempts = 5
    $count = 0

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
                throw [GraphiteInstallationException] "Service $Service can not be stopped. Increase the value of the timer and try again."
            }                
        }
        Start-Process -FilePath "C:\Windows\System32\cmd.exe" -ArgumentList "/c sc delete $Service" -Wait      
    }        
}
 
function CheckingStateNSSM {
    $Timer = 10
    $MaxAttempts = 5
    $count = 0
    
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
                    $exception = New-Object -TypeName System.Exception -ArgumentList "Service $Service can not be stopped. Increase the value of the timer and try again." 
                    throw $exception
                }
            }

            # nssm_id process is stopped
            if ($null -eq $nssm_id) {
                Remove-Item  -Path $Path -Recurse -Force
            }
        }
        # nssm.exe in path not found
        else {
            Remove-Item  -Path $Path -Recurse -Force
        }
    }
}

function Unzip {
    param(
        [string]$zipfile,
        [string]$output
    )

    try {
       $shell = new-object -com shell.application
        $zip = $shell.NameSpace($zipfile)
        foreach ($item in $zip.items()) {
            $shell.Namespace($output).copyhere($item)
        }
    }
    catch  {
        $_.Exception.Message
        return
    }
}

function InstallGraphitePowerShell {
    $debugmod = $false
    
    try {
        DeleteGraphitePowerShell
    }
    catch  {
        $_.Exception.Message
        return
    }

    New-Item -Path $Path -ItemType 'directory'
 
    # Download zip file and unzip       
    $url = "https://codeload.github.com/skbkontur/Graphite-PowerShell-Functions/zip/master"
    $zipfile = "$env:TEMP\Graphite-PowerShell-Functions-master.zip"

    $WebClient = New-Object System.Net.WebClient
    $WebClient.DownloadFile($url, $zipfile)
    
    Unzip -zipfile $zipfile -output "$env:TEMP\"
 
    $url = "https://nssm.cc/release/nssm-2.24.zip"
    $zipfile = "$env:TEMP\nssm-2.24.zip"
    $WebClient.DownloadFile($url, $zipfile)

    Unzip -zipfile $zipfile -output "$env:TEMP\"
 
    if (Test-Path -Path "$env:TEMP\$Service") {
        Remove-Item -Path "$env:TEMP\$Service" -Recurse -Force
    }
 
    Rename-Item -Path "$env:TEMP\Graphite-PowerShell-Functions-master" -NewName "$env:TEMP\$Service"
    Copy-Item -Path "$env:TEMP\$Service" -Destination "$env:SystemDrive\"  -Recurse -Force
 
    # Checking the size of registers
    if (Test-Path -Path 'HKLM:\Software\Wow6432Node') {
        Copy-Item -Path "$env:TEMP\nssm-2.24\win64\nssm.exe" -Destination $Path -Recurse -Force
    }      
    else {
        Copy-Item -Path "$env:TEMP\nssm-2.24\win32\nssm.exe" -Destination $Path -Recurse -Force
    }
 
    Remove-Item -Path "$env:TEMP\nssm*" -Recurse -Force
    Remove-Item -Path "$env:TEMP\$Service*" -Recurse -Force
 
    # Configure nssm
    Start-Process -FilePath $Path\nssm.exe -ArgumentList "install $Service  powershell.exe -command & { Import-Module $Path\Graphite-PowerShell.psm1 ; Start-StatsToGraphite }" -Wait
    Start-Process -FilePath "C:\Windows\System32\cmd.exe" -ArgumentList "/c sc failure $Service actions= restart/60000/restart/60000/restart/60000// reset= 240"
    Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service  AppRotateFiles 1"
    Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service  AppRotateOnline 1"
    if ($debugmod){
        Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service AppStderr $Path\stdout.txt"
        Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service AppStdout $Path\stdout.txt"
    }
    Start-Process -FilePath $Path\nssm.exe -ArgumentList "set  $Service  AppThrottle 1500"
}

InstallGraphitePowerShell