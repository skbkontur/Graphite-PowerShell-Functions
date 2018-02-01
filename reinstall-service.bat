pushd C:\Graphite-Powershell
nssm stop Graphite-Powershell
nssm remove Graphite-Powershell confirm
nssm install Graphite-Powershell powershell.exe -command "& { Import-Module C:\Graphite-Powershell\Graphite-PowerShell.psm1 ; Start-StatsToGraphite }"
sc failure Graphite-Powershell actions= restart/60000/restart/60000/restart/60000// reset= 240
nssm set  Graphite-Powershell AppRotateFiles 1
nssm set  Graphite-Powershell AppRotateOnline 1
nssm set  Graphite-Powershell AppThrottle 1500
nssm start Graphite-Powershell
