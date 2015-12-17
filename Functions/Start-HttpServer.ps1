Function Start-HttpServer
{
<#
    .Synopsis
        Starts the http server to return Windows Performance Counters.

    .Description
        Starts the http server to return Windows Performance Counters.

    .Parameter Verbose
        Provides Verbose output which is useful for troubleshooting

    .Parameter ExcludePerfCounters
        Excludes Performance counters defined in XML config

    .Parameter SqlMetrics
        Includes SQL Metrics defined in XML config

    .Parameter Url
        Specified the listening http prefix. Default http://localhost:8080/

    .Example
        PS> Start-StatHttpServer -Url http://*:8080/ (the last slash is important)

        Will start http server for specified prefix
        
        GET http://localhost:8080/cpu.usage returns cpu.usage current value

    .Notes
        NAME:      Start-HttpServer
        AUTHOR:    Alexey Larkov
#>
    [CmdletBinding()]
    Param
    (
        # Enable Test Mode. Metrics will not be sent to Graphite
        [String]$Url = "http://localhost:8080/",
        [switch]$ExcludePerfCounters = $false,
        [switch]$SqlMetrics = $false
    )

    # Run The Load XML Config Function
    $Config = Import-XMLConfig -ConfigPath $configPath
    $Config.HostName = "localhost"

    $configFileLastWrite = (Get-Item -Path $configPath).LastWriteTime

    if($ExcludePerfCounters -and -not $SqlMetrics) {
        throw "Parameter combination provided will prevent any metrics from being collected"
    }

    if($SqlMetrics) {
        if ($Config.MSSQLServers.Length -gt 0)
        {
            # Check for SQLPS Module
            if (($listofSQLModules = Get-Module -List SQLPS).Length -eq 1)
            {
                # Load The SQL Module
                Import-Module SQLPS -DisableNameChecking
            }
            # Check for the PS SQL SnapIn
            elseif ((Test-Path ($env:ProgramFiles + '\Microsoft SQL Server\100\Tools\Binn\Microsoft.SqlServer.Management.PSProvider.dll')) `
                -or (Test-Path ($env:ProgramFiles + ' (x86)' + '\Microsoft SQL Server\100\Tools\Binn\Microsoft.SqlServer.Management.PSProvider.dll')))
            {
                # Load The SQL SnapIns
                Add-PSSnapin SqlServerCmdletSnapin100
                Add-PSSnapin SqlServerProviderSnapin100
            }
            # If No Snapin's Found end the function
            else
            {
                throw "Unable to find any SQL CmdLets. Please install them and try again."
            }
        }
        else
        {
            Write-Warning "There are no SQL Servers in your configuration file. No SQL metrics will be collected."
        }
    }

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add($Url)
    $listener.Start()

    Write-Host "Listening at $Url..."
    
    $metrics = @{}
    $lastCollected = Get-Date -Date 0
    
    # Start Endless Loop
    while ($listener.IsListening)
    {
        $context = $listener.GetContext()
        $requestUrl = $context.Request.Url
        $response = $context.Response

        Write-Verbose $requestUrl
    
        $nowUtc = [datetime]::UtcNow
        
        if($requestUrl.LocalPath -eq "/kill"){
            $response.StatusCode = 200
            $response.Close()
            break
        }
        
        # Loop until enough time has passed to run the process again.
        if($nowUtc.AddSeconds(-$Config.MetricSendIntervalSeconds) -gt $lastCollected) {
            # Used to track execution time
            $collectionTime = [System.Diagnostics.Stopwatch]::StartNew()

            $collectMetricsParams = @{
                "Config" = $Config
                "ExcludePerfCounters" = $ExcludePerfCounters
                "SqlMetrics" = $SqlMetrics
                "AddConfigMetricPath" = $false
            }

            $metrics = CollectMetrics @collectMetricsParams
    
            $lastCollected = [datetime]::UtcNow
            if ($Config.ShowOutput)
            {
                # Write To Console How Long Execution Took
                $VerboseOutPut = 'Collection metrics time: ' + $collectionTime.Elapsed.TotalSeconds + ' seconds'
                Write-Verbose $VerboseOutPut
            }
        }
        
        $metricKey = "/" + $env:COMPUTERNAME + "." + $requestUrl.LocalPath.Substring(1)
        $metric = $metrics.Get_Item($metricKey)
    
        if ($metric -eq $null)
        {
            $response.StatusCode = 404
            
            $metrics.Keys | % {
                $VerboseOutPut = $_ + ": " + $metrics.Item($_)
                Write-Verbose $VerboseOutPut
            }
        }
        else
        {
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($metric)
            $response.ContentLength64 = $buffer.Length
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
        }
        
        $response.Close()
    
        $responseStatus = $response.StatusCode
        Write-Verbose "$responseStatus"

    }
    $listener.Stop()
    $listener.Close()
}