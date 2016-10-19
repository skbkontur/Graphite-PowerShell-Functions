Function Start-StatsToGraphite
{
<#
    .Synopsis
        Starts the loop which sends Windows Performance Counters to Graphite.

    .Description
        Starts the loop which sends Windows Performance Counters to Graphite. Configuration is all done from the StatsToGraphiteConfig.xml file.

    .Parameter Verbose
        Provides Verbose output which is useful for troubleshooting

    .Parameter TestMode
        Metrics that would be sent to Graphite is shown, without sending the metric on to Graphite.

    .Parameter ExcludePerfCounters
        Excludes Performance counters defined in XML config

    .Parameter SqlMetrics
        Includes SQL Metrics defined in XML config

    .Example
        PS> Start-StatsToGraphite

        Will start the endless loop to send stats to Graphite

    .Example
        PS> Start-StatsToGraphite -Verbose

        Will start the endless loop to send stats to Graphite and provide Verbose output.

    .Example
        PS> Start-StatsToGraphite -SqlMetrics

        Sends perf counters & sql metrics

    .Example
        PS> Start-StatsToGraphite -SqlMetrics -ExcludePerfCounters

        Sends only sql metrics

    .Notes
        NAME:      Start-StatsToGraphite
        AUTHOR:    Matthew Hodgkins
        WEBSITE:   http://www.hodgkins.net.au
#>
    [CmdletBinding()]
    Param
    (
        # Enable Test Mode. Metrics will not be sent to Graphite
        [Parameter(Mandatory = $false)]
        [switch]$TestMode,
        [switch]$ExcludePerfCounters = $false,
        [switch]$SqlMetrics = $false,
        [switch]$NtpOffset = $true
    )

    # Run The Load XML Config Function
    $Config = Import-XMLConfig -ConfigPath $configPath

    # Get Last Run Time
    $sleep = 0

    $configFileLastWrite = (Get-Item -Path $configPath).LastWriteTime

    if($ExcludePerfCounters -and -not $SqlMetrics -and -not $NtpOffset) {
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

    # Start Endless Loop
    while ($true)
    {
        # Loop until enough time has passed to run the process again.
        if($sleep -gt 0) {
            Start-Sleep -Milliseconds $sleep
        }

        # Used to track execution time
        $iterationStopWatch = [System.Diagnostics.Stopwatch]::StartNew()

        $nowUtc = [datetime]::UtcNow

        # Round Time to Nearest Time Period
        $nowUtc = $nowUtc.AddSeconds(- ($nowUtc.Second % $Config.MetricSendIntervalSeconds))

        $collectMetricsParams = @{
            "Config" = $Config
            "ExcludePerfCounters" = $ExcludePerfCounters
            "SqlMetrics" = $SqlMetrics
            "AddConfigMetricPath" = $true
            "NtpOffset" = $NtpOffset
        }
        $metricsToSend = CollectMetrics @collectMetricsParams

        # Send To Graphite Server

        $sendBulkGraphiteMetricsParams = @{
            "CarbonServer" = $Config.CarbonServer
            "CarbonServerPort" = $Config.CarbonServerPort
            "Metrics" = $metricsToSend
            "DateTime" = $nowUtc
            "UDP" = $Config.SendUsingUDP
            "Verbose" = $Config.ShowOutput
            "TestMode" = $TestMode
        }

        Send-BulkGraphiteMetrics @sendBulkGraphiteMetricsParams

        # Reloads The Configuration File After the Loop so new counters can be added on the fly
        if((Get-Item $configPath).LastWriteTime -gt (Get-Date -Date $configFileLastWrite)) {
            $Config = Import-XMLConfig -ConfigPath $configPath
        }

        $iterationStopWatch.Stop()
        $collectionTime = $iterationStopWatch.Elapsed
        $sleep = $Config.MetricTimeSpan.TotalMilliseconds - $collectionTime.TotalMilliseconds
        if ($Config.ShowOutput)
        {
            # Write To Console How Long Execution Took
            $VerboseOutPut = 'PerfMon Job Execution Time: ' + $collectionTime.TotalSeconds + ' seconds'
            Write-Output $VerboseOutPut
        }
    }
}

function CollectMetrics
{

    param (
        [hashtable]$Config,
        [switch]$ExcludePerfCounters,
        [switch]$SqlMetrics,
        [bool]$AddConfigMetricPath,
        [switch]$NtpOffset
    )

    $metrics = @{}

    if(-not $ExcludePerfCounters)
    {
        # Take the Sample of the Counter
        $collections = Get-Counter -Counter $Config.Counters -SampleInterval 1 -MaxSamples 1
        # Filter the Output of the Counters
        $samples = $collections.CounterSamples

        # Verbose
        Write-Verbose "All Samples Collected"

        # Loop Through All The Counters
        foreach ($sample in $samples)
        {
            if ($Config.ShowOutput)
            {
                Write-Verbose "Sample Name: $($sample.Path)"
            }

            # Create Stopwatch for Filter Time Period
            $filterStopWatch = [System.Diagnostics.Stopwatch]::StartNew()

            # Check if there are filters or not
            if ([string]::IsNullOrWhiteSpace($Config.Filters) -or $sample.Path -notmatch [regex]$Config.Filters)
            {
                # Run the sample path through the ConvertTo-GraphiteMetric function
                $cleanNameOfSample = ConvertTo-GraphiteMetric -MetricToClean $sample.Path -HostName $Config.NodeHostName -MetricReplacementHash $Config.MetricReplace

                # Build the full metric path
                if($AddConfigMetricPath){
                    $metricPath = $Config.MetricPath + '.' + $cleanNameOfSample
                }else{
                    $metricPath = "/" + $cleanNameOfSample
                }

                $metrics[$metricPath] = $sample.Cookedvalue
            }
            else
            {
                Write-Verbose "Filtering out Sample Name: $($sample.Path) as it matches something in the filters."
            }

            $filterStopWatch.Stop()

            Write-Verbose "Job Execution Time To Get to Clean Metrics: $($filterStopWatch.Elapsed.TotalSeconds) seconds."

        }# End for each sample loop

        if ($Config.Services -ne $null)
        {
          Write-Verbose "Service monitor enabled"
          Foreach ($service in $Config.Services)
          {
            # Create Stopwatch for Filter Time Period
            $filterStopWatch = [System.Diagnostics.Stopwatch]::StartNew()

            $cleanNameOfService = ConvertTo-GraphiteMetric -MetricToClean $service -HostName $Config.NodeHostName -MetricReplacementHash $Config.MetricReplace
            $serviceStatus = Get-Service -Name $service -ErrorAction SilentlyContinue
            if ($serviceStatus -eq $null)
            {
              Write-Warning "Service $service not found."
              $status = 10
            }
            elseif ($serviceStatus.Status -eq [System.ServiceProcess.ServiceControllerStatus]::Running)
            {
              $status = 20
            }
            else
            {
              $status = 10
            }

            # Build the full metric path
            if($AddConfigMetricPath){
                $metricPath = $Config.MetricPath + '.' + $Config.NodeHostName.ToLower() + '.services.' + $cleanNameOfService
            }else{
                $metricPath = "/" + $Config.NodeHostName.ToLower() + '.services.' + $cleanNameOfService
            }

            $metrics[$metricPath] = $status

            $filterStopWatch.Stop()
            Write-Verbose "Job Execution Time To Get to Clean Metrics: $($filterStopWatch.Elapsed.TotalSeconds) seconds."
          }

        }
    }# end if Config.Services

    if ($Config.NtpTimeSource -ne $null) {
        $sample = & W32TM /stripchart /computer:$($Config.NtpTimeSource) /dataonly /period:1 /samples:1 2>&1 | Select-String '[+-]\d\d\.\d\d\d\d\d\d\d' -AllMatches | Foreach {$_.Matches} | Foreach {[decimal]$_.Value*1000}
        # Build the full metric path
        if ($sample) {
          $metricPath = $Config.MetricPath + '.' + $Config.NodeHostName.toLower() + '.' + $Config.NtpMetricPath
          $metrics[$metricPath] = $sample
        } else {
          Write-Verbose "Unable to test time-diff with $($Config.NtpTimeSource)"
        }
    }#endif Config.NtpTimeSource

    if($SqlMetrics) {
        # Loop through each SQL Server
        foreach ($sqlServer in $Config.MSSQLServers)
        {
            Write-Verbose "Running through SQLServer $($sqlServer.ServerInstance)"
            # Loop through each query for the SQL server
            foreach ($query in $sqlServer.Queries)
            {
                Write-Verbose "Current Query $($query.TSQL)"

                $sqlCmdParams = @{
                    'ServerInstance' = $sqlServer.ServerInstance;
                    'Database' = $query.Database;
                    'Query' = $query.TSQL;
                    'ConnectionTimeout' = $Config.MSSQLConnectTimeout;
                    'QueryTimeout' = $Config.MSSQLQueryTimeout
                }

                # Run the Invoke-SqlCmd Cmdlet with a username and password only if they are present in the config file
                if (-not [string]::IsNullOrWhitespace($sqlServer.Username) `
                    -and -not [string]::IsNullOrWhitespace($sqlServer.Password))
                {
                    $sqlCmdParams['Username'] = $sqlServer.Username
                    $sqlCmdParams['Password'] = $sqlServer.Password
                }

                # Run the SQL Command
                try
                {
                    $commandMeasurement = Measure-Command -Expression {
                        $sqlresult = Invoke-SQLCmd @sqlCmdParams

                        # Build the MetricPath that will be used to send the metric to Graphite
                        $metricPath = $Config.MSSQLMetricPath + '.' + $query.MetricName

                        $metrics[$metricPath] = $sqlresult[0]
                    }

                    Write-Verbose ('SQL Metric Collection Execution Time: ' + $commandMeasurement.TotalSeconds + ' seconds')
                }
                catch
                {
                    $exceptionText = GetPrettyProblem $_
                    throw "An error occurred with processing the SQL Query. $exceptionText"
                }
            } #end foreach Query
        } #end foreach SQL Server
    }#endif SqlMetrics

    # Send To Graphite Server
    Return $metrics
}
