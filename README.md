# Graphite PowerShell Functions

A group of PowerShell functions that allow you to send Windows Performance counters to a Graphite Server, all configurable from a simple XML file.

[![GitHub Version](https://img.shields.io/github/release/MattHodge/Graphite-PowerShell-Functions.svg)](https://github.com/MattHodge/Graphite-PowerShell-Functions/releases)

More details at [http://www.hodgkins.net.au/mswindows/using-powershell-to-send-metrics-graphite/](http://www.hodgkins.net.au/mswindows/using-powershell-to-send-metrics-graphite/)

## Features

* Sends Metrics to Graphite's Carbon daemon using TCP or UDP
* Can collect Windows Performance Counters
* Can collect values by using T-SQL queries against MS SQL databases
* Converts time to UTC on sending
* All configuration can be done from a simple XML file
* Allows you to override the hostname in Windows Performance Counters before sending on to Graphite
* Allows renaming of metric names using regex via the configuration file
* Reloads the XML configuration file automatically. For example, if more counters are added to the configuration file, the script will notice and start sending metrics for them to Graphite in the next send interval
* Additional functions are exposed that allow you to send data to Graphite from PowerShell easily. [Here](#functions) is the list of included functions
* Script can be installed to run as a service
* Installable by Chef Cookbook [which is available here](https://github.com/tas50/chef-graphite_powershell_functions/)
* Installable by Puppet [which is available here](https://forge.puppetlabs.com/opentable/graphite_powershell)
* Supports Hosted Graphite [which is available here](https://www.hostedgraphite.com)

## Installation

	Set-ExecutionPolicy Bypass -Scope Process -Force; Install_Graphite-Powershell.ps1

Change StatsToGraphiteConfig.xml, more in the wiki https://github.com/skbkontur/Graphite-PowerShell-Functions/wiki

    Start-Service Graphite-Powershell