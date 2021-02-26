##
# Parse-Files
# -----------
# This function iterates through the files of a specified directory, parsing each file and 
# exctracting relevant data from each file. Data is then packaged into an array of objects,
# one object per file, and returned along with some meta data.
#
# Parameters
# ----------
# DirName (String) - Path to the directory whose files are to be parsed
# Tool (String) - Name of the tool whose data is being parsed (NTTTCP, LATTE, CTStraffic, etc.)
#
# Return
# ------
# HashTable - Object containing an array of dataEntry objects and meta data
#
##
function Parse-Files {
    param (
        [Parameter(Mandatory=$true)] [string]$DirName, 
        [Parameter()] [string] $Tool
    )

    $files = Get-ChildItem $DirName

    if ($Tool -eq "NTTTCP") {
        [Array] $dataEntries = @()
        foreach ($file in $files) {
            $fileName = $file.FullName
            $dataEntry = Parse-NTTTCP -FileName $fileName
            if ($dataEntry) {
                $dataEntries += ,$dataEntry
            }
        }

        $rawData = @{
            "meta" = @{
                "units" = @{
                    "cycles"     = "cycles/byte"
                    "throughput" = "Gbps"
                }
                "goal" = @{
                    "throughput" = "increase"
                    "cycles"     = "decrease"
                }
                "format" = @{
                    "throughput" = "0.00"
                    "cycles"     = "0.00"
                    "% change"   = "+#.0%;-#.0%;0.0%"
                }
                "noTable"  = [Array] @("filename", "sessions", "bufferLen", "bufferCount")
            }
            "data" = $dataEntries
        }

        return $rawData
    } 
    elseif ($Tool -eq "LATTE") {
        [Array] $dataEntries = @() 
        foreach ($file in $files) {
            $fileName = $file.FullName
            $dataEntry = Parse-LATTE -FileName $fileName

            $dataEntries += ,$dataEntry
        }

        $rawData = @{
            "meta" = @{
                "units" = @{
                    "latency"  = "us"
                }
                "goal" = @{
                    "latency"  = "decrease"
                }
                "format" = @{
                    "latency"  = "#.0"
                    "% change" = "+#.0%;-#.0%;0.0%"
                }
                "noTable"  = [Array]@("filename", "sendMethod", "protocol")
            }
            "data" = $dataEntries
        }

        return $rawData
    }
    elseif ($Tool -eq "CTStraffic") {
        [Array] $dataEntries = @() 
        foreach ($file in $files) {
            $fileName = $file.FullName
            $ErrorActionPreference = "Stop"
            $dataEntry = Parse-CTStraffic -FileName $fileName

            if ($dataEntry) {
                $dataEntries += ,$dataEntry
            }
        }

        $rawData = @{
            "meta" = @{
                "units" = @{
                    "throughput" = "Gbps"
                }
                "goal" = @{
                    "throughput" = "increase"
                }
                "format" = @{
                    "throughput" = "0.00"
                    "% change"   = "+#.0%;-#.0%;0.0%"
                }
                "noTable"  = [Array]@("filename", "sessions")
            }
            "data" = [Array]$dataEntries
        }

        return $rawData
    }
}


##
# Parse-NTTTCP
# ------------
# This function parses a single file containing NTTTCP data in an XML format. Relevant data
# is then extracted, packaged into a dataEntry object, and returned.
#
# Parameters
# ----------
# Filename (String) - Path of file to be parsed
#
# Return
# ------
# HashTable - Object containing extracted data
#
## 
function Parse-NTTTCP ([string] $FileName) {
    if ($FileName.Split(".")[-1] -ne "xml"){
        return
    }

    [XML]$file = Get-Content $FileName

    if (-not $file) {
        Write-Warning "Unable to parse file $FileName"
        return
    }

    [Decimal] $cycles = $file.ChildNodes.cycles.'#text'
    [Decimal] $throughput = ($file.ChildNodes.throughput | where {$_.metric -eq "mbps"})."#text" / 1000
    [Int] $sessions = $file.ChildNodes.parameters.max_active_threads
    [Int] $bufferLen = $file.ChildNodes.bufferLen
    [Int] $bufferCount = $file.ChildNodes.io

    $dataEntry = @{
        "sessions"    = $sessions
        "throughput"  = $throughput
        "cycles"      = $cycles
        "filename"    = $FileName
        "bufferLen"   = $bufferLen
        "bufferCount" = $bufferCount
    }

    return $dataEntry
}

<#
.SYNOPSIS
    Parses CTSTraffic output.
.DESCRIPTION
    This function parses a CTStraffic status log file, generated from
    the -StatusFilename option. Desired data is collected and returned
    in a Hashtable.
.PARAMETER Filename
    Path of the status log file to parse.
#>
function Parse-CTStraffic {
    param(
        [String] $Filename
    )

    $data = (Get-Content $Filename) -replace '^"|"$','' | ConvertFrom-Csv

    if (-not ($data | Get-Member -Name "In-Flight")) {
        Write-Error "'$Filename' is not a valid ctsTraffic status log. Please verify that the log was generated by the -StatusFilename option."
        return
    }

    $bytesToGigabits = 8 / (1000 * 1000 * 1000)

    $throughput = ($data.SendBps | measure -Average).Average * $bytesToGigabits
    $maxSessions = ($data."In-Flight" | measure -Max).Max

    $dataEntry = @{
        "sessions"   = $maxSessions
        "throughput" = $throughput
        "filename"   = $Filename
    }

    return $dataEntry
}

##
# Parse-LATTE
# ----------
# This function parses a single file containing LATTE data. This function can parse files 
# containing either raw LATTE data, or a LATTA summary. For raw data, each line contains
# a latency sample which is extracted into an array, packaged into a dataEntry 
# object, and returned. For summary data, the latency histogram and a few other measures 
# are parsed from the file, packaged into a dataEntry object, and returned.
#
# Parameters
# ----------
# Filename (String) - Path of file to be parsed
#
# Return
# ------
# HashTable - Object containing extracted data
#
##
function Parse-LATTE ([string] $FileName) {
    $file = Get-Content $FileName

    $dataEntry = @{
        "filename" = $FileName
    }

    $splitline = Remove-EmptyStrings -Arr (([Array]$file)[0]).split(' ')
    if ($splitline[0] -eq "Protocol") {
        $histogram = $false

        foreach ($line in $file) {
            $splitLine = Remove-EmptyStrings -Arr $line.split(' ')

            if ($splitLine.Count -eq 0) {
                continue
            }

            if ($splitLine[0] -eq "Protocol") {
                $dataEntry.protocol = $splitLine[-1]
            }
            if ($splitLine[0] -eq "MsgSize") {
                $dataEntry.msgSize = $splitLine[-1]  # Not currently used for anything
            }

            if ($splitLine[0] -eq "Interval(usec)") {
                $dataEntry.latency = [HashTable] @{} 
                $histogram = $true
                continue
            }

            if ($histogram) {
                $dataEntry.latency.([Int32]$splitLine[0]) = [Int32] $splitLine[-1]
            }
        }

        if (-not $histogram) {
            Write-Warning "No histogram in file $filename"
            return
        }

    } 
    else {
        
        [Array] $latency = @()
        foreach ($line in $file) {
            if (-not ($line -match "\d+")) {
                Write-Warning "Error Parsing file $FileName"
                return
            }
            $latency += ,[int]$line
        }
        $dataEntry.latency = $latency
        $dataEntry.protocol = (($FileName.Split('\'))[-1].Split('.'))[0].ToUpper()
    }

    $dataEntry.sendMethod = (($FileName.Split('\'))[-1].Split('.'))[2]

    return $dataEntry
}


##
# Remove-EmptyStrings
# -------------------
# This function removes all empty strings from the given array
#
# Parameters
# ----------
# Arr (string[]) - Array of strings
# 
# Return
# ------
# Array of strings with all empty strings removed
#
##
function Remove-EmptyStrings ($Arr) {
    $newArr = [Array] @()
    foreach ($val in $arr) {
        if ($val -ne "") {
            $newArr += $val.Trim()
        }
    }
    return $newArr
}