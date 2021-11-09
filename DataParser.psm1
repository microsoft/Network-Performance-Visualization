<#
.SYNOPSIS
    This function parses each file in the specified directory for
    the given tool. The data is packaged into an array, one entry
    per file, along with meta data.
.PARAMETER DirName
    Path to the directory with the data files.
.PARAMETER Tool
    Name of the tool that generated the data. (NTTTCP, etc.)
.PARAMETER Mode
    Whether the given directory contains 'Baseline' or 'Test' data
.PARAMETER InnerPivot
    Name of inner pivot property
.PARAMETER OuterPivot
    Name of outer pivot property 
#>
function Get-RawData {
    param (
        [Parameter(Mandatory=$true)] [String] $DirName, 
        [Parameter()] [String] $Tool,
        [Parameter()] [String] $Mode="Baseline",
        [Parameter()] [String] $InnerPivot,
        [Parameter()] [String] $OuterPivot
    )

    $output = @{}
    $files = Get-ChildItem -File $DirName
    if ($files.Count -eq 0) {
        Throw "'$DirName' does not contain any data files."
    }

    switch ($Tool) {
        "NTTTCP" {
            $parseFunc = ${Function:Parse-NTTTCP}

            $output."meta" = @{
                "props" = [Array] @(
                    "throughput",
                    "cycles"
                )
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
                }
                "noTable"  = [Array] @("filename", "sessions", "bufferLen", "bufferCount")
            }
        }
        "LATTE" {
            $parseFunc = ${Function:Parse-LATTE}

            $output."meta" = @{
                "props" = [Array] @(
                    "latency"
                )
                "units" = @{
                    "latency"  = "us"
                }
                "goal" = @{
                    "latency"  = "decrease"
                }
                "format" = @{
                    "latency"  = "#.0"
                }
                "noTable"  = [Array]@("filename", "sendMethod", "protocol")
            }
        }
        "CTStraffic" {
            $parseFunc = ${Function:Parse-CTSTraffic}

            $output."meta" = @{
                "props" = [Array] @(
                    "throughput"
                )
                "units" = @{
                    "throughput" = "Gbps"
                }
                "goal" = @{
                    "throughput" = "increase"
                }
                "format" = @{
                    "throughput" = "0.00"
                }
                "noTable"  = [Array]@("filename", "sessions")
            }
        }
        "CPS" {
            $parseFunc = ${Function:Parse-CPS}

            $output."meta" = @{
                "props" = [Array] @(
                    "conn/s",
                    "close/s"
                )
                "units" = @{
                    "conn/s" = ""
                    "close/s" = ""  
                }
                "goal" = @{
                    "conn/s" = "increase"
                    "close/s" = "increase"    
                }
                "format" = @{ 
                    "conn/s" = "0.0"
                    "close/s" = "0.0"
                }
                "noTable" = [Array]@("filename")
            }
        }
    } 

    $PathCosts = @{}
    $InnerPivotKeys = @{}
    $OuterPivotKeys = @{}

    $id = if ($Mode -eq "Baseline") {0} else {1}
    $output.data = [Array]@()
    $addedPathCosts = $false
    for($i = 0; $i -lt $files.Count; $i++) { 
        Write-Progress -Activity "Parsing $($Mode) Data Files..." -Status "Parsing..." -Id $id -PercentComplete (100 * (($i) / $files.Count))
        $output.data += , (& $parseFunc -FileName $files[$i].FullName -InnerPivot $InnerPivot -OuterPivot $OuterPivot `
                            -InnerPivotKeys $InnerPivotKeys  -OuterPivotKeys $OuterPivotKeys -PathCosts $PathCosts) 
        if (($PathCosts.Count -gt 0) -and (-not $addedPathCosts)) {
            $output.data = $output.data[0..($output.Count - 1)]
            $addedPathCosts = $true 
        }
    }

    if ($Tool -in @("CTStraffic", "NTTTCP")) {
        if ($PathCosts.Count -gt 0) {
            # This can be expanded to include the other metrics captured by the pathcosts tool
            Incorporate-PathCosts -Data $output.data -PathCosts $PathCosts 
            $output.meta.props += "cycles/byte"
            $output.meta.goal["cycles/byte"] = "decrease"
            $output.meta.format["cycles/byte"] = "0.00"
        }
        
    }

    Write-Progress -Activity "Parsing $($Mode) Data Files..." -Status "Done" -Id $id -PercentComplete 100
 
    if ($output."data".Count -eq 0) {
        Write-Error "Failed to parse any file in '$DirName'."
    }

    $output.meta.innerPivotKeys = $InnerPivotKeys.Keys 
    $output.meta.outerPivotKeys = $OuterPivotKeys.Keys 

    return $output
}


function Incorporate-PathCosts ($Data, $PathCosts) {
    foreach ($entry in $Data) {
        $file = $entry.filename.Split("\")[-1]
        if ($PathCosts.ContainsKey($file)) {
            # There is a bug in Get-VswitchPathCost where sometimes throughput counters return 0
            # and CPB calculation leads to a very large number.
            # This is to skip those outlier cases. 
            if ($PathCosts[$file]["Byte path cost (cycles/byte)"] -gt 1000) {
                continue
            }
            $entry["cycles/byte"] = $PathCosts[$file]["Byte path cost (cycles/byte)"]
        }
    }
}


<#
.SYNOPSIS
    Parses XML-formated NTTTCP output data file.
.PARAMETER FileName
    Path of file to be parsed.
.PARAMETER InnerPivotKeys
    Set containing all inner pivot keys encountered across all data files
.PARAMETER OuterPivotKeys
    Set containing all outer pivot keys encountered across all data files
.PARAMETER InnerPivot
    Name of inner pivot property
.PARAMETER OuterPivot
    Name of outer pivot property
#>
function Parse-NTTTCP ([String] $FileName, $InnerPivot, $OuterPivot, $InnerPivotKeys, $OuterPivotKeys, $PathCosts) {
    if ($Filename -match "pathcost") {
        Extract-PathCosts -Filename $Filename -PathCosts $PathCosts 
        return
    }
    
    if ($FileName -notlike "*.xml") {
        return
    }

    $file = (Get-Content $FileName) -as [XML]
    if (-not $file) {
        Write-Warning "Skipped '$FileName' because it is not valid XML."
        return
    }

    [Decimal] $cycles = $file.ChildNodes.cycles.'#text'
    [Decimal] $throughput = ($file.ChildNodes.throughput | where {$_.metric -eq "mbps"})."#text" / 1000
    [Int] $sessions = $file.ChildNodes.parameters.max_active_threads #should this be .num_processors or .parametes.max_active_threads?
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

    $iPivotKey = if ($dataEntry[$InnerPivot]) {$dataEntry[$InnerPivot]} else {""}
    $oPivotKey = if ($dataEntry[$OuterPivot]) {$dataEntry[$OuterPivot]} else {""}

    $InnerPivotKeys[$iPivotKey] = $true
    $OuterPivotKeys[$oPivotKey] = $true

    return $dataEntry
}


function Extract-PathCosts ($Filename, $PathCosts) {
    (Get-Content -Path $Filename | ConvertFrom-Json).psobject.properties | Foreach {
        $key = $_.Name     
        $values = @{}
        ($_.Value).psobject.properties | Foreach {
            $values[$_.Name] = [Decimal]$_.Value 
        }
        $PathCosts[$key] = $values
    }
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
.PARAMETER InnerPivotKeys
    Set containing all inner pivot keys encountered across all data files
.PARAMETER OuterPivotKeys
    Set containing all outer pivot keys encountered across all data files
.PARAMETER InnerPivot
    Name of inner pivot property
.PARAMETER OuterPivot
    Name of outer pivot property
#>
function Parse-CTStraffic ( [String] $Filename, $InnerPivot, $OuterPivot , $InnerPivotKeys, $OuterPivotKeys, $PathCosts) {

    if ($Filename -match "pathcost") {
        Extract-PathCosts -Filename $Filename -PathCosts $PathCosts 
        return
    }

    $data = (Get-Content $Filename) -replace '^"|"$','' | ConvertFrom-Csv



    if (-not ($data | Get-Member -Name "In-Flight" -ErrorAction "SilentlyContinue")) {
        Write-Warning "Skipped '$Filename' because it's not a valid ctsTraffic status log. Please verify that it was generated by the -StatusFilename option."
        return
    }

    $bytesToGigabits = [Decimal] 8 / (1000 * 1000 * 1000)

    $throughput = [Array]@()
    $warmupPadding = 2
    $cooldownPadding = 2

    for ($i = $warmupPadding; $i -lt $data.Count - $cooldownPadding; $i += 1) { 

        $tputVal = ($data[$i].SendBps, $data[$i].RecvBps | Measure-Object -Maximum).Maximum
        $throughput += [Decimal] $tputVal * $bytesToGigabits
    }  #($data.SendBps | measure -Average).Average * $bytesToGigabits

    $maxSessions = ($data."In-Flight" | measure -Max).Maximum

    $dataEntry = @{
        "sessions"   = $maxSessions
        "throughput" = $throughput
        "filename"   = $Filename
    }

    $iPivotKey = if ($dataEntry[$InnerPivot]) {$dataEntry[$InnerPivot]} else {""}
    $oPivotKey = if ($dataEntry[$OuterPivot]) {$dataEntry[$OuterPivot]} else {""}

    $InnerPivotKeys[$iPivotKey] = $true
    $OuterPivotKeys[$oPivotKey] = $true

    return $dataEntry
}

 
<#
.SYNOPSIS
    Parses LATTE output data files
.DESCRIPTION
    This function parses a CTStraffic status log file, generated from
    the -StatusFilename option. Desired data is collected and returned
    in a Hashtable.
.PARAMETER Filename
    Path of the status log file to parse.
.PARAMETER InnerPivotKeys
    Set containing all inner pivot keys encountered across all data files
.PARAMETER OuterPivotKeys
    Set containing all outer pivot keys encountered across all data files
.PARAMETER InnerPivot
    Name of inner pivot property
.PARAMETER OuterPivot
    Name of outer pivot property
#>
function Parse-LATTE ([string] $FileName, $InnerPivot, $OuterPivot, $InnerPivotKeys, $OuterPivotKeys, $PathCosts) {

    if ($Filename -match "pathcost") {
        Extract-PathCost -Filename $Filename -PathCosts $PathCosts 
        return
    }

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
                $dataEntry.latency[[Int32]$splitLine[0]] = [Int32] $splitLine[-1]
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

    $iPivotKey = if ($dataEntry[$InnerPivot]) {$dataEntry[$InnerPivot]} else {""}
    $oPivotKey = if ($dataEntry[$OuterPivot]) {$dataEntry[$OuterPivot]} else {""}

    $InnerPivotKeys[$iPivotKey] = $true
    $OuterPivotKeys[$oPivotKey] = $true
    return $dataEntry
}
 

<#
.SYNOPSIS 
    Parses CPS output data files
.DESCRIPTION
    This function parses a single file containing CPS data. Each line contains
    conn/s and close/s samples which are extracted into arrays, packaged into a dataEntry 
    object, and returned. 
.PARAMETER Filename
    Path of the status log file to parse.
.PARAMETER InnerPivotKeys
    Set containing all inner pivot keys encountered across all data files
.PARAMETER OuterPivotKeys
    Set containing all outer pivot keys encountered across all data files
.PARAMETER InnerPivot
    Name of inner pivot property
.PARAMETER OuterPivot
    Name of outer pivot property
#>
function Parse-CPS ([string] $FileName, $InnerPivot, $OuterPivot, $InnerPivotKeys, $OuterPivotKeys, $PathCosts) {
    if ($Filename -match "pathcost") {
        Extract-PathCost -Filename $Filename -PathCosts $PathCosts 
        return
    }

    $file = Get-Content $FileName

    $dataEntry = @{
        "filename" = $FileName
        "conn/s" = [Array] @()
        "close/s" = [Array] @()
    }

    foreach ($line in $file[1..($file.Count - 1)]) {
        $splitLine = Remove-EmptyStrings -Arr $line.split(' ')
         
        if ($splitLine.Count -eq 0) {
            break
        }

        $dataEntry."conn/s" += ,[Decimal]($splitLine[5])
        $dataEntry."close/s" += ,[Decimal]($splitLine[6]) 
    } 

    $iPivotKey = if ($dataEntry[$InnerPivot]) {$dataEntry[$InnerPivot]} else {""}
    $oPivotKey = if ($dataEntry[$OuterPivot]) {$dataEntry[$OuterPivot]} else {""}

    $InnerPivotKeys[$iPivotKey] = $true
    $OuterPivotKeys[$oPivotKey] = $true
    
    return $dataEntry
}


<#
.SYNOPSIS
    This function removes all empty strings from the given array
.PARAMETER Arr
    Array of strings 
#>
function Remove-EmptyStrings ($Arr) {
    $newArr = [Array] @()
    foreach ($val in $arr) {
        if ($val -ne "") {
            $newArr += $val.Trim()
        }
    }
    return $newArr
}