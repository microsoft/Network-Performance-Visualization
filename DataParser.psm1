<#
.SYNOPSIS
    This function parses every file in a specified directory. The data from each individual
    file is represented through a HashSet. The Hashtables containing data for each file are 
    stored in an array and returned along with meta data.
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
        [Parameter()] [String] $OuterPivot,
        [Parameter()] [Int] $Warmup,
        [Parameter()] [Int] $Cooldown
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
                    "cycles/byte"
                )
                "units" = @{
                    "cycles/byte"     = "cycles/byte"
                    "throughput" = "Gbps"
                }
                "goal" = @{
                    "throughput" = "increase"
                    "cycles/byte"     = "decrease"
                }
                "format" = @{
                    "throughput" = "0.00"
                    "cycles/byte"     = "0.00"
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
        "LagScope" {
            $parseFunc = ${Function:Parse-LagScope}

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

    $PathCosts = @{
        "meta" = @{
            "props" = @()
        }
    }
    $InnerPivotKeys = @{}
    $OuterPivotKeys = @{}

    $id = if ($Mode -eq "Baseline") {0} else {1}
    $output.data = [Array]@()  
    for($i = 0; $i -lt $files.Count; $i++) { 
        Write-Progress -Activity "Parsing $($Mode) Data Files..." -Status "Parsing..." -Id $id -PercentComplete (100 * (($i) / $files.Count))
        $output.data += , (& $parseFunc -FileName $files[$i].FullName -InnerPivot $InnerPivot -OuterPivot $OuterPivot `
                            -InnerPivotKeys $InnerPivotKeys  -OuterPivotKeys $OuterPivotKeys -PathCosts $PathCosts -Warmup $Warmup -Cooldown $Cooldown)
        if (-not $output.data[-1]) {
            $output.data = $output.data[0..($output.data.Count - 1)] 
        } 
    }

    if ($Tool -in @("CTStraffic", "NTTTCP")) {
        if ($PathCosts.Count -gt 0) {
            # This can be expanded to include the other metrics captured by the pathcosts tool
            Incorporate-PathCosts -Data $output.data -PathCosts $PathCosts   

            if ($PathCosts.meta.props -contains "total root VP utilization") {
                $output.meta.props += "total root VP utilization"
                $output.meta.goal["total root VP utilization"] = "decrease"
                $output.meta.format["total root VP utilization"] = "0.00"
                $output.meta.units["total root VP utilization"] = "% Utilization"
            }
            
            if ($PathCosts.meta.props -contains "vSwitch root VP utilization") {
                $output.meta.props += "vSwitch root VP utilization"
                $output.meta.goal["vSwitch root VP utilization"] = "decrease"
                $output.meta.format["vSwitch root VP utilization"] = "0.00"
                $output.meta.units["vSwitch root VP utilization"] = "% Utilization"
            }

            if ($PathCosts.meta.props -contains "CPU Utilization") { 
                $output.meta.props += "CPU Utilization"
                $output.meta.goal["CPU Utlization"] = "decrease"
                $output.meta.format["CPU Utlization"] = "0.00"
                $output.meta.units["CPU Utlization"] = "% Utilization"
            }

            if ($PathCosts.meta.props -contains "cycles/packet") { 
                $output.meta.props += "cycles/packet"
                $output.meta.goal["cycles/packet"] = "decrease"
                $output.meta.format["cycles/packet"] = "0.00"
            }

            if ($PathCosts.meta.props -contains "cycles/byte") {
                $output.meta.props += "cycles/byte"
                $output.meta.goal["cycles/byte"] = "decrease"
                $output.meta.format["cycles/byte"] = "0.00"
            }

            if ($PathCosts.meta.props -contains "RSS Core Utilization") { 
                $output.meta.props += "RSS Core Utilization"
                $output.meta.goal["RSS Core Utilization"] = "decrease"
                $output.meta.format["RSS Core Utilization"] = "0.00"
                $output.meta.units["cpu utlization"] = "% Utilization"
            }

            if ($PathCosts.meta.props -contains "UVMS Data Core Utilization") { 
                $output.meta.props += "UVMS Data Core Utilization"
                $output.meta.goal["UVMS Data Core Utilization"] = "decrease"
                $output.meta.format["UVMS Data Core Utilization"] = "0.00"
                $output.meta.units["cpu utlization"] = "% Utilization"
            }
        }
        
    }

    Write-Progress -Activity "Parsing $($Mode) Data Files..." -Status "Done" -Id $id -PercentComplete 100
 
    if ($output."data".Count -eq 0) {
        Write-Error "Failed to parse any files in '$DirName'."
    }

    $output.meta.innerPivotKeys = $InnerPivotKeys.Keys 
    $output.meta.outerPivotKeys = $OuterPivotKeys.Keys 

    return $output
}

<#
.SYNOPSIS
    Reads data from HashTable containing pathcost data and writes the values to the DataEntry 
    objects corresponding to each individual file.

.PARAMETER Data
    Array of DataEntry objects which each correspond to a single data file.
.PARAMETER PathCosts
    Hashtable containing a mapping between data filenames and pathcosts data
#>
function Incorporate-PathCosts ($Data, $PathCosts) {
    foreach ($entry in $Data) {
        $file = $entry.filename.Split("\")[-1]
        if ($PathCosts.ContainsKey($file)) {
            
            $cpb = $PathCosts[$file]["cycles/byte"]
            $cpu = $PathCosts[$file]["CPU Utilization"]
            $cpp = $PathCosts[$file]["cycles/packet"]
            $trvp = $PathCosts[$file]["Total Root VP Utilization"]
            $vsrvp = $PathCosts[$file]["vSwitch Root VP Utilization"]
            $udcu = $PathCosts[$file]["UVMS Data Core Utilization"]
            $rcu = $PathCosts[$file]["RSS Core Utilization"]

            # TPUT measures from dedicated tools are more reliable than vswitch counters 
            # (which sometimes return 0 erroneously), thus we perform the cycles/byte calculation
            # using TPUT measures from ntttcp/ctstraffic when possible
            if ($PathCosts[$file].ContainsKey("Total CPU cycles used per second")) {
                $avgTput = ((1000 * 1000 * 1000) / 8) * ($entry["throughput"] | Measure-Object -Average).Average
                if ($avgTput -ne 0) {
                    $cpb = $PathCosts[$file]["Total CPU cycles used per second"] / $avgTput
                }

                if ("cycles/byte" -notin $PathCosts.meta.props) {
                    $PathCosts.meta.props += "cycles/byte"
                }

            }

            # Sometimes counters mess up and record nearly-zero for tput and it causes 
            # cycle/byte calculations to return extremely large numbers. These outliers
            # make visualizations unusable, thus we filter outliers here. We 
            # should look for a more sustainable solution in the future
            
            if ($cpb -lt 100 -and $cpb -gt 0) {
                $entry["cycles/byte"] = $cpb
            } 

            if ($vsrvp -and $vsrvp -le 100) { 
                $entry["vSwitch root VP utilization"] = $vsrvp
            }

            if ($udcu) { 
                $entry["UVMS Data Core Utilization"] = $udcu 
            }

            if ($rcu) { 
                $entry["RSS Core Utilization"] = $rcu 
            }

            if ($cpp -gt 0) {
                $entry["cycles/packet"] = $cpp
            }
            if ($cpu -gt 0) {
                $entry["CPU Utilization"] = $cpu
            }
            if ($trvp -gt 0) {
                $entry["total root VP utilization"] = $trvp
            }
            
        }
    }
}


<#
.SYNOPSIS
    Parses a single XML-formated NTTTCP output data file. Relevant data is collected and returned
    as a Hashtable.
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
function Parse-NTTTCP ([String] $FileName, $InnerPivot, $OuterPivot, $InnerPivotKeys, $OuterPivotKeys, $PathCosts, $Warmup, $Cooldown) {
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
    [Int] $sessions = [Int] $filename.split("\")[-1].split(".")[2].substring(1) #should this be .num_processors or .parametes.max_active_threads?
    [Int] $bufferLen = $file.ChildNodes.bufferLen
    [Int] $bufferCount = $file.ChildNodes.io

    $dataEntry = @{
        "sessions"    = $sessions
        "throughput"  = $throughput
        "cycles/byte" = $cycles
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

<#
.SYNOPSIS
    Parses a single pathcosts data file, extracts relevant data, and stores the data in a 
    HashTable.
.PARAMETER Filename
    Path of file to be parsed
.PARAMETER PathCosts
    Hashtable to which pathcosts data is written
#>
function Extract-PathCosts ($Filename, $PathCosts) {
    (Get-Content -Path $Filename | ConvertFrom-Json).psobject.properties | Foreach {
        $key = $_.Name     
        $values = @{}
        $obj = $_.Value 
        if ($_.Value[1]) {
            $obj = $_.Value[1]
        }
        $obj.psobject.properties | Foreach {
            try {
                $propName = $_.Name
                if ($propName -Match ".*\((.*)\).*") {
                    $propName = $Matches[1]
                }
                $value = [Decimal]$_.Value
                if ($value -gt 0) {
                    if ($propName -notin $PathCosts.meta.props) {
                        $PathCosts.meta.props += $propName
                    }
                    $values[$propName] = $value
                }
                 
            } catch {}
        }
        $PathCosts[$key] = $values
    }
}
 

<#
.SYNOPSIS
    This function parses a single CTStraffic status log file. Desired data is collected and returned
    as a Hashtable.
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
function Parse-CTStraffic ( [String] $Filename, $InnerPivot, $OuterPivot , $InnerPivotKeys, $OuterPivotKeys, $PathCosts, $Warmup, $Cooldown) {

    if ($Filename -match "pathcost") {
        Extract-PathCosts -Filename $Filename -PathCosts $PathCosts 
        return
    }

    $data = (Get-Content $Filename) -replace '^"|"$','' | ConvertFrom-Csv


    if (-not ($data | Get-Member -Name "In-Flight" -ErrorAction "SilentlyContinue")) {
        Write-Warning "Skipped '$Filename' because it's not a valid ctsTraffic status log. Please verify that it was generated by the -StatusFilename option."
        return
    }

    $fileSplit = $Filename.Split("\")[-1]
    $fileSplit = $fileSplit.Split("/")[-1]
    $sessionsStr = $fileSplit.Split(".")[2] 
    $bytesToGigabits = [Decimal] 8 / (1000 * 1000 * 1000)

    $throughput = [Array]@()
    $warmupPadding = $Warmup
    $dynamicWarmup = $true
    $dynamicCooldown = $true 
    $cooldownPadding = $Cooldown

    $avgSend = ($data.SendBps | Measure-Object -Average).Average
    $avgRecv = ($data.RecvBps | Measure-Object -Average).Average
    $tputProp = if ($avgSend -gt $avgRecv) {"SendBps"} else {"RecvBps"} 

    for ($i = $warmupPadding; $i -lt $data.Count - $cooldownPadding; $i += 1) { 

        $tputVal = $data[$i].$tputProp
        if ($tputVal -eq 0 -and $dynamicWarmup) {
            continue
        } else {
            $dynamicWarmup = $false
        }
        $throughput += [Decimal] $tputVal * $bytesToGigabits
    }  #($data.SendBps | measure -Average).Average * $bytesToGigabits

    if ($dynamicCooldown) {
        for ($j = $throughput.Length - 1; $j -ge 0; $j -= 1) {
            if ($throughput[$j] -eq 0) {
                $throughput = $throughput[0..($j-1)] 
            } else {
                break
            }
        }
    }

    $maxSessions = ($data."In-Flight" | Measure-Object -Max).Maximum
    
    $sessionsStr = $fileSplit.Split(".")[2]
    $dataEntry = @{
        "sessions"   =  [Int] $maxSessions
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
    This function parses a single LATTE data file. Relevant data is collected and returned
    as a Hashtable.
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
function Parse-LATTE ([string] $FileName, $InnerPivot, $OuterPivot, $InnerPivotKeys, $OuterPivotKeys, $PathCosts, $Warmup, $Cooldown) {

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
        $dataEntry.latency = $latency[$Warmup..($latency.Count - ($Cooldown + 1))]
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
    This function parses a single LagScope data file. Relevant data is collected and returned
    as a Hashtable. 
.PARAMETER Filename
    Path of the data file to parse.
.PARAMETER InnerPivotKeys
    Set containing all inner pivot keys encountered across all data files
.PARAMETER OuterPivotKeys
    Set containing all outer pivot keys encountered across all data files
.PARAMETER InnerPivot
    Name of inner pivot property
.PARAMETER OuterPivot
    Name of outer pivot property
#>
function Parse-LagScope ([string] $FileName, $InnerPivot, $OuterPivot, $InnerPivotKeys, $OuterPivotKeys, $PathCosts, $Warmup, $Cooldown) {
    
    if ($Filename -match "pathcost") {
        Extract-PathCost -Filename $Filename -PathCosts $PathCosts 
        return
    }

    if ($FileName -notlike "*.json") {return}
    
    $latencies = (Get-Content $Filename | ConvertFrom-Json).latencies 
    $dataEntry = @{
        "filename" = $FileName
        "latency" = [Array] @()
    }

    for ($i = 0; $i -lt $latencies.Count; $i++) {
        $dataEntry.latency += @($latencies[$i].latency) * $latencies[$i].frequency
    }
    
    $iPivotKey = if ($dataEntry[$InnerPivot]) {$dataEntry[$InnerPivot]} else {""}
    $oPivotKey = if ($dataEntry[$OuterPivot]) {$dataEntry[$OuterPivot]} else {""}

    $InnerPivotKeys[$iPivotKey] = $true
    $OuterPivotKeys[$oPivotKey] = $true

    return $dataEntry
} 
 

<#
.SYNOPSIS 
    This function parses a single file containing CPS data. Each line contains
    conn/s and close/s samples which are extracted into arrays, packaged into a 
    HashTable, and returned. 
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
function Parse-CPS ([string] $FileName, $InnerPivot, $OuterPivot, $InnerPivotKeys, $OuterPivotKeys, $PathCosts, $Warmup, $Cooldown) {
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
        $trimVal = $val.Trim()
        if ($trimVal -ne "") {
            $newArr += $trimVal
        }
    }
    return $newArr
}