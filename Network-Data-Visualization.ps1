Set-ExecutionPolicy -ExecutionPolicy  ByPass

$MBTOGB = 1 / 1024
$BTOGB = 1 / (1024 * 1024 * 1024) 

$LIGHTGREEN = 10416289
$GREEN = 1268766
$LIGHTRED = 10396159
$RED = 2108032
$BLUES = @(10249511, 14058822, 16758932)
$ORANGES = @(294092, 1681916, 6014716)

$XLENUM = New-Object -TypeName PSObject

$Percentiles = @(0, 1, 5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90, 95, 96, 97, 98,`
                                         99, 99.9, 99.99, 99.999, 99.9999, 99.99999, 100)
# Interface ---------------------------------------------------------------------------------------

Function Network-Data-Visualization 
{
    <#
    .Description
    This cmdlet parses, processes, and visualizes network performance data files produced by one of various 
    possible network performance tools. This tool is capable of visualizing data from the following tools:

        NTTTCP
        LATTE
        CTStraffic

    This tool can aggregate data over several iterations of test runs, and can be used to visualize comparisons
    between a baseline and test set of data.

    .PARAMETER NTTTCP
    Flag that sets NetData-Visualizer to run in NTTTCP mode

    .PARAMETER LATTE
    Flag that sets NetData-Visualizer to run in LATTE mode

    .PARAMETER CTStraffic
    Flag that sets NetData-Visualizer to run in CTStraffic mode

    .PARAMETER BaselineDir
    Path to directory containing network performance data files to be consumed as baseline data.

    .PARAMETER TestDir
    Path to directory containing network performance data files to be consumed as test data. Providing
    this parameter runs the tool in comparison mode.

    .PARAMETER SaveDir
    Path to directory where excel file will be saved with an auto-generated name if no SavePath provided

    .PARAMETER SavePath
    Path to exact file where excel file will be saved. 

    .SYNOPSIS
    Visualizes network performance data via excel tables and charts
    #>  
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ParameterSetName="NTTTCP")]
        [switch]$NTTTCP,

        [Parameter(Mandatory=$true, ParameterSetName="LATTE")]
        [switch]$LATTE,

        [Parameter(Mandatory=$true, ParameterSetName="CTStraffic")]
        [switch]$CTStraffic, 

        [Parameter(Mandatory=$true, ParameterSetName = "NTTTCP")]
        [Parameter(Mandatory=$true, ParameterSetName = "LATTE")]
        [Parameter(Mandatory=$true, ParameterSetName = "CTStraffic")]
        [string]$BaselineDir, 

        [Parameter()]
        [string]$TestDir=$null,

        [Parameter()]
        [string]$SaveDir = "$home\Documents\PSreports",

        [Parameter()]
        [string]$SavePath = $none,

        [Parameter(ParameterSetName = "LATTE")]
        [int]$SubsampleRate = 50
    )
    
    Init-XLENUM

    $ErrorActionPreference = "Stop"

    $tool = ""
    if ($NTTTCP) 
    {
        $tool = "NTTTCP"
    } 
    elseif ($LATTE) 
    {
        $tool = "LATTE"

    } 
    elseif ($CTStraffic) 
    {
        $tool = "CTStraffic"
    }

    $baselineRaw = Parse-Files -Tool $tool -DirName $BaselineDir
    $testRaw = $null
    if ($TestDir) 
    {
        $testRaw = Parse-Files -Tool $tool -DirName $TestDir
    } 

    $processedData = Process-Data -BaselineDataObj $baselineRaw -TestDataObj $testRaw

    [Array] $tables = @() 
    if (@("NTTTCP", "CTStraffic") -contains $tool) 
    {
        $tables += Format-RawData -DataObj $processedData -TableTitle $tool
        $tables += "NEW"
        $tables += Format-Stats -DataObj $processedData -TableTitle $tool -Metrics @("min", "mean", "max", "std dev")
        $tables += Format-Quartiles -DataObj $processedData -TableTitle $tool
        $tables += Format-MinMaxChart -DataObj $processedData -TableTitle $tool
        $tables += "NEW" 
    } 
    elseif (@("LATTE") -contains $tool ) 
    {
        $tables += Format-Distribution -DataObj $processedData -TableTitle $tool -SubSampleRate $SubsampleRate
        $tables += "NEW"
        $tables += Format-Stats -DataObj $processedData -TableTitle $tool
    }

    $tables += Format-Percentiles -DataObj $processedData -TableTitle $tool
    $fileName = Create-ExcelSheet -Tables $tables -SaveDir $SaveDir -Tool $tool -SavePath $SavePath

    Write-Host "Created report at $filename"
}


Function Init-XLENUM {
    $xl = New-Object -ComObject Excel.Application -ErrorAction Stop
    $xl.Quit() | Out-Null

    $xl.GetType().Assembly.GetExportedTypes() | Where-Object {$_.IsEnum} | ForEach-Object {
        $enum = $_
        $enum.GetEnumNames() | ForEach-Object {
            $XLENUM | Add-Member -MemberType NoteProperty -Name $_ -Value $enum::($_) -Force
        }
    }
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    [System.GC]::Collect() | Out-Null
    [System.GC]::WaitForPendingFinalizers() | Out-Null
}

# File Parsing ------------------------------------------------------------------------------------


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
Function Parse-Files 
{
    param (
        [Parameter(Mandatory=$true)] [string]$DirName, 
        [Parameter()] [string] $Tool
    )

    try
    {
        $files = Get-ChildItem $DirName
    } 
    catch
    {
        Write-Warning "Error at Parse-Files: failed to open directory at path: $DirName"
        Write-Error $_.Exception.Message
    }

    if ($Tool -eq "NTTTCP") 
    {
        [Array] $dataEntries = @()
        foreach ($file in $files) 
        {
            $fileName = $file.FullName
            try 
            {
                $dataEntry = Parse-NTTTCP -FileName $fileName
            } 
            catch 
            {
                Write-Warning "Error at Parse-NTTTCP: failed to parse file $fileName"
                Write-Error $_.Exception.Message
            }

            $dataEntries += ,$dataEntry
        }

        $rawData = @{
            "meta" = @{
                "units" = @{
                    "cycles" = "cycles/byte"
                    "throughput" = "Gb/s"
                }
                "goal" = @{
                    "throughput" = "increase"
                    "cycles" = "decrease"
                }
                "format" = @{
                    "throughput" = "0.0"
                    "cycles" = "0.00"
                    "% change" = "+#.0%;-#.0%;0.0%"
                }
                "noTable" = [Array] @("filename")
                "sortProp" = "sessions"
            }
            "data" = $dataEntries
        }

        return $rawData
    } 
    elseif ($Tool -eq "LATTE") 
    {
        [Array] $dataEntries = @() 
        foreach ($file in $files) 
        {
            $fileName = $file.FullName
            try 
            {
                $dataEntry = Parse-LATTE -FileName $fileName
            } 
            catch 
            {
                Write-Warning "Error at Parse-LATTE: failed to parse file $fileName"
                Write-Error $_.Exception.Message
            }

            $dataEntries += ,$dataEntry
        }

        $rawData = @{
            "meta" = @{
                "units" = @{
                    "latency" = "us"
                }
                "goal" = @{
                    "latency" = "decrease"
                }
                "format" = @{
                    "latency" = "#.0"
                    "% change" = "+#.0%;-#.0%;0.0%"
                }
                "noTable" = [Array]@("filename")
            }
            "data" = $dataEntries
        }

        return $rawData
    }
    elseif ($Tool -eq "CTStraffic") 
    {
        [Array] $dataEntries = @() 
        foreach ($file in $files) 
        {
            $fileName = $file.FullName
            try 
            {
                $ErrorActionPreference = "Stop"
                $dataEntry = Parse-CTStraffic -FileName $fileName
            } 
            catch 
            {
                Write-Warning "Error at Parse-CTStraffic: failed to parse file $fileName"
                Write-Error $_.Exception.Message
            }

            $dataEntries += ,$dataEntry
        }

        $rawData = @{
            "meta" = @{
                "units" = @{
                    "throughput" = "Gb"
                }
                "goal" = @{
                    "throughput" = "increase"
                }
                "format" = @{
                    "throughput" = "0.0"
                    "% change" = "+#.0%;-#.0%;0.0%"
                }
                "sortProp" = "sessions"
                "noTable" = [Array]@("filename")
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
Function Parse-NTTTCP ([string] $FileName) 
{
    [XML]$file = Get-Content $FileName

    [decimal] $cycles = $file.ChildNodes.cycles.'#text'
    [decimal] $throughput = $MBTOGB * [decimal]$file.ChildNodes.throughput[0].'#text'
    [int] $sessions = $file.ChildNodes.parameters.max_active_threads

    $dataEntry = @{
        "sessions" = $sessions
        "throughput" = $throughput
        "cycles" = $cycles
        "filename" = $FileName.Split('\')[-1]
    }

    return $dataEntry
}


##
# Parse-CTStraffic
# ----------------
# This function parses a single file containing CTStraffic data in an CSV format. 
# Relevant data is then extracted, packaged into a dataEntry object, and returned.
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
Function Parse-CTStraffic ([string] $Filename) 
{
    $file = Get-Content $Filename

    $firstLine = $true
    $idxs = @{}
    [Array] $throughputs = @()
    [Array] $sessions = @()

    foreach ($line in $file) 
    {
        if ($firstLine) 
        {
            $firstLine = $false
            $splitLine = $line.Split(',')
            $col = 0
            foreach($token in $splitLine) 
            {
                if (@("SendBps", "In-Flight") -contains $token) 
                {
                    $idxs[$token] = $col
                }
                $col++
            }
        } 
        else 
        {
            $splitLine = $line.Split(',')

            $throughputs += ($BTOGB * [decimal]$splitLine[$idxs["SendBps"]])
            $sessions += $splitLine[$idxs["In-Flight"]]
        }
    }

    $dataEntry = @{
        "sessions" = [int]($sessions | Measure -Maximum).Maximum
        "throughput" = [decimal]($throughputs | Measure -Average).Average
        "filename" = $Filename.Split('\')[-1] 
    }

    return $dataEntry
}


##
# Parse-LATTE
# -----------
# This function parses a single file containing LATTE data. Each line containing
# a latency sample is then extracted into an array, packaged into a dataEntry 
# object, and returned.
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
Function Parse-LATTE ([string] $FileName) 
{
    $file = Get-Content $FileName

    [Array] $latency = @()
    foreach ($line in $file) 
    {
        $latency += ,[int]$line
    }

    $dataEntry = @{
        "latency" = $latency
        "filename" = $FileName.Split('\')[-1]
    }

    return $dataEntry
}



# Data Processing --------------------------------------------------------------------

##
# Process-Data
# ------------
# This function organizes raw data by property and sortProp (if applicable), 
# and then calculates statistics and percentiles over each sub-category of data. Processed data, 
# the original raw data, and some meta data are then stored together in an object and returned. 
#
# Parameters
# ----------
# BaselineDataObj (HashTable) - Object containing baseline raw data
# TestDataObj (HashTable) - Object containing test raw data (optional) 
#
# Return 
# ------
# HashTable - Object containing processed data, raw data, and meta data
##
Function Process-Data 
{
    param (
        [Parameter(Mandatory=$true)] [PSobject[]] $BaselineDataObj,
        [Parameter()] [PSobject[]] $TestDataObj
    )
    try 
    {
        $processedDataObj = @{
            "meta" = $BaselineDataObj.meta
            "data" = @{}
            "rawData" = @{
                "baseline" = $BaselineDataObj.data
            }
        }
        if ($TestDataObj) 
        {
            $processedDataObj.meta.comparison = $true
            $processedDataObj.rawData.test = $TestDataObj.data
        }

        $sortProp = $BaselineDataObj.meta.sortProp
        foreach ($prop in ([Array]$BaselineDataObj.data)[0].Keys) 
        {
            if (($prop -eq $BaselineDataObj.meta.sortProp) -or ($BaselineDataObj.meta.noTable -contains $prop)) 
            {
                continue
            }

            # Organize baseline data by sortProp values
            $processedDataObj.data.$prop = @{}
            $modes = @("baseline")
            foreach($item in $BaselineDataObj.data) 
            {
                $sortKey = "allData"
                if ($sortProp) 
                {
                    $sortKey = $item.$sortProp 
                } 
                if (-not ($processedDataObj.data.$prop.Keys -contains $sortKey)) 
                {
                    $processedDataObj.data.$prop.$sortKey = @{
                        "baseline" = @{
                            "orderedData" = @()
                        }
                    }
                }
                $processedDataObj.data.$prop.$sortKey.baseline.orderedData += $item.$prop
            }

            # Organize test data by sortProp values, if test data is provided
            if ($TestDataObj) 
            {
                $modes += "test"
                foreach ($item in $TestDataObj.data) 
                {
                    $sortKey = "allData"
                    if ($sortProp) 
                    {
                        $sortKey = $item.$sortProp 
                    }
                    if (-not ($processedDataObj.data.$prop.$sortKey.Keys -contains "test")) 
                    {
                        $processedDataObj.data.$prop.$sortKey.test = @{
                            "orderedData" = @()
                        }
                        $processedDataObj.data.$prop.$sortKey."% change" = @{}
                    }
                    $processedDataObj.data.$prop.$sortKey.test.orderedData += $item.$prop
                }
            }

            # Calculate stats and percentiles for each sortKey, calculate % change if necessary
            foreach ($sortKey in $processedDataObj.data.$prop.Keys) 
            {
                foreach ($mode in $modes) 
                {
                    $processedDataObj.data.$prop.$sortKey.$mode.orderedData = $processedDataObj.data.$prop.$sortKey.$mode.orderedData | Sort
                    $stats = Calculate-Stats -arr $processedDataObj.data.$prop.$sortKey.$mode.orderedData
                    $processedDataObj.data.$prop.$sortKey.$mode.stats = $stats
                    $processedDataObj.data.$prop.$sortKey.$mode.percentiles = @{}
                    foreach ($percentile in $Percentiles) 
                    {
                        $idx = [int] (($percentile / 100) * ($processedDataObj.data.$prop.$sortKey.$mode.orderedData.Count - 1))
                        $value = $processedDataObj.data.$prop.$sortKey.$mode.orderedData[$idx]
                        $processedDataObj.data.$prop.$sortKey.$mode.percentiles.$percentile = $value
                    }
                } 
                if ($TestDataObj) 
                {
                    $processedDataObj.data.$prop.$sortKey."% change".stats = @{}
                    foreach ($metric in $processedDataObj.data.$prop.$sortKey.$mode.stats.Keys) 
                    {
                        $diff = $processedDataObj.data.$prop.$sortKey."test".stats.$metric - $processedDataObj.data.$prop.$sortKey."baseline".stats.$metric
                        $percentChange = 100 * ($diff / [math]::Abs( $processedDataObj.data.$prop.$sortKey."baseline".stats.$metric))
                        $processedDataObj.data.$prop.$sortKey."% change".stats.$metric = $percentChange
                    }
                    $processedDataObj.data.$prop.$sortKey."% change".percentiles = @{}
                    foreach ($percentile in $Percentiles) 
                    {
                        $percentChange = 100 * (($processedDataObj.data.$prop.$sortKey."test".percentiles.$percentile / $processedDataObj.data.$prop.$sortKey."baseline".percentiles.$percentile) - 1)
                        $processedDataObj.data.$prop.$sortKey."% change".percentiles.$percentile = $percentChange
                    }
                } 
            }
        }
        return $processedDataObj
    } 
    catch 
    {
        Write-Warning "Error in Process-Data"
        Write-Error $_.Exception.Message
    }
}
 
# Table Formatting -------------------------------------------------------------------


##
# Format-RawData
# --------------
# This function formats raw data into tables, one for each dataEntry property. Data samples are
# organized by their sortProp and labeled with the name of the file from which the data sample was extracted.
#
# Parameters
# ----------
# DataObj (HashTable) - Object containing processed data, raw data, and meta data
# TableTitle (String) - Title to be displayed at the top of each table
# 
# Return
# ------
# HashTable[] - Array of HashTable objects which each store a table of formatted raw data
#
##
Function Format-RawData 
{
    param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter()] [String] $TableTitle = ""
    )

    try 
    {
        $tables = @() 
        $meta = $DataObj.meta
        $sortProp = $meta.sortProp 
        $legend = @{
            "meta" = @{
                "colLabelDepth" = 1
                "rowLabelDepth" = 1
                "dataWidth" = 2
                "dataHeight" = 3
            }
            "rows" = @{
                " " = 0
                "  " = 1
                "   " = 2
            }
            "cols" = @{
                "legend" = @{
                    " " = 0
                    "  " = 1
                }
            }
            "data" = @{
                "legend" = @{
                    " " = @{
                        " " = @{
                            "value" = "Test values are compared against the mean basline value."
                        }
                        "  " = @{
                            "value" = "Test values which show improvement are colored green:"
                        }
                        "   " = @{
                            "value" = "Test values which show regression are colored red:"
                        }
                    }
                    "  " = @{
                        "  " = @{
                            "value" = "Improvement"
                            "fontColor" = $GREEN
                            "cellColor" = $LIGHTGREEN
                        }
                        "   " = @{
                            "value" = "Regression"
                            "fontColor" = $RED
                            "cellColor" = $LIGHTRED
                        }
                    } 
                }
            }
        }

        if ($meta.comparison) 
        {
            $tables += $legend
        }

        # Fill single array with all data and sort, label data as baseline/test if necessary
        [Array] $data = @()
        $baseData = $DataObj.rawData.baseline
        foreach ($entry in $baseData) 
        {
            if ($meta.comparison) 
            {
                $entry.baseline = $true
            }
            $data += $entry
        }

        if ($meta.comparison) 
        {
            $testData = $DataObj.rawData.test
            foreach ($entry in $testData) 
            {
                $data += $entry
            }
        }

        $data = Sort-ByProp -Data $data -Prop $sortProp

        foreach ($prop in $dataObj.data.Keys) 
        {
            $table = @{
                "rows" = @{
                    $prop = @{}
                }
                "cols" = @{
                    $TableTitle = @{}
                }
                "meta" = @{
                    "columnFormats" = @()
                }
                "data"  = @{
                    $TableTitle = @{}
                }
            }
            $col = 0
            $row = 0
            foreach ($entry in $data)
            {
                $sortKey = $entry.$sortProp

                # Add column labels to table
                if (-not ($table.cols.$TableTitle.Keys -contains $sortKey)) 
                {
                    if ($meta.comparison) 
                    {
                        $table.cols.$TableTitle.$sortKey = @{
                            "baseline" = $col
                            "test" = $col + 1
                        }
                        $table.meta.columnFormats += $meta.format.$prop
                        $table.meta.columnFormats += $meta.format.$prop
                        $col += 2
                        $table.data.$TableTitle.$sortKey = @{
                            "baseline" = @{
                                $prop = @{}
                            }
                            "test" = @{
                                $prop = @{}
                            }
                        }
                    } 
                    else 
                    {
                        $table.cols.$TableTitle.$sortKey = $col
                        $table.meta.columnFormats += $meta.format.$prop
                        $col += 1
                        $table.data.$TableTitle.$sortKey = @{
                            $prop = @{}
                        }
                    }
                }

                # Add row labels and fill data in table
                $filename = $entry.fileName
                $table.rows.$prop.$filename = $row
                $row += 1
                if ($meta.comparison) 
                {
                    if ($entry.baseline) 
                    {
                        $table.data.$TableTitle.$sortKey.baseline.$prop.$filename = @{
                            "value" = $entry.$prop
                        }
                    } 
                    else 
                    {
                        $table.data.$TableTitle.$sortKey.test.$prop.$filename = @{
                            "value" = $entry.$prop
                        }
                        $params = @{
                            "Cell" = $table.data.$TableTitle.$sortKey.test.$prop.$filename
                            "TestVal" = $entry.$prop
                            "BaseVal" = $DataObj.data.$prop.$sortKey.baseline.stats.mean
                            "Goal" = $meta.goal.$prop
                        }
                        $table.data.$TableTitle.$sortKey.test.$prop.$filename = Select-Color @params
                    }
                } 
                else 
                {
                    $table.data.$TableTitle.$sortKey.$prop.$filename = @{
                        "value" = $entry.$prop
                    }
                }
            }
            $table.meta.dataWidth = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
        }

        foreach ($entry in $data) 
        {
            if ($entry.baseline) 
            {
                $entry.Remove("baseline")
            }
        }
        return $tables
    } 
    catch 
    {
        Write-Warning "Error at Format-RawData"
        Write-Error $_.Exception.Message
    }
}


##
# Format-Stats
# -------------------
# This function formats statistical metrics (min, mean, max, etc) into a table, one per property.
# When run in comparison mode, the table also displays % change and is color-coded to indicate 
# improvement/regression.
#
# Parameters
# ----------
# DataObj (HashTable) - Object containing processed data, raw data, and meta data
# TableTitle (String) - Title to be displayed at the top of each table
# Metrics (String[]) - Array containing statistical metrics that should be displayed on generated 
#                      tables. All metrics are displayed if this parameter is null. 
#
# Return
# ------
# HashTable[] - Array of HashTable objects which each store a table of formatted statistical data
#
##
Function Format-Stats 
{
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter()] [String] $TableTitle = "",

        [Parameter()] [Array] $Metrics=$null
    )
    
    try 
    {
        $tables = @()
        $data = $DataObj.data
        $meta = $DataObj.meta
        foreach ($prop in $data.keys) 
        { 
            $table = @{
                "rows" = @{
                    $prop = @{}
                }
                "cols" = @{
                    $TableTitle = @{}
                }
                "meta" = @{
                    "columnFormats" = @()
                }
                "data" = @{
                    $TableTitle = @{}
                }
            }
            $col = 0
            $row = 0
            foreach ($sortKey in $data.$prop.Keys | Sort) 
            { 

                # Add column labels to table
                if (-not $meta.comparison) 
                {
                    $table.cols.$TableTitle.$sortKey = $col
                    $table.meta.columnFormats += $meta.format.$prop 
                    $table.data.$TableTitle.$sortKey = @{
                        $prop = @{}
                    }
                    $col += 1
                } 
                else 
                {
                    $table.cols.$TableTitle.$sortKey = @{
                        "baseline" = $col
                        "% Change" = $col + 1
                        "test" = $col + 2
                    }
                    $table.meta.columnFormats += $meta.format.$prop
                    $table.meta.columnFormats += $meta.format."% change"
                    $table.meta.columnFormats += $meta.format.$prop
                    $col += 3
                    $table.data.$TableTitle.$sortKey = @{
                        "baseline" = @{
                            $prop = @{}
                        }
                        "% Change" = @{
                            $prop = @{}
                        }
                        "test" = @{
                            $prop = @{}
                        }
                    }
                }

                if (-not $Metrics) 
                {
                    $Metrics = ($data.$prop.$sortKey.baseline.stats.Keys | Sort)
                }

                # Add row labels and fill data in table
                foreach ($metric in $Metrics) 
                {
                    if (-not ($table.rows.$prop.Keys -contains $metric)) 
                    {
                        $table.rows.$prop.$metric = $row
                        $row += 1
                    }
                    if (-not $meta.comparison) 
                    {
                        $table.data.$TableTitle.$sortKey.$prop.$metric = @{"value" = $data.$prop.$sortKey.baseline.stats.$metric}
                    } 
                    else 
                    {
                        $table.data.$TableTitle.$sortKey.baseline.$prop.$metric = @{"value" = $data.$prop.$sortKey.baseline.stats.$metric}
                        $table.data.$TableTitle.$sortKey.test.$prop.$metric = @{"value" = $data.$prop.$sortKey.test.stats.$metric}
                    
                        $percentChange = $data.$prop.$sortKey."% change".stats.$metric
                        $table.data.$TableTitle.$sortKey."% change".$prop.$metric = @{"value" = "$percentChange %"}
                        $params = @{
                            "Cell" = $table.data.$TableTitle.$sortKey."% change".$prop.$metric
                            "TestVal" = $data.$prop.$sortKey.test.stats.$metric
                            "BaseVal" = $data.$prop.$sortKey.baseline.stats.$metric
                            "Goal" = $meta.goal.$prop
                        }
                        if (@("std dev", "variance", "kurtosis", "std err", "range") -contains $metric) 
                        {
                            $params.goal = "decrease"
                            $table.data.$TableTitle.$sortKey."% change".$prop.$metric = Select-Color @params
                        } 
                        elseif ( -not (@("sum", "count") -contains $metric)) 
                        {
                            $table.data.$TableTitle.$sortKey."% change".$prop.$metric = Select-Color @params
                        }
                    }
                }
            }

            $table.meta.dataWidth = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
        }
        return $tables
    } catch {
        Write-Warning "Error at Format-Stats"
        Write-Error $_.Exception.Message
    }
}


##
# Format-Quartiles
# ----------------
# This function formats a table in order to create a chart that displays the quartiles
# of each data subcategory (organized by sortProp), one chart per property.
#
# Parameters
# ----------
# DataObj (HashTable) - Object containing processed data, raw data, and meta data
# TableTitle (String) - Title to be displayed at the top of each table
#
# Return
# ------
# HashTable[] - Array of HashTable objects which each store a table of formatted quartile data
#
##
Function Format-Quartiles 
{
    param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter()] [String] $TableTitle = ""
    )
    try 
    {
        $tables = @()
        $data = $DataObj.data
        $meta = $DataObj.meta
        $sortProp = $meta.sortProp
        foreach ($prop in $data.Keys) 
        { 
            $format = $meta.format.$prop
            $table = @{
                "rows" = @{
                    $prop = @{
                        $sortProp = @{}
                    }
                }
                "cols" = @{
                    $TableTitle = @{
                        "min" = 0
                        "Q1" = 1
                        "Q2" = 2
                        "Q3" = 3
                        "Q4" = 4
                    }
                }
                "meta" = @{
                    "columnFormats" = @($format, $format, $format, $format, $format )
                    "dataWidth" = 5
                }
                "data" = @{
                    $TableTitle = @{
                        "min" = @{
                            $prop = @{
                                $sortProp = @{}
                            }
                        }
                        "Q1" = @{
                            $prop = @{
                                $sortProp = @{}
                            }
                        }
                        "Q2" = @{
                            $prop = @{
                                $sortProp = @{}
                            }
                        }
                        "Q3" = @{
                            $prop = @{
                                $sortProp = @{}
                            }
                        }
                        "Q4" = @{
                            $prop = @{
                                $sortProp = @{}
                            }
                        }
                    }
                }
                "chartSettings" = @{ 
                    "chartType"= $XLENUM.xlColumnStacked
                    "plotBy" = $XLENUM.xlColumns
                    "xOffset" = 1
                    "YOffset" = 1
                    "title"="$prop quartiles"
                    "hideLegend" = $true
                    "dataTable" = $true
                    "seriesSettings"= @{
                        1 = @{
                            "hide"=$true
                        }
                    }
                    "axisSettings" = @{
                        1 = @{
                            "majorGridlines" = $true
                        }
                        2 = @{
                            "minorGridlines" = $true
                            "title" = $meta.units[$prop]
                        }
                    }
                }
            }
        
            $row = 0
            # Add row labels and fill data in table
            foreach ($sortKey in $data.$prop.Keys | Sort) 
            {
                if (-not $meta.comparison)
                {
                    $table.rows.$prop.$sortProp.$sortKey = $row
                    $row += 1
                    $table.data.$TableTitle.min.$prop.$sortProp.$sortKey = @{ "value" = $data.$prop.$sortKey.baseline.stats.min }
                    $table.data.$TableTitle.Q1.$prop.$sortProp.$sortKey = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[25] - $data.$prop.$sortKey.baseline.stats.min }
                    $table.data.$TableTitle.Q2.$prop.$sortProp.$sortKey = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[50] - $data.$prop.$sortKey.baseline.percentiles[25] } 
                    $table.data.$TableTitle.Q3.$prop.$sortProp.$sortKey = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[75] - $data.$prop.$sortKey.baseline.percentiles[50]}
                    $table.data.$TableTitle.Q4.$prop.$sortProp.$sortKey = @{ "value" = $data.$prop.$sortKey.baseline.stats.max - $data.$prop.$sortKey.baseline.percentiles[75] }
                } 
                else 
                {
                    $table.rows.$prop.$sortProp.$sortKey = @{
                        "baseline" = $row
                        "test" = $row + 1
                    }
                    $row += 2
                    $table.data.$TableTitle.min.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.stats.min }
                        "test" = @{ "value" = $data.$prop.$sortKey.test.stats.min}
                    }
                    $table.data.$TableTitle.Q1.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[25] - $data.$prop.$sortKey.baseline.stats.min }
                        "test" = @{ "value" = $data.$prop.$sortKey.test.percentiles[25] - $data.$prop.$sortKey.test.stats.min }
                    }
                    $table.data.$TableTitle.Q2.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[50] - $data.$prop.$sortKey.baseline.percentiles[25] } 
                        "test" = @{ "value" = $data.$prop.$sortKey.test.percentiles[50] - $data.$prop.$sortKey.test.percentiles[25] } 
                    }
                    $table.data.$TableTitle.Q3.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[75] - $data.$prop.$sortKey.baseline.percentiles[50] } 
                        "test" = @{ "value" = $data.$prop.$sortKey.test.percentiles[75] - $data.$prop.$sortKey.test.percentiles[50] }
                    }
                    $table.data.$TableTitle.Q4.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.stats.max - $data.$prop.$sortKey.baseline.percentiles[75] }
                        "test" = @{ "value" = $data.$prop.$sortKey.test.stats.max - $data.$prop.$sortKey.test.percentiles[75] }
                    }
                }

            }
            $table.meta.dataWidth = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
        }
        return $tables
    } 
    catch 
    {
        Write-Warning "Error at Format-Quartiles"
        Write-Error $_.Exception.Message
    }
}


##
# Format-MinMaxChart
# ----------------
# This function formats a table that displays min, mean, and max of each data subcategory, 
# one table per property. This table primarily serves to generate a line chart for the
# visualization of this data.
#
# Parameters
# ----------
# DataObj (HashTable) - Object containing processed data, raw data, and meta data
# TableTitle (String) - Title to be displayed at the top of each table
#
# Return
# ------
# HashTable[] - Array of HashTable objects which each store a table of formatted data
#
##
Function Format-MinMaxChart 
{
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter()] [String] $TableTitle = ""
    )
    
    try 
    {
        $tables = @()
        $data = $DataObj.data
        $meta = $DataObj.meta
        $sortProp = $meta.sortProp

        foreach ($prop in $data.keys) 
        {
            $cappedProp = (Get-Culture).TextInfo.ToTitleCase($prop) 
            $table = @{
                "rows" = @{
                    $prop = @{}
                }
                "cols" = @{
                    $TableTitle = @{
                        $sortProp = @{}
                    }
                }
                "meta" = @{
                    "columnFormats" = @()
                }
                "data" = @{
                    $TableTitle = @{
                        $sortProp = @{}
                    }
                }
                "chartSettings" = @{
                    "chartType" = $XLENUM.xlLineMarkers
                    "plotBy" = $XLENUM.xlRows
                    "title" = $cappedProp
                    "xOffset" = 1
                    "yOffset" = 1
                    "dataTable" = $true
                    "hideLegend" = $true
                    "axisSettings" = @{
                        1 = @{
                            "majorGridlines" = $true
                        }
                        2 = @{
                            "minorGridlines" = $true
                            "title" = $meta.units.$prop
                        }
                    }
                }
            }
            if ($meta.comparison) 
            {
                $table.chartSettings.seriesSettings = @{
                    1 = @{
                        "color" = $BLUES[2]
                        "markerColor" = $BLUES[2]
                    }
                    2 = @{
                        "color" = $ORANGES[2]
                        "markerColor" = $ORANGES[2]
                    }
                    3 = @{
                        "color" = $BLUES[1]
                        "markerColor" = $BLUES[1]
                    }
                    4 = @{
                        "color" = $ORANGES[1]
                        "markerColor" = $ORANGES[1]
                    }
                    5 = @{
                        "color" = $BLUES[0]
                        "markerColor" = $BLUES[0]
                    }
                    6 = @{
                        "color" = $ORANGES[0]
                        "markerColor" = $ORANGES[0]
                    }
                }
            } 
            else 
            {
                $table.chartSettings.seriesSettings = @{
                    1 = @{
                        "color" = $BLUES[2]
                        "markerColor" = $BLUES[2]
                    }
                    2 = @{
                        "color" = $BLUES[1]
                        "markerColor" = $BLUES[1]
                    }
                    3 = @{
                        "color" = $BLUES[0]
                        "markerColor" = $BLUES[0]
                    }
                }
            }
            $col = 0
            $row = 0
            foreach ($sortKey in $data.$prop.Keys | Sort) 
            {
                # Add column labels to table
                $table.cols.$TableTitle.$sortProp.$sortKey = $col
                $table.meta.columnFormats += $meta.format.$prop
                $col += 1
                $table.data.$TableTitle.$sortProp.$sortKey = @{
                    $prop = @{}
                }
            
                # Add row labels and fill data in table
                foreach ($metric in @("min", "mean", "max")) 
                {
                    if (-not ($table.rows.$prop.Keys -contains $metric)) { 
                        if (-not $meta.comparison){
                            $table.rows.$prop.$metric = $row
                            $row += 1
                        } else {
                            $table.rows.$prop.$metric = @{
                                "baseline" = $row
                                "test" = $row + 1
                            } 
                            $row += 2
                        }
                    }
                    if (-not ($table.data.$TableTitle.$sortProp.$sortKey.$prop.Keys -contains $metric)) {
                        $table.data.$TableTitle.$sortProp.$sortKey.$prop.$metric = @{}
                    }

                    if (-not $meta.comparison) {
                        $table.data.$TableTitle.$sortProp.$sortKey.$prop.$metric = @{"value" = $data.$prop.$sortKey.baseline.stats.$metric}
                    } 
                    else 
                    {
                        $table.data.$TableTitle.$sortProp.$sortKey.$prop.$metric.baseline = @{"value" = $data.$prop.$sortKey.baseline.stats.$metric}
                        $table.data.$TableTitle.$sortProp.$sortKey.$prop.$metric.test = @{"value" = $data.$prop.$sortKey.test.stats.$metric}
                    }
                }

            }
            $table.meta.dataWidth = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
        }
        return $tables
    } 
    catch 
    {
        Write-Warning "Error at Format-MinMaxChart"
        Write-Error $_.Exception.Message
    }
}


##
# Format-Percentiles
# ----------------
# This function formats a table displaying percentiles of each data subcategory, one
# table per property + sortProp combo. When in comparison mode, percent change is also
# plotted and is color-coded to indicate improvement/regression. A chart is also formatted
# with each table.  
#
# Parameters
# ----------
# DataObj (HashTable) - Object containing processed data, raw data, and meta data
# TableTitle (String) - Title to be displayed at the top of each table
#
# Return
# ------
# HashTable[] - Array of HashTable objects which each store a table of formatted percentile data
#
##
Function Format-Percentiles 
{
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter()] [String] $TableTitle = ""
    )
    try 
    {
        $tables = @()
        $data = $DataObj.data
        $meta = $DataObj.meta
        $sortProp = $meta.sortProp
        $baseTitle = $TableTitle
        foreach ($prop in $data.Keys) 
        {
            foreach ($sortKey in $data.$prop.Keys | Sort) 
            {
                $note = ""
                if ($sortProp) {
                    $note = " - $sortProp $sortKey"
                    $TableTitle = "$baseTitle$note"
                }
                $table = @{
                    "rows" = @{
                        "percentiles" = @{}
                    }
                    "cols" = @{
                        $TableTitle = @{
                            $prop = 0
                        }
                    }
                    "meta" = @{
                        "columnFormats" = @($meta.format.$prop)
                    }
                    "data" = @{
                        $TableTitle = @{
                            $prop = @{
                                "percentiles" = @{}
                            }
                        }
                    }
                    "chartSettings"= @{
                        "title" = "$prop Percentiles$note"
                        "yOffset" = 1
                        "xOffset" = 1
                        "chartType" = $XLENUM.xlXYScatterLinesNoMarkers
                        "seriesSettings" = @{
                            1 = @{
                                "color" = $BLUES[1]
                            }
                        }
                        "axisSettings" = @{
                            1 = @{
                                "max" = 100
                                "title" = "Percentiles"
                                "minorGridlines" = $true
                            }
                            2 = @{
                                "title" = $meta.units[$prop]
                                "logarithmic" = $true
                                #"min" = 10
                            }
                        }
                    }
                }
                if ($meta.comparison) 
                {
                    $table.cols.$TableTitle.$prop = @{
                        "baseline" = 0
                        "% change" = 1
                        "test" = 2
                    }
                    $table.data.$TableTitle.$prop = @{
                        "baseline" = @{
                            "percentiles" = @{}
                        }
                        "% change" = @{
                            "percentiles" = @{}
                        }
                        "test" = @{
                            "percentiles" = @{}
                        }
                    }
                    $table.chartSettings.seriesSettings[2] = @{
                        "delete" = $true
                    }
                    $table.chartSettings.seriesSettings[3] = @{
                        "color" = $ORANGES[1]
                    }
                    $table.meta.columnFormats = @($meta.format.$prop, $meta.format."% change", $meta.format.$prop)
                }
                $row = 0
                # Add row labels and fill data in table
                foreach ($percentile in $data.$prop.$sortKey.baseline.percentiles.Keys | Sort) 
                {
                    $table.rows.percentiles.$percentile = $row
                    if ($meta.comparison) 
                    {
                        $percentage = $data.$prop.$sortKey."% change".percentiles.$percentile
                        $percentage = "$percentage %"
                        $table.data.$TableTitle.$prop.baseline.percentiles.$percentile = @{"value" = $data.$prop.$sortKey.baseline.percentiles.$percentile}
                        $table.data.$TableTitle.$prop."% change".percentiles.$percentile = @{"value" = $percentage}
                        $table.data.$TableTitle.$prop.test.percentiles.$percentile = @{"value" = $data.$prop.$sortKey.test.percentiles.$percentile}
                        $params = @{
                            "Cell" = $table.data.$TableTitle.$prop."% change".percentiles.$percentile
                            "TestVal" = $data.$prop.$sortKey.test.percentiles.$percentile
                            "BaseVal" = $data.$prop.$sortKey.baseline.percentiles.$percentile
                            "Goal" = $meta.goal.$prop
                        }
                        $table.data.$TableTitle.$prop."% change".percentiles.$percentile = Select-Color @params
                    } 
                    else 
                    {
                        $table.data.$TableTitle.$prop.percentiles.$percentile = @{"value" = $data.$prop.$sortKey.baseline.percentiles.$percentile}
                    }
                    $row += 1
                
                }
                $table.meta.dataWidth = Get-TreeWidth $table.cols
                $table.meta.colLabelDepth = Get-TreeDepth $table.cols
                $table.meta.dataHeight = Get-TreeWidth $table.rows
                $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
                $tables = $tables + $table
            }
        }
        return $tables  
    } 
    catch 
    {
        Write-Warning "Error at Format-Percentiles"
        Write-Error $_.Exception.Message
    }
}


##
# Format-Distribution
# -------------------
# This function formats a table in order to create a chart that displays the the
# distribution of data over time.
#
# Parameters
# ----------
# DataObj (HashTable) - Object containing processed data, raw data, and meta data
# TableTitle (String) - Title to be displayed at the top of each table
# Prop (String) - The name of the property for which a table should be created (raw data must be in array form)
# SubSampleRate (int) - How many time samples should be grouped together for a single data point on the chart
#
# Return
# ------
# HashTable[] - Array of HashTable objects which each store a table of formatted distribution data
#
##
Function Format-Distribution 
{
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter()] [String] $TableTitle = "",

        [Parameter()] [String] $Prop = "latency",
    
        [Parameter()] [Int] $SubSampleRate = 50
    )
    try 
    {
        $meta = $DataObj.meta
        $modes = @("baseline")
        if ($meta.comparison) 
        {
            $modes += "test"
        }
        $tables = @()
        $originalTitle = $TableTitle
        foreach ($mode in $modes) 
        {
            if ($tables.Count -gt 0) 
            {
                $tables += "NEW"
            }
            $label = ""
            if ($modes.Count -gt 1) 
            {
                $label = "$mode -"
                $TableTitle = "$label$originalTitle"
            }
            $data = $dataObj.rawData.$mode
            $timeSamples = $data[0][$Prop].Count
            $table = @{
                "meta" = @{}
                "rows" = @{
                    "Data Point" = @{}
                }
                "cols" = @{
                    $TableTitle = @{
                        "Time Segment" = 0
                        $Prop = 1
                    }
                }
                "data" = @{
                    $TableTitle = @{
                        "Time Segment" = @{
                            "Data Point" = @{}
                        }
                        "latency" = @{
                            "Data Point" = @{}
                        }
                    }
                }
                "chartSettings" = @{
                    "chartType" = $XLENUM.xlXYScatter
                    "yOffset" = 2
                    "xOffset" = 2
                    "title" = "$label Temporal Latency Distribution"
                    "axisSettings" = @{
                        1 = @{
                            "title" = "Time Series"
                            "max" = $timeSamples
                            "minorGridlines" = $true
                            "majorGridlines" = $true
                        }
                        2 = @{
                            "title" = "us"
                            "logarithmic" = $true
                            "min" = 10
                        }
                    }
                }
            }

            # Add row labels and fill data in table
            $i = 0
            $row = 0
            $NumSegments = $data[0].$Prop.Count / $SubSampleRate
            while ($i -lt $NumSegments) 
            {
                [Array]$segmentData = @()
                foreach ($entry in $data) 
                {
                    $segmentData += $entry[$Prop][($i * $SubSampleRate) .. ((($i + 1) * $SubSampleRate) - 1)]
                }
                $segmentData = $segmentData | Sort
                $time = $i * $subSampleRate
                if ($segmentData.Count -ge 10) 
                {
                    $table.rows."Data Point".$row = $row
                    $table.rows."Data Point".($row + 1) = $row + 1
                    $table.rows."Data Point".($row + 2) = $row + 2
                    $table.rows."Data Point".($row + 3) = $row + 3
                    $table.rows."Data Point".($row + 4) = $row + 4
                    $table.data.$TableTitle."Time Segment"."Data Point".$row = @{"value" = $time}
                    $table.data.$TableTitle."Time Segment"."Data Point".($row + 1) = @{"value" = $time}
                    $table.data.$TableTitle."Time Segment"."Data Point".($row + 2) = @{"value" = $time}
                    $table.data.$TableTitle."Time Segment"."Data Point".($row + 3) = @{"value" = $time}
                    $table.data.$TableTitle."Time Segment"."Data Point".($row + 4) = @{"value" = $time}
                    $table.data.$TableTitle.$Prop."Data Point".$row = @{"value" = $segmentData[0]}
                    $table.data.$TableTitle.$Prop."Data Point".($row + 1) = @{"value" = $segmentData[[int]($segmentData.Count / 4)]}
                    $table.data.$TableTitle.$Prop."Data Point".($row + 2) = @{"value" = $segmentData[[int]($segmentData.Count / 2)]}
                    $table.data.$TableTitle.$Prop."Data Point".($row + 3) = @{"value" = $segmentData[[int]((3 * $segmentData.Count) / 4)]}
                    $table.data.$TableTitle.$Prop."Data Point".($row + 4) = @{"value" = $segmentData[-1]}
                    $row += 5
                } 
                elseif ($segmentData.Count -ge 2) 
                {
                    $table.rows."Data Point".$row = $row
                    $table.rows."Data Point".($row + 1) = $row + 1
                    $table.data.$TableTitle."Time Segment"."Data Point".$row = @{"value" = $time}
                    $table.data.$TableTitle."Time Segment"."Data Point".($row + 1) = @{"value" = $time}
                    $table.data.$TableTitle.$Prop."Data Point".$row = @{"value" = $segmentData[0]}
                    $table.data.$TableTitle.$Prop."Data Point".($row + 1) = @{"value" = $segmentData[-1]}
                    $row += 2
                } 
                else 
                {
                    $table.rows."Data Point".$row = $row
                    $table.data.$TableTitle."Time Segment"."Data Point".$row = @{"value" = $time}
                    $table.data.$TableTitle.$Prop."Data Point".$row = @{"value" = $segmentData[0]}
                    $row += 1
                }
                $i += 1
            }
            $table.meta.dataWidth = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows

            $tables += $table
        }
        return $tables
    } 
    catch 
    {
        Write-Warning "Error at Format-Distribution"
        Write-Error $_.Exception.Message
    }
}


##
# Select-Color
# ------------
# This function selects the color of a cell, indicating whether a test value
# shows an improvement when compared to a baseline value. Improvement is defined
# by the goal (increase/decrease) for the given value.
# 
# Parameters
# ----------
# Cell (HashTable) - Object containg a cell's value and other settings
# TestVal (decimal) - Test value
# BaseVal (decimal) - Baseline value
# Goal (String) - Defines improvement ("increase" or "decrease")
#
# Return
# ------
# HashTable - Object containing a cell's value and other settings
#
##  
Function Select-Color ($Cell, $TestVal, $BaseVal, $Goal) 
{
    if ( $Goal -eq "increase") 
    {
        if ($TestVal -ge $BaseVal) 
        {
            $Cell["fontColor"] = $GREEN
            $Cell["cellColor"] = $LIGHTGREEN
        } 
        else 
        {
            $Cell["fontColor"] = $RED
            $Cell["cellColor"] = $LIGHTRED
        }
    } 
    else 
    {
        if ($TestVal -le $BaseVal) 
        {
            $Cell["fontColor"] = $GREEN
            $Cell["cellColor"] = $LIGHTGREEN
        } 
        else 
        {
            $Cell["fontColor"] = $RED
            $Cell["cellColor"] = $LIGHTRED
        }
    }
    return $cell
}


##
# Calculate-Stats
# ---------------
# Calculates and returns statistical metrics calculated over an array of values
#
# Parameters
# ----------
# Arr (decimal[]) - Array of values to calculate statistics over
#
# Return
# ------
# HashTable - Object containing statistical metric calculated over Arr
#
## 
Function Calculate-Stats ($Arr) 
{
    $measures = ($Arr | Measure -Average -Maximum -Minimum -Sum)
    $stats = @{
        "count" = $measures.Count
        "sum" = $measures.Sum
        "min" = $measures.Minimum
        "mean" = $measures.Average
        "max" = $measures.Maximum
    }
    $N = $measures.Count
    $squareDiffSum = 0
    $cubeDiffSum = 0
    $quadDiffSum = 0
    $curCount = 0
    $curVal = $null
    $mode = $null
    $modeCount = 0
    $Arr = $Arr | Sort
    foreach ($val in $Arr) 
    {
        if ($val -ne $curVal) 
        {
            $curVal = $val
            $curCount = 1
        } 
        else 
        {
            $curCount++ 
        }

        if ($curCount -gt $modeCount) 
        {
            $mode = $val
            $modeCount = $curCount
        }

        $squareDiffSum += [Math]::Pow(($val - $measures.Average), 2)
        $quadDiffSum += [Math]::Pow(($val - $measures.Average), 4)
    }
    $stats["median"] = $Arr[[int]($N / 2)]
    $stats["mode"] = $mode
    $stats["range"] = $stats["max"] - $stats["min"]
    $stats["std dev"] = [Math]::Sqrt(($squareDiffSum / ($N - 1)))
    $stats["variance"] = $squareDiffSum / ($N - 1)
    $stats["std err"] = $stats["std dev"] / [math]::Sqrt($N)

    if ($N -gt 3) 
    {
        $stats["kurtosis"] = (($N * ($N + 1))/( ($N - 1) * ($N - 2) * ($N - 3))) * ($quadDiffSum / [Math]::Pow($stats["variance"], 2)) - 3 * ([Math]::Pow($N - 1, 2) / (($N - 2) * ($N - 3)) )
        foreach ($val in $Arr | Sort) 
        { 
            $cubeDiffSum += [Math]::Pow(($val - $measures.Average) / $stats["std dev"], 3) 
        }
        $stats["skewness"] = ($N / (($N - 1) * ($N - 2))) * $cubeDiffSum
    }
    return $stats
}


##
# Get-TreeWidth
# -------------
# Calculates the width of a tree structure
#
# Parameters 
# ----------
# Tree (HashTable) - Object with a heirarchical tree structure
#
# Return
# ------
# int - Width of Tree
#
##
Function Get-TreeWidth ($Tree) 
{
    if ($Tree.GetType().Name -eq "Int32") 
    {
        return 1
    }
    $width = 0
    foreach ($key in $Tree.Keys) 
    {
        $width += [int](Get-TreeWidth -Tree $Tree[$key])
    }
    return $width
}

##
# Get-TreeWidth
# -------------
# Calculates the depth of a tree structure
#
# Parameters 
# ----------
# Tree (HashTable) - Object with a heirarchical tree structure
#
# Return
# ------
# int - Depth of Tree
#
##
Function Get-TreeDepth ($Tree)
{
    if ($Tree.GetType().Name -eq "Int32") 
    {
        return 0
    }
    $depths = @()
    foreach ($key in $Tree.Keys) 
    {
        $depths = $depths + [int](Get-TreeDepth -Tree $Tree[$key])
    }
    return ($depths | Measure -Maximum).Maximum + 1
}

##
# Sort-ByProp
# -------------
# Sorts an array of objects by the value of a specified property in each object
#
# Parameters 
# ----------
# Data (HashTable[]) - Array of objects
# Prop (String) - Name of property to sort by
#
# Return
# ------
# HashTable[] - Array of objects, sorted by property value
#
##
Function Sort-ByProp 
{
    param(
        [Parameter()]
        [PSObject] $Data,

        [Parameter()]
        [string] $Prop
    )

    if ($Data.length -eq 1) 
    {
        $sorted = @()
        $sorted = $sorted + $Data
        return $sorted
    }
    $arr1 = $Data[0 .. ([int]($Data.length / 2) - 1)]
    $arr2 = $Data[[int]($Data.length / 2) .. ($Data.length - 1)]

    [array] $arr1 = Sort-ByProp -Data $arr1 -Prop $prop
    [array] $arr2 = Sort-ByProp -Data $arr2 -Prop $prop
    $sorted = @()
    $idx1 = 0
    $idx2 = 0
    
    while ($idx1 -lt $arr1.length -and $idx2 -lt $arr2.length) 
    {
        if ($arr1[$idx1].$prop -le $arr2[$idx2].$prop) 
        {
            $sorted = $sorted + $arr1[$idx1]
            $idx1 += 1
        } 
        else 
        {
            $sorted = $sorted + $arr2[$idx2]
            $idx2 += 1
        }
    }

    while ($idx1 -lt $arr1.length) 
    {
        $sorted = $sorted + $arr1[$idx1]
        $idx1 += 1
    }

    while ($idx2 -lt $arr2.length) 
    {
        $sorted = $sorted + $arr2[$idx2]
        $idx2 += 1
    }
    return $sorted
}

# Excel Population --------------------------------------------------------------------------------

##
# Create-ExcelSheet
# -----------------
# This function plots tables in excel and can use those tables to generate charts. The function
# receives data in a custom format which it uses to create tables. 
# 
# Parameters
# ----------
# Tables (HashTable[]) - Array of table objects
# SaveDir (String) - Directory to save excel workbook w/ auto-generated name at if no savePath is provided
# Tool (String) - Name of tool for which a report is being created
# SavePath (String) - Path and filename for where to save excel workbook, if none is supplied a file name is 
#                     auto-generated
#
# Returns
# -------
# (String) - Name of saved file
#
##
Function Create-ExcelSheet 
{
    param (
        [Parameter(Mandatory=$true)] 
        [PSObject[]]$Tables,

        [Parameter()]
        [String] $SaveDir="$home\Documents",

        [Parameter(Mandatory=$true)]
        [string]$Tool,

        [Parameter()]
        [string]$SavePath = $null
    )

    if  ( (-not $SavePath) -and !( Test-Path -Path $SaveDir -PathType "Container" ) ) 
    { 
        New-Item -Path $SaveDir -ItemType "Container" -ErrorAction Stop | Out-Null
    }

    $date = Get-Date -UFormat "%Y-%m-%d_%H-%M-%S"
    if ($SavePath) 
    {
        $excelFile = $SavePath
    } 
    else 
    {
        $excelFile = "$SaveDir\$Tool-Report-$($date).xlsx"
    }
    $excelFile = $excelFile.Replace(" ", "_")

    try 
    {
        $excelObject = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excelObject.Visible = $true
        $workbookObject = $excelObject.Workbooks.Add()
        $worksheetObject = $workbookObject.Worksheets.Item(1)
            
        [int]$rowOffset = 1
        [int] $chartNum = 1
        foreach ($table in $Tables) 
        {
            if ($table -eq "NEW") 
            {
                $chartNum = 1
                $worksheetObject.UsedRange.Columns.Autofit() | Out-Null
                $worksheetObject = $workbookObject.worksheets.Add()
                [int]$rowOffset = 1
                continue
            }

            Fill-ColLabels -Worksheet $worksheetObject -cols $table.cols -startCol ($table.meta.rowLabelDepth + 1) -row $rowOffset | Out-Null
            Fill-RowLabels -Worksheet $worksheetObject -rows $table.rows -startRow ($table.meta.colLabelDepth + $rowOffset) -col 1 | Out-Null
            Fill-Data -Worksheet $worksheetObject -Data $table.data -Cols $table.cols -Rows $table.rows -StartCol ($table.meta.rowLabelDepth + 1) -StartRow ($table.meta.colLabelDepth + $rowOffset) | Out-Null
            if ($table.chartSettings) 
            {
                Create-Chart -Worksheet $worksheetObject -Table $table -StartCol 1 -StartRow $rowOffset -chartNum $chartNum | Out-Null
                $chartNum += 1
            }
            if ($table.meta.columnFormats)
            {
                for ($i = 0; $i -lt $table.meta.columnFormats.Count; $i++) 
                {
                    if ($table.meta.columnFormats[$i]) 
                    {
                        $column = $worksheetObject.Range($worksheetObject.Cells($rowOffset + $table.meta.colLabelDepth, 1 + $table.meta.rowLabelDepth + $i), $worksheetObject.Cells($rowOffset + $table.meta.colLabelDepth + $table.meta.dataHeight - 1, 1 + $table.meta.rowLabelDepth + $i))
                        $column.select() | Out-Null
                        $column.NumberFormat = $table.meta.columnFormats[$i]
                    }
                }
            }
            $selection = $worksheetObject.Range($worksheetObject.Cells($rowOffset, 1), $worksheetObject.Cells($rowOffset + $table.meta.colLabelDepth + $table.meta.dataHeight - 1, $table.meta.rowLabelDepth + $table.meta.dataWidth))
            $selection.select() | Out-Null
            $selection.BorderAround(1, 4) | Out-Null

            $rowOffset += $table.meta.colLabelDepth + $table.meta.dataHeight + 1
        }
        
        $worksheetObject.UsedRange.Columns.Autofit() | Out-Null

        $workbookObject.SaveAs($excelFile,51) | Out-Null # http://msdn.microsoft.com/en-us/library/bb241279.aspx 
        $workbookObject.Saved = $true 
        $workbookObject.Close() | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) | Out-Null  

        $excelObject.Quit() | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) | Out-Null
        [System.GC]::Collect() | Out-Null
        [System.GC]::WaitForPendingFinalizers() | Out-Null

        return [string]$excelFile
    } 
    catch 
    {
        Write-Warning "Error at Create-ExcelSheet"
        Write-Error $_.Exception.Message
    } 
}

##
# Create-Chart
# ------------
# This function uses a table's chartSettings to create and customize a chart
# that visualizes the table's data. 
#
# Parameters
# ----------
# WorkSheet (ComObject) - Object containing the current worksheet's internal state
# Table (HashTable) - Object containing formatted data and chart settings
# StartRow (int) - The row number on which the top of the already-plotted table begins
# StartCol (int) - The column number on which the left side of the already-plotted table begins
# ChartNum (int) - The index this chart will occupy in the worksheet's internally-stored lisrt of charts
#
# Return
# ------
# None
#
##
Function Create-Chart ($Worksheet, $Table, $StartRow, $StartCol, $chartNum) 
{
    $chart = $Worksheet.Shapes.AddChart().Chart 

    $width = $Table.meta.dataWidth + $Table.meta.rowLabelDepth
    $height = $Table.meta.dataHeight + $Table.meta.colLabelDepth
    if ($Table.chartSettings.yOffset) 
    {
        $height -= $Table.chartSettings.yOffset
        $StartRow += $Table.chartSettings.yOffset
    }
    if ($Table.chartSettings.xOffset) 
    {
        $width -= $Table.chartSettings.xOffset
        $StartCol += $Table.chartSettings.xOffset
    }
    if ($Table.chartSettings.chartType) 
    {
        $chart.ChartType = $Table.chartSettings.chartType
    }
    $chart.SetSourceData($Worksheet.Range($Worksheet.Cells($StartRow, $StartCol), $Worksheet.Cells($StartRow + $height - 1, $StartCol + $width - 1)))
    
    if ($Table.chartSettings.plotBy) 
    {
        $chart.PlotBy = $Table.chartSettings.plotBy
    }
     
    if ($Table.chartSettings.seriesSettings) {
        foreach($seriesNum in $Table.chartSettings.seriesSettings.Keys) 
        {
            if ($Table.chartSettings.seriesSettings.$seriesNum.hide) 
            {
                $chart.SeriesCollection($seriesNum).format.fill.ForeColor.TintAndShade = 1
                $chart.SeriesCollection($seriesNum).format.fill.Transparency = 1
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.color) 
            {
                $chart.SeriesCollection($seriesNum).Border.Color = $Table.chartSettings.seriesSettings.$seriesNum.color
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.delete) 
            {
                $chart.SeriesCollection($seriesNum).Delete()
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.markerColor) 
            {
                $chart.SeriesCollection($seriesNum).MarkerBackgroundColor = $Table.chartSettings.seriesSettings.$seriesNum.markerColor
                $chart.SeriesCollection($seriesNum).MarkerForegroundColor = $Table.chartSettings.seriesSettings.$seriesNum.markerColor
                $chart.SeriesCollection($seriesNum).MarkerStyle = $XLENUM.xlMarkerStyleCircle
            }
        }
    }

    if ($Table.chartSettings.axisSettings) 
    {
        foreach($axisNum in $Table.chartSettings.axisSettings.Keys) 
        {
            if ($Table.chartSettings.axisSettings.$axisNum.min) 
            { 
                $Worksheet.chartobjects($chartNum).chart.Axes($axisNum).MinimumScale = $Table.chartSettings.axisSettings.$axisNum.min
            }
            if ($Table.chartSettings.axisSettings.$axisNum.max) 
            { 
                $Worksheet.chartobjects($chartNum).chart.Axes($axisNum).MaximumScale = $Table.chartSettings.axisSettings.$axisNum.max
            }
            if ($Table.chartSettings.axisSettings.$axisNum.logarithmic) 
            {
                $Worksheet.chartobjects($chartNum).chart.Axes($axisNum).scaleType = $XLENUM.xlScaleLogarithmic
            }
            if ($Table.chartSettings.axisSettings.$axisNum.title) 
            {
                $Worksheet.chartobjects($chartNum).chart.Axes($axisNum).HasTitle = $true
                $Worksheet.chartobjects($chartNum).chart.Axes($axisNum).AxisTitle.Caption = $Table.chartSettings.axisSettings.$axisNum.title
            }
            if ($Table.chartSettings.axisSettings.$axisNum.minorGridlines) 
            {
                $Worksheet.chartobjects($chartNum).chart.Axes($axisNum).HasMinorGridlines = $true
            }
            if ($Table.chartSettings.axisSettings.$axisNum.majorGridlines) 
            {
                $Worksheet.chartobjects($chartNum).chart.Axes($axisNum).HasMajorGridlines = $true
            }
        }
    }

    if ($Table.chartSettings.title) 
    {
        $chart.HasTitle = $true
        $chart.ChartTitle.Caption = [string]$Table.chartSettings.title
    }
    if ($Table.chartSettings.hideLegend) 
    {
        $chart.HasLegend = $false
    }
    if ($Table.chartSettings.dataTable) 
    {
        $chart.HasDataTable = $true
    }

    $Worksheet.Shapes.Item("Chart " + $chartNum ).top = $Worksheet.Cells($StartRow, $StartCol + $width + 1).top
    $Worksheet.Shapes.Item("Chart " + $chartNum ).left = $Worksheet.Cells($StartRow, $StartCol + $width + 1).left
}


##
# Fill-Cell
# ---------
# This function fills an excel cell with a value, and optionally also customizes the cell style.
# 
# Parameters
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Row (int) - Row index of the cell to fill
# Col (int) - Column index of the cell to fill
# CellSettings (HashTable) - Object containing the value and style settings for the cell
#
# Return
# ------
# None
#
##
Function Fill-Cell ($Worksheet, $Row, $Col, $CellSettings) 
{
    $Worksheet.Cells($Row, $Col).Borders.LineStyle = $XLENUM.xlContinuous
    if ($CellSettings.fontColor) 
    {
        $Worksheet.Cells($Row, $Col).Font.Color = $CellSettings.fontColor
    }
    if ($CellSettings.cellColor) 
    {
        $Worksheet.Cells($Row, $Col).Interior.Color = $CellSettings.cellColor
    }
    if ($CellSettings.bold) 
    {
        $Worksheet.Cells($Row, $Col).Font.Bold = $true
    }
    if ($CellSettings.center) 
    {
        $Worksheet.Cells($Row, $Col).HorizontalAlignment = $XLENUM.xlHAlignCenter
        $Worksheet.Cells($Row, $Col).VerticalAlignment = $XLENUM.xlHAlignCenter
    }
    if ($CellSettings.value -ne $null) 
    {
        $Worksheet.Cells($Row, $Col) = $CellSettings.value
    }
}

##
# Merge-Cells
# -----------
# This function merges a range of cells into a single cell and adds a border
#
# Parameters
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Row1 (int) - Row index of top left cell of range to merge
# Col1 (int) - Column index of top left cell of range to merge
# Row2 (int) - Row index of bottom right cell of range to merge
# Col2 (int) - Column index of bottom right cell of range to merge
#
# Return 
# ------
# None
#
##
Function Merge-Cells ($Worksheet, $Row1, $Col1, $Row2, $Col2) 
{
    $cells = $Worksheet.Range($Worksheet.Cells($Row1, $Col1), $Worksheet.Cells($Row2, $Col2))
    $cells.Select()
    $cells.MergeCells = $true
    $cells.Borders.LineStyle = $XLENUM.xlContinuous
}


##
# Fill-ColLabels
# --------------
# This function consumes the cols field of a table object, and plots the column labels by recursing 
# through the object. 
#
# Parameters 
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Cols (HashTable) - Object storing column label structure and column indices of labels. 
# StartCol (int) - The column index on which the labels should start being drawn (left edge)
# Row (int) - The row at which the current level of labels should be drawn
#
# Return 
# ------
# (int[]) - Tuple of integers capturing the column index range across which the just-drawn label spans
#
##
Function Fill-ColLabels ($Worksheet, $Cols, $StartCol, $Row) 
{
    $range = @(-1, -1)
    foreach ($label in $Cols.Keys) 
    {
        if ($Cols.$label.GetType().Name -ne "Int32") 
        {
            $subRange = Fill-ColLabels -Worksheet $Worksheet -Cols $Cols.$label -StartCol $StartCol -Row ($Row + 1)
            Merge-Cells -Worksheet $Worksheet -Row1 $Row -Col1 $subRange[0] -Row2 $Row -Col2 $subRange[1] | Out-Null
            $cellSettings = @{
                "value" = $label
                "bold" = $true
                "center" = $true
            }
            Fill-Cell -Worksheet $Worksheet -Row $Row -Col $subRange[0] -CellSettings $cellSettings | Out-Null
            if (($subRange[0] -lt $range[0]) -or ($range[0] -eq -1))
            {
                $range[0] = $subRange[0]
            } 
            if (($subRange[1] -gt $range[1]) -or ($range[0] -eq -1)) 
            {
                $range[1] = $subRange[1]
            }
        } 
        else 
        {
            $cellSettings = @{
                "value" = $label
                "bold" = $true
                "center" = $true
            }
            Fill-Cell $Worksheet -Row $Row -Col ($StartCol + $Cols.$label) -CellSettings $cellSettings | Out-Null
            if (($StartCol + $Cols.$label -lt $range[0]) -or ($range[0] -eq -1)) 
            {
                $range[0] = $StartCol + $Cols.$label
            }
            if (($StartCol + $Cols.$label -gt $range[1]) -or ($range[1] -eq -1)) 
            {
                $range[1] = $StartCol + $Cols.$label
            }
        }    
    }
    return $range
}


##
# Fill-RowLabels
# --------------
# This function consumes the rows field of a table object, and plots the row labels by recursing 
# through the object. 
#
# Parameters 
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Rows (HashTable) - Object storing row label structure and row indices of labels. 
# StartRow (int) - The row index on which the labels should start being drawn (top edge)
# Col (int) - The column at which the current level of labels should be drawn
#
# Return 
# ------
# (int[]) - Tuple of integers capturing the row index range across which the just-drawn label spans
#
##
Function Fill-RowLabels ($Worksheet, $Rows, $StartRow, $Col) 
{
    $range = @(-1, -1)
    foreach ($label in $rows.Keys) 
    {
        if ($Rows.$label.GetType().Name -ne "Int32") 
        {
            $subRange = Fill-RowLabels -Worksheet $Worksheet -Rows $Rows.$label -StartRow $StartRow -Col ($Col + 1)
            Merge-Cells -Worksheet $Worksheet -Row1 $subRange[0] -Col1 $Col -Row2 $subRange[1] -Col2 $Col | Out-Null
            $cellSettings = @{
                "value" = $label
                "bold" = $true
                "center" = $true
            }
            Fill-Cell -Worksheet $Worksheet -Row $subRange[0] -Col $Col -CellSettings $cellSettings | Out-Null
            if (($subRange[0] -lt $range[0]) -or ($range[0] -eq -1))
            {
                $range[0] = $subRange[0]
            } 
            if (($subRange[1] -gt $range[1]) -or ($range[0] -eq -1)) 
            {
                $range[1] = $subRange[1]
            }
        } 
        else 
        {
            $cellSettings = @{
                "value" = $label
                "bold" = $true
                "center" = $true
            }
            Fill-Cell $Worksheet -Row ($StartRow + $Rows.$label) -Col $Col -CellSettings $cellSettings | Out-Null
            if (($StartRow + $Rows.$label -lt $range[0]) -or ($range[0] -eq -1)) 
            {
                $range[0] = $StartRow + $Rows.$label
            }
            if (($StartRow + $Rows.$label -gt $range[1]) -or ($range[1] -eq -1)) 
            {
                $range[1] = $StartRow + $Rows.$label
            }
        }    
    }
    return $range
}


##
# Fill-Data
# ---------
# This function uses the data, rows, and cols fields of a table object to fill in the data
# values of a table. The objects are recursed through depth-first, and the path followed is used to retreive 
# row and column indices from the Rows and Cols objects while data values and cell formatting are retreived from the
# data object. 
#
# Parameters
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Data (HashTable) - Object containing data values and cell formatting
# Cols (HashTable) - Object storing column label structure and column indices of the labels. 
# Rows (HashTable) - Object storing row label structure and row indices of the labels.
# StartCol (int) - Column index where the data range of the table begins (left edge) 
# StartRow (int) - Row index where the data range of the table begins (top edge)
#
# Return
# ------
# None
#
## 
Function Fill-Data ($Worksheet, $Data, $Cols, $Rows, $StartCol, $StartRow) 
{
    if($Cols.GetType().Name -eq "Int32" -and $Rows.GetType().Name -eq "Int32") 
    {
        Fill-Cell -Worksheet $Worksheet -Row ($StartRow + $Rows) -Col ($StartCol + $Cols) -CellSettings $Data
        return
    }  
    foreach ($label in $Data.Keys) 
    {
        if ($Cols.getType().Name -ne "Int32") 
        {
            Fill-Data -Worksheet $Worksheet -Data $Data.$label -Cols $Cols.$label -Rows $Rows -StartCol $StartCol -StartRow $StartRow
        } 
        else 
        {
            Fill-Data -Worksheet $Worksheet -Data $Data.$label -Cols $Cols -Rows $Rows.$label -StartCol $StartCol -StartRow $StartRow
        }
    }
}