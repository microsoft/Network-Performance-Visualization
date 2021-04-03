$NTTTCPPivots = @("sessions", "bufferLen", "bufferCount", "none")
$LATTEPivots = @("protocol", "sendMethod", "none")
$CTSPivots = @("sessions", "none")

$starttime = (Get-Date)

#
# Define atrribute that dynamically tab completes pivot params.
# This is not a form of validation like [ValidateScript()]
#
[ScriptBlock] $Global:PACScript = { 
    param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)

    if ($FakeBoundParameters.ContainsKey("NTTTCP")) {
        return $NTTTCPPivots | where {$_ -like "$WordToComplete*"}
    } elseif ($FakeBoundParameters.ContainsKey("LATTE")) {
        return $LATTEPivots | where {$_ -like "$WordToComplete*"}
    } elseif ($FakeBoundParameters.ContainsKey("CTStraffic")) {
        return $CTSPivots | where {$_ -like "$WordToComplete*"}
    } else {
        return @("")
    }
}

class PivotArgumentCompleter : ArgumentCompleter {
    PivotArgumentCompleter() : base($Global:PACScript) {
    }
}

function processing_start {
    Write-Host "$(get-date): Start"
}

function processing_time {
    Write-Host "$(get-date): Progress"
}

function processing_end {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$true)] [String] $Output
    )

    # Collect statistics
    # $timestamp = $start | Get-Date -f yyyy.MM.dd_hh.mm.ss

    # Display version and file save location

    Write-Host "$(get-date): End"
    Write-Host "---------------"
    $endtime = (Get-Date)
    $delta   = $endtime - $starttime
    Write-Host "Time:   $($delta.Minutes) Min $($delta.Seconds) Sec"
    Write-Host "Report: $OutPut"
    Write-Host " "
}

<#
.SYNOPSIS
    Visualizes network performance data via excel tables and charts

.Description
    This cmdlet parses raw data files produced from various network performance monitoring tools, processes them, and produces 
    visualizations of the data. This tool is capable of visualizing data from the following tools:

        NTTTCP
        LATTE
        CTStraffic

    This tool can aggregate data over several iterations of test runs, and can be used to visualize comparisons
    between a baseline and test set of data.

.PARAMETER NTTTCP
    Flag that sets New-Visualization command to run in NTTTCP mode

.PARAMETER LATTE
    Flag that sets New-Visualization command to run in LATTE mode

.PARAMETER CTStraffic
    Flag that sets New-Visualization command to run in CTStraffic mode

.PARAMETER BaselineDir
    Path to directory containing network performance data files to be consumed as baseline data.

.PARAMETER TestDir
    Path to directory containing network performance data files to be consumed as test data. Providing
    this parameter runs the tool in comparison mode.

.PARAMETER InnerPivot
    Name of the property to use as a pivot variable. Valid values for this parameter are:
        NTTTCP:     sessions, bufferLen, bufferCount, none
        LATTE:      protocol, sendMethod, none
        CTSTraffic: sessions, none
    
    The default for all cases is none.

    The inner pivot varies in value within each worksheet; the different pivot values are used as table column labels.

.PARAMETER OuterPivot
    Name of the property to use as a pivot variable. This parameter has the same valid values as InnerPivot.
    The outer pivot varies in value across different worksheets and remains constant within each worksheet. 

.PARAMETER SavePath
    Full path to excel output file (with extension) where the report should be generated. ex: C:\TempDirName\Filename.xlsx

.PARAMETER SubsampleRate
    Only relevant when run running tool in LATTE context. Sets the subsampling rate used to create the raw
    data chart displaying the latency distribution over time. 

.LINK
    https://github.com/microsoft/Network-Performance-Visualization
#>
function New-NetworkVisualization {
    [CmdletBinding()]
    [Alias("New-NetVis")]
    param (
        [Parameter(Mandatory=$true, ParameterSetName="NTTTCP")]
        [Switch] $NTTTCP,

        [Parameter(Mandatory=$true, ParameterSetName="LATTE")]
        [Switch] $LATTE,

        [Parameter(Mandatory=$true, ParameterSetName="CTStraffic")]
        [Switch] $CTStraffic,

        [Parameter(Mandatory=$true)]
        [String] $BaselineDir, 

        [Parameter(Mandatory=$false)]
        [String] $TestDir = $null,

        [Parameter(Mandatory=$false)]
        [PivotArgumentCompleter()]
        [String] $InnerPivot = "none",

        [Parameter(Mandatory=$false)]
        [PivotArgumentCompleter()]
        [String] $OuterPivot = "none",

        [Parameter(Mandatory=$true)]
        [ValidatePattern('[.]*\.xl[txms]+')]
        [String] $SavePath,

        [Parameter(Mandatory=$false, ParameterSetName = "LATTE")]
        [Int] $SubsampleRate = 50
    )

    processing_start

    $ErrorActionPreference = "Stop"

    $tool = $PSCmdlet.ParameterSetName
    Confirm-Pivots -Tool $tool -InnerPivot $InnerPivot -OuterPivot $OuterPivot
    processing_time

    # Temporarily replace "none" with empty string to ensure compat.
    $InnerPivot = if ($InnerPivot -eq "none") {""} else {$InnerPivot}
    $OuterPivot = if ($OuterPivot -eq "none") {""} else {$OuterPivot}

    Add-ExcelTypes
    processing_time

    # Parse Data
    $baselineRaw = Get-RawData -Tool $tool -DirName $BaselineDir
    processing_time
    $testRaw     = $null
    if ($TestDir) {
        $testRaw = Get-RawData -Tool $tool -DirName $TestDir
    }
    processing_time

    $processedData = Process-Data -BaselineRawData $baselineRaw -TestRawData $testRaw -InnerPivot $InnerPivot -OuterPivot $OuterPivot
    processing_time

    foreach ($oPivotKey in $processedData.data.Keys) {
        if ($tool -in @("NTTTCP", "CTStraffic")) {
            $tables += Format-RawData      -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
            $tables += Format-Stats        -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -Metrics @("min", "mean", "max", "std dev")
            $tables += Format-Quartiles    -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -NoNewWorksheets
            $tables += Format-MinMaxChart  -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -NoNewWorksheets
        } elseif ($tool -in @("LATTE")) {
            $tables += Format-Distribution -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -SubSampleRate $SubsampleRate
            $tables += Format-Stats        -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
            $tables += Format-Histogram    -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
        }
        $tables  += Format-Percentiles -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
        processing_time
    }
    $filePathAndName = Create-ExcelFile -Tables $tables -SavePath $SavePath 
    
    processing_end -OutPut $filePathAndName
}

<#
.SYNOPSIS
    This function loads Microsoft.Office.Interop.Excel.dll
    from the GAC so exported enums can be accessed. 
#>
function Add-ExcelTypes {
    $gac = "$env:WINDIR\assembly\GAC_MSIL"
    $version = (Get-ChildItem "$gac\office" | select -Last 1).Name # e.g. 15.0.0.0__71e9bce111e9429c

    Add-Type -Path "$gac\office\$version\office.dll"
    Add-Type -Path "$gac\Microsoft.Office.Interop.Excel\$version\Microsoft.Office.Interop.Excel.dll"
}

<#
.SYNOPSIS
    Check if user-provided pivots are valid pivots given the
    chosen context, and that the two pivots are not the same. 
.PARAMETER Tool
    Name of the tool that 
.PARAMETER InnerPivot
    Property to be used as the inner pivot variable
.PARAMETER OuterPivot
    Property to be used as the outer pivot variable
#>
function Confirm-Pivots ($Tool, $InnerPivot, $OuterPivot) {
    $validPivots = switch ($Tool) {
        "NTTTCP" {
            $NTTTCPPivots
            break
        }
        "LATTE" {
            $LATTEPivots
            break
        }
        "CTSTraffic" {
            $CTSPivots
            break
        }
    }

    foreach ($curPivot in @($InnerPivot, $OuterPivot)) {
        if ($curPivot -notin $validPivots) {
            Write-Error "Invalid pivot property '$curPivot'. Supported pivots for $Tool are: $($validPivots -join ", ")." -ErrorAction "Stop"
        }
    }

    if (($InnerPivot -eq $OuterPivot) -and ($InnerPivot -ne "none")) {
        Write-Error "Cannot use the same property for both inner and outer pivots." -ErrorAction "Stop"
    }
}