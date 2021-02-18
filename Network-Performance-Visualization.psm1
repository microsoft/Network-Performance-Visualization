. "$PSScriptRoot.\Data-Parser.ps1"
. "$PSScriptRoot.\Data-Processor.ps1"
. "$PSScriptRoot.\Data-Formatters.ps1"
. "$PSScriptRoot.\Excel-Plotter.ps1"

$NTTTCPPivots = @("sessions", "bufferLen", "bufferCount", "none")
$LATTEPivots = @("protocol", "sendMethod", "none")
$CTSPivots = @("sessions", "none")

#
# Define atrribute that dynamically tab completes pivot params.
# This is not a form of validation like [ValidateScript()]
#
[ScriptBlock] $Global:PACScript = { 
    param($CommandName, $ParameterName, $WordToComplete, $CommandAst, $FakeBoundParameters)

    if ($FakeBoundParameters.ContainsKey("NTTTCP")) {
        return $NTTTCPPivots | where {$_ -like "$WordToComplete*"}
    }
    elseif ($FakeBoundParameters.ContainsKey("LATTE")) {
        return $LATTEPivots | where {$_ -like "$WordToComplete*"}
    }
    elseif ($FakeBoundParameters.ContainsKey("CTStraffic")) {
        return $CTSPivots | where {$_ -like "$WordToComplete*"}
    }
    else {
        return @("")
    }
}

class PivotArgumentCompleter : ArgumentCompleter {
    PivotArgumentCompleter() : base($Global:PACScript) {
    }
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

    $ErrorActionPreference = "Stop"
    
    # Save tool name
    $tool = Get-ToolName -NTTTCP $NTTTCP -LATTE $LATTE -CTStraffic $CTStraffic
    if (-Not (Validate-Pivots -tool $tool -InnerPivot $InnerPivot -OuterPivot $OuterPivot)) {
        return
    }

    # Temporarily replace "none" with empty string to ensure compat.
    $InnerPivot = if ($InnerPivot -eq "none") {""} else {$InnerPivot}
    $OuterPivot = if ($OuterPivot -eq "none") {""} else {$OuterPivot}

    Load-ExcelDll

    # Parse Data 
    $baselineRaw = Parse-Files -Tool $tool -DirName $BaselineDir
    $testRaw     = $null
    if ($TestDir) {
        $testRaw = Parse-Files -Tool $tool -DirName $TestDir
    } 

    $processedData = Process-Data -BaselineRawData $baselineRaw -TestRawData $testRaw -InnerPivot $InnerPivot -OuterPivot $OuterPivot

    foreach ($oPivotKey in $processedData.data.Keys) {
        if (@("NTTTCP", "CTStraffic") -contains $tool) {
            $tables += Format-RawData -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
            $tables += Format-Stats -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -Metrics @("min", "mean", "max", "std dev")
            $tables += Format-Quartiles -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -NoNewWorksheets
            $tables += Format-MinMaxChart -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -NoNewWorksheets
        }
        elseif (@("LATTE") -contains $tool ) {
            $tables += Format-Distribution -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -SubSampleRate $SubsampleRate
            $tables += Format-Stats -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
            $tables += Format-Histogram -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
        } 
        $tables  += Format-Percentiles -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
    }
    $fileName = Create-ExcelFile -Tables $tables -SavePath $SavePath 
    Write-Host "Created report at $filename"
}

<#
.SYNOPSIS
    This function loads Microsoft.Office.Interop.Excel.dll
    so exported enums can be accessed. 
#>
function Load-ExcelDll {
    $gac = "$env:WINDIR\assembly\GAC_MSIL"
    $version = (Get-ChildItem "$gac\office" | select -Last 1).Name # e.g. 15.0.0.0__71e9bce111e9429c

    Add-Type -Path "$gac\office\$version\office.dll"
    Add-Type -Path "$gac\Microsoft.Office.Interop.Excel\$version\Microsoft.Office.Interop.Excel.dll"
}

##
# Get-ToolName
# ------------
# Given three tool flags, this function returns a string containing the name of the tool whose data is being visualized
# 
# Parameters
# ----------
# NTTTCP (bool) - Whether tool is being run in NTTTCP context
# LATTE (bool) - Whether tool is being run in LATTE context
# CTStraffic (bool) - Whether tool is being run in CTStraffic context
#
# Return
# ------
# Name of tool whose data is being visualized
#
##
function Get-ToolName ($NTTTCP, $LATTE, $CTStraffic) {
    if ($NTTTCP) {
        return "NTTTCP"
    }
    if ($LATTE) {
        return "LATTE"
    }
    if ($CTStraffic) {
        return "CTStraffic"
    }
}


##
# Validate-Pivots
# ---------------
# This function verifies that the user-provided pivots are valid pivots given the chosen context, and that the 
# two pivots are not the same. 
#
# Parameters
# ----------
# Tool (string) - Name of the tool whose data is being visualized
# InnerPivot (string) - Name of property to be used as the inner pivot variable
# OuterPivot (string) - Name of property to be used as the outer pivot variable  
#
# Return
# ------
# Whether pivots are valid
#
##
function Validate-Pivots ($Tool, $InnerPivot, $OuterPivot) {
    if ($Tool -eq "NTTTCP") {
        foreach ($curPivot in @($InnerPivot, $OuterPivot)) {
            if (-Not ($NTTTCPPivots -contains $curPivot)) {
                $msg = "This tool does not support using '$curPivot' as a pivot for NTTTCP data.`n"
                $msg += "Supported pivots are:`n"

                foreach ($pivot in $NTTTCPPivots) {
                    if ($null -eq $pivot) {
                        continue
                    }
                    $msg += "$pivot`n"
                }
                Write-Warning $msg
                return $false
            }    
        }
    }

    if ($Tool -eq "LATTE") {
        foreach ($curPivot in @($InnerPivot, $OuterPivot)) {
            if (-Not ($LATTEPivots -contains $curPivot)) {
                $msg = "This tool does not support using '$curPivot' as a pivot for LATTE data.`n"
                $msg += "Supported pivots are:`n"

                foreach ($pivot in $LATTEPivots) {
                    if ($null -eq $pivot) {
                        continue
                    }
                    $msg += "$pivot`n"
                }
                Write-Warning $msg
                return $false
            }    
        } 
    }

    if ($Tool -eq "CTStraffic") {
        foreach ($curPivot in @($InnerPivot, $OuterPivot)) {
            if (-Not ($CTSPivots -contains $curPivot)) {
                $msg = "This tool does not support using '$curPivot' as a pivot for CTStraffic data.`n"
                $msg += "Supported pivots are:`n"

                foreach ($pivot in $CTSPivots) {
                    if ($null -eq $pivot) {
                        continue
                    }
                    $msg += "$pivot`n"
                }
                Write-Warning $msg
                return $false
            }    
        }
    }
    if (($InnerPivot -eq $OuterPivot) -and ($InnerPivot -ne "none")) {
        Write-Warning "Cannot use the same property for both inner and outer pivots."
        return $false
    }
    return $true
}