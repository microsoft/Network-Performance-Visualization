. "$PSScriptRoot.\Data-Parser.ps1"
. "$PSScriptRoot.\Data-Processor.ps1"
. "$PSScriptRoot.\Data-Formatters.ps1"
. "$PSScriptRoot.\Excel-Plotter.ps1"

$XLENUM = New-Object -TypeName PSObject

$NTTTCPPivots = @("", "sessions", "bufferLen", "bufferCount")
$LATTEPivots = @("", "protocol", "sendMethod")
$CTSPivots = @("", "sessions")


function New-NetworkVisualization {
    <#
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
        NTTTCP:     sessions, bufferCount, bufferLen
        LATTE:      protocol
        CTSTraffic: sessions

    The inner pivot varies in value within each worksheet; the different pivot values are used as table column labels.


    .PARAMETER OuterPivot
    Name of the property to use as a pivot variable. Valid values for this parameter are:
        NTTTCP:     sessions, bufferCount, bufferLen
        LATTE:      protocol
        CTSTraffic: sessions

    The outer pivot varies in value across different worksheets and remains constant within each worksheet. 

    .PARAMETER SavePath
    Full path to excel output file (with extension) where the report should be generated. ex: C:\TempDirName\Filename.xlsx

    .PARAMETER SubsampleRate
    Only relevant when run running tool in LATTE context. Sets the subsampling rate used to create the raw
    data chart displaying the latency distribution over time. 

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
        [string]$InnerPivot="",

        [Parameter()]
        [string]$OuterPivot="",

        [Parameter(Mandatory=$true)]
        [ValidatePattern('[.]*\.xl[txms]+')]
        [string]$SavePath,

        [Parameter(ParameterSetName = "LATTE")]
        [int]$SubsampleRate = 50
    )
    
    Initialize-XLENUM

    $ErrorActionPreference = "Stop"

    # Save tool name
    $tool = Get-ToolName -NTTTCP $NTTTCP -LATTE $LATTE -CTStraffic $CTStraffic
    if (-Not (Validate-Pivots -tool $tool -InnerPivot $InnerPivot -OuterPivot $OuterPivot)) {
        return
    }
    
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


##
# Initialize-XLENUM
# -----------------
# This function fills the content of the global object XLENUM with every enum value defined 
# by the Excel application. 
#
# Parameters
# ----------
# None
# 
# Return
# ------
# None
#
##
function Initialize-XLENUM {
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
    if (($InnerPivot -eq $OuterPivot) -and ($InnerPivot -ne "")) {
        Write-Warning "Cannot use the same property for both inner and outer pivots."
        return $false
    }
    return $true
}