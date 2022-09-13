$NTTTCPPivots = @("sessions", "bufferLen", "bufferCount", "none")
$LATTEPivots = @("protocol", "sendMethod", "none")
$CTSPivots = @("sessions", "none")
$CPSPivots = @("none")
$LagScopePivots = @("none")

$DEFAULT_SAVEPATH = ".\visualization.xlsx"
$DEFAULT_JSONPATH = ".\stats.xlsx"

 

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
    elseif ($FakeBoundParameters.ContainsKey("LagScope")) {
        return $LagScopePivots | where {$_ -like "$WordsToComplete*"}
    } 
    elseif ($FakeBoundParameters.ContainsKey("CPS")) {
        return $CPSPivots | where {$_ -like "$WordsToComplete*"}
    } 
    else {
        return @("")
    }
}

class PivotArgumentCompleter : ArgumentCompleter {
    PivotArgumentCompleter() : base($Global:PACScript) {
    }
}
 
function processing_end {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)] [String] $Output,
        [Parameter()] $Starttime 
    )

    # Collect statistics
    # $timestamp = $start | Get-Date -f yyyy.MM.dd_hh.mm.ss

    # Display version and file save location
    
    $endtime = (Get-Date)
    $delta   = $endtime - $Starttime
    Clear-Host
    Write-Host "$(get-date): xcvEnd"
    Write-Host "---------------`nTime:   $($delta.Minutes) Min $($delta.Seconds) Sec"  
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
        CPS
        Lagscope

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
        LagScope:   protocol, none
        CPS:        none
    
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

        [Parameter(Mandatory=$true, ParameterSetName="CPS")]
        [Switch] $CPS,

        [Parameter(Mandatory=$true, ParameterSetName="LagScope")]
        [Switch] $LagScope,

        [Parameter(Mandatory=$false, ParameterSetName="LagScope")]
        [Parameter(Mandatory=$false, ParameterSetName="LATTE")]
        [Int] $NumHistogramBuckets = 50,

        [Parameter(Mandatory=$false)]
        [String] $BaselineDir = "", 

        [Parameter(Mandatory=$false)]
        [String] $BaselineDatasetName = "",

        [Parameter(Mandatory=$false)]
        [String] $TestDir = "",

        [Parameter(Mandatory=$false)]
        [String] $TestDatasetName = "",

        [Parameter(Mandatory=$false)]
        [String] $JsonInputPath = "",

        [Parameter(Mandatory=$false)]
        [PivotArgumentCompleter()]
        [String] $InnerPivot = "none",

        [Parameter(Mandatory=$false)]
        [PivotArgumentCompleter()]
        [String] $OuterPivot = "none",

        [Parameter(Mandatory=$false)] 
        [ValidatePattern('[.]*\.xl[txms]+')]
        [String] $SavePath=$DEFAULT_SAVEPATH,

        
        [Parameter(Mandatory=$false)] 
        [string] $JsonSavePath = "",

        [Parameter(Mandatory=$false)]
        [Int] $Warmup = 0,

        [Parameter(Mandatory=$false)]
        [Int] $Cooldown = 0,

        [Parameter(Mandatory=$false, ParameterSetName = "LATTE")]
        [Parameter(Mandatory=$false, ParameterSetName = "CPS")]
        [Parameter(Mandatory=$false, ParameterSetName = "LagScope")]
        [Int] $SubsampleRate = -1, 
        
        [Parameter(Mandatory=$false)]
        [Switch] $NoExcel
    )

    $ErrorActionPreference = "Stop"

    Confirm-Parameters -BaselineDir $BaselineDir -TestDir $TestDir -JsonInputPath $JsonInputPath -NoExcel $NoExcel -JsonSavePath $JsonSavePath

    if ($NoExcel -and ($JsonSavePath -eq "")) {
        $JsonSavePath = $DEFAULT_JSONPATH
    }
    if ($JsonInputPath -ne "") {
        $JsonSavePath = ""
    }

    $starttime = (Get-Date)
    Clear-Host 

    # Normalize SavePath
    $file = Split-Path $SavePath -Leaf
    $dir = Convert-Path $(Split-Path $SavePath -Parent)
    $fullPath = Join-Path $dir $file

    # Normalize JsonSavePath
    if ($JsonSavePath) {
        $JsonFile = Split-Path $JsonSavePath -Leaf
        $jsonDir = Convert-Path $(Split-Path $JsonSavePath -Parent)
        $jsonFullPath = Join-Path $jsonDir $JsonFile 
    }

    $tool = $PSCmdlet.ParameterSetName
    Confirm-Pivots -Tool $tool -InnerPivot $InnerPivot -OuterPivot $OuterPivot 

    # Temporarily replace "none" with empty string to ensure compat.
    $InnerPivot = if ($InnerPivot -eq "none") {""} else {$InnerPivot}
    $OuterPivot = if ($OuterPivot -eq "none") {""} else {$OuterPivot}

    if (-not $NoExcel) {
        Add-ExcelTypes 
    }


    if ($JsonInputPath) {
        $jsonInputObj = Get-Content -Path $JsonInputPath | ConvertFrom-Json 
        $processedData =  Parse-RestructedData -DataObj $jsonInputObj
    } else {
        # Parse Data
        $baselineRaw = Get-RawData -Tool $tool -DirName $BaselineDir -InnerPivot $InnerPivot -OuterPivot $OuterPivot -Warmup $Warmup -Cooldown $Cooldown

        $testRaw     = $null
        if ($TestDir) {
            $testRaw = Get-RawData -Tool $tool -DirName $TestDir -Mode "Test" -InnerPivot $InnerPivot -OuterPivot $OuterPivot -Warmup $Warmup -Cooldown $Cooldown
        } 
        $processedData = Process-Data -BaselineRawData $baselineRaw -TestRawData $testRaw -InnerPivot $InnerPivot -OuterPivot $OuterPivot -NumHistogramBuckets $NumHistogramBuckets
        
        $processedData.meta.datasetNames = @{}
        $processedData.meta.datasetNames["baseline"] = "baseline"
        $processedData.meta.datasetNames["test"] = "test"
        $processedData.meta.usedCustomNames = @{
            "baseline"  = $false
            "test"      = $false
        }
        if ($BaselineDatasetName) {
            $processedData.meta.datasetNames["baseline"] = $BaselineDatasetName
            $processedData.meta.usedCustomNames["baseline"] = $true
        }
        if ($TestDatasetName) {
            $processedData.meta.datasetNames["test"] = $TestDatasetName
            $processedData.meta.usedCustomNames["test"] = $true
        }
    }

    if ($JsonSavePath) {
        $output = Restructure-Data -DataObj $processedData
        ($output | ConvertTo-Json -Depth 100) | Out-File -FilePath $jsonFullPath
    }

    if ($NoExcel) {
        $endtime = (Get-Date)
        $delta   = $endtime - $starttime 
        Write-Host "Time: $($delta.Minutes) Min $($delta.Seconds) Sec"   
        Write-Host "Json File: $jsonFullPath `n`n" 
        return
    }

    $tables = @()
    foreach ($oPivotKey in $processedData.data.Keys) { 
        $tables += Format-Stats -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
        if ($tool -in @("NTTTCP", "CTStraffic")) {
            #$tables += Format-RawData      -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool   
            $tables += Format-Quartiles    -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -NoNewWorksheets
            $tables += Format-MinMaxChart  -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -NoNewWorksheets
        } elseif ($tool -in @("LATTE", "LagScope")) { 
            #$tables += Format-Distribution -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -Prop "latency" -SubSampleRate $SubsampleRate
            $tables += Format-Histogram    -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool
            #$tables  += Format-Percentiles -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool  
        } elseif ($tool -in @("CPS")) {
            $tables += Format-Distribution -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -Prop "conn/s" -SubSampleRate $SubsampleRate
            $tables += Format-Distribution -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool -Prop "close/s" -SubSampleRate $SubsampleRate
            $tables += Format-Histogram    -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool  
        }
        $tables  += Format-Percentiles -DataObj $processedData -OPivotKey $oPivotKey -Tool $tool 
    }
    
    
    Create-ExcelFile -Tables $tables -SavePath $fullPath 
    for ($i = 0; $i -lt 4; $i++) 
    {
        if ( (-not $processedData.meta.comparison) -and ($i -eq 1)) {
            continue
        }
        Write-Progress -Activity "Processing" -Id $i -Completed
    } 
    $endtime = (Get-Date)
    $delta   = $endtime - $starttime 

    # Powershell wont print more than two lines here
    Write-Host "Time: $($delta.Minutes) Min $($delta.Seconds) Sec"   
    Write-Host "Report: $fullPath" 
    Write-Host "Json File: $jsonFullPath `n`n"

    #processing_end -OutPut $fullPath -Starttime $starttime     
    if ($JsonSavePath) { 
        ($output | ConvertTo-Json -Depth 100) | Out-File -FilePath $JsonSavePath  
    }
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
        "CPS" {
            $CPSPivots
            break
        }
        "LagScope" {
            $LagScopePivots
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
function Confirm-Parameters($BaselineDir, $TestDir, $JsonInputPath, $NoExcel, $JsonSavePath) {
    if ($TestDir -and ($BaselineDir -eq "")) {
        Write-Error "TestDir should only be used if BaselineDir is non-null"
    }
    if (($BaselineDir -ne "") -and ($JsonInputPath -ne "")) { 
        Write-Error "Cannot injest raw data and pre-processed data at the same time. Use either BaselineDir+TestDir OR JsonInputPath"
    }
    if (($BaselineDir -eq "") -and ($JsonInputPath -eq "")) {
        Write-Error "Must provide a non-null value to either BaselineDir or InputJsonPath parameters"
    } 
    if (($JsonInputPath -ne "") -and ($JsonSavePath -ne "")) {
        Write-Warning "Ignoring JsonSavePath parameter since JsonInputPath was used to provide input (input is already pre-processed json file)"
    }
    if (($JsonInputPath -ne "") -and ($NoExcel)) {
        Write-Error "Cannot run in NoExcel mode when providing input through JsonInputPath"
    }
}
 