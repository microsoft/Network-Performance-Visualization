using namespace Microsoft.Office.Interop

$TextInfo = (Get-Culture).TextInfo
$WorksheetMaxLen = 31
$HeaderRows = 4
$EPS = 0.0001

# Excel uses BGR color values
$ColorPalette = @{
    "LightGreen" = 0x9EF0A1
    "Green"      = 0x135C1E
    "LightRed"   = 0x9EA1FF
    "Red"        = 0x202A80
    "Blue"      = @(0x633f16, 0x9C6527, 0xD68546, 0xFFB894) # Dark -> Light
    "Orange"    = @(0x005b97, 0x047CCC, 0x19A9FC, 0x5BC6FC)
    "LightGray" = @(0xf5f5f5, 0xd9d9d9)
}

$ABBREVIATIONS = @{
    "sessions" = "sess."
    "bufferLen" = "bufLen."
    "bufferCount" = "bufCt."
    "protocol" = ""
    "sendMethod" = "sndMthd" 
}

<#
.SYNOPSIS
    Converts an index to an Excel column name.
.NOTES
    Valid for A to ZZ
#>
function Get-ColName($n) {
    if ($n -ge 26) {
        $a = [Int][Math]::floor($n / 26)
        $c1 = [Char]($a + 64)
    }

    $c2 = [Char](($n % 26) + 65)

    return "$c1$c2"
}


function Get-WorksheetTitle ($BaseName, $OuterPivot, $OPivotKey, $InnerPivot, $IPivotKey, $Prop="") {
    if ($OuterPivot -and $InnerPivot) {
        $OAbv = $ABBREVIATIONS[$OuterPivot]
        $IAbv = $ABBREVIATIONS[$InnerPivot]

        $name = "$BaseName - $OPivotKey $OAbv - $IPivotKey $IAbv"

        if ($name.Length -gt $WorksheetMaxLen) {
            $name = "$BaseName - $OPivotKey - $IPivotKey"
        }

        return $name
    } 
    elseif ($OuterPivot) {
        $OAbv = $ABBREVIATIONS[$OuterPivot]
        $name = "$BaseName - $OPivotKey $OAbv"

        if ($name.Length -gt $WorksheetMaxLen) {
            $name = "$BaseName - $OPivotKey"
        }

        return $name
    } 
    elseif ($InnerPivot) {
        $IAbv = $ABBREVIATIONS[$InnerPivot]
        $name = "$BaseName - $IPivotKey $IAbv"

        if ($name.Length -gt $WorksheetMaxLen) {
            $name = "$BaseName - $IPivotKey"
        }

        return $name 
    }
    else {
        if ($Prop) { 
            return "$BaseName $($Prop.Replace("/", " per "))"
        }
        return "$BaseName"
    }
}

function Get-TableTitle ($Tool, $OuterPivot, $OPivotKey, $InnerPivot, $IPivotKey, $Comparison, $DatasetName, $UsedCustomName) { 
    $title = $Tool
      
    if ($OuterPivot) {
        $OAbv = $ABBREVIATIONS[$OuterPivot]
        $title = "$title - $OPivotKey $OAbv"
    } 

    if ($InnerPivot) {
        $IAbv = $ABBREVIATIONS[$InnerPivot]
        $title = "$title - $IPivotKey $IAbv"
    }

    if ((-not $Comparison) -and $DatasetName -and $UsedCustomName) {
        $title = "$DatasetName - $title"
    }

    return $title
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
function Format-Stats {
    Param (
        [Parameter(Mandatory=$true)]
        [PSObject[]] $DataObj,

        [Parameter(Mandatory=$true)]
        $OPivotKey,

        [Parameter()]
        [String] $Tool = "",

        [Parameter()]
        [Switch] $NoNewWorksheets
    )
    
    $tables = @()
    $data = $DataObj.data
    $meta = $DataObj.meta
    $innerPivot = $meta.InnerPivot
    $outerPivot = $meta.OuterPivot
    $baselineName = $meta.datasetNames["baseline"]
    $testName = $meta.datasetNames["test"]
   
    $numIters =  $meta.props.Count * $meta.innerPivotKeys.Count
    $completeIters = 0

    foreach ($prop in $meta.props) {
        $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -DatasetName $baselineName -Comparison $meta.comparison -UsedCustomName $meta.usedCustomNames.baseline
        $table = @{
            "rows" = @{
                $prop = @{}
            }
            "cols" = @{
                $tableTitle = @{
                    $innerPivot = @{}
                }
            }
            "meta" = @{
                "columnFormats" = @()
                "name"          = "Stats"  
                "numWrites"     = 1 + 2
            }
            "data" = @{
                $tableTitle = @{
                    $innerPivot = @{}
                }
            }
        }

        $col = 0
        $row = 0
        foreach ($IPivotKey in $meta.innerPivotKeys | Sort) { 

            # Add column labels to table
            if (-not $meta.comparison) {
                $table.cols.$tableTitle.$innerPivot.$IPivotKey  = $col 
                $table.data.$tableTitle.$innerPivot.$IPivotKey  = @{
                    $prop = @{}
                }
                $col += 1
                $table.meta.columnFormats += $meta.format.$prop
                $table.meta.numWrites += 1
            } 
            else {
                $table.cols.$tableTitle.$innerPivot.$IPivotKey = @{
                    $baselineName = $col
                    "% Change" = $col + 1
                    $testName     = $col + 2
                }
                $table.meta.numWrites += 4
                $table.meta.columnFormats += $meta.format.$prop
                $table.meta.columnFormats += "0.0%"
                $table.meta.columnFormats += $meta.format.$prop
                $col += 3
                $table.data.$tableTitle.$innerPivot.$IPivotKey = @{
                    $baselineName = @{
                        $prop = @{}
                    }
                    "% Change" = @{
                        $prop = @{}
                    }
                    $testName = @{
                        $prop = @{}
                    }
                }
            }

            # Add row labels and fill data in table
            
            if ($data.$OPivotKey.$prop.$IPivotKey.baseline.stats) {
                $metrics = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.Keys
            }
            else {
                return
            }

            foreach ($metric in $Metrics) {
                if ($table.rows.$prop.Keys -notcontains $metric) {
                    $table.rows.$prop.$metric = $row
                    $row += 1
                    $table.meta.numWrites += 1
                }

                if (-not $meta.comparison) {
                    $table.data.$tableTitle.$innerPivot.$IPivotKey.$prop.$metric = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.$metric}
                } else {
                    if ($data.$OPivotKey.$prop.$IPivotKey.baseline.stats) {
                        $table.data.$tableTitle.$innerPivot.$IPivotKey.$baselineName.$prop.$metric = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.$metric}
                    }
                    if ($data.$OPivotKey.$prop.$IPivotKey.test.stats) {
                        $table.data.$tableTitle.$innerPivot.$IPivotKey.$testName.$prop.$metric     = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.$metric}
                    }

                    if ($data.$OPivotKey.$prop.$IPivotKey.baseline.stats -and $data.$OPivotKey.$prop.$IPivotKey.test.stats) {
                        $table.data.$tableTitle.$innerPivot.$IPivotKey."% change".$prop.$metric = @{"value" = "=IF([col-1][row]=0, `"--`", ([col+1][row]-[col-1][row])/ABS([col-1][row]))"}
                        
                        $params = @{
                            "Cell"    = $table.data.$tableTitle.$innerPivot.$IPivotKey."% change".$prop.$metric
                            "TestVal" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.$metric
                            "BaseVal" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.$metric
                            "Goal"    = $meta.goal.$prop
                        }

                        # Certain statistics always have the same goal.
                        if ($metric -eq "n") {
                            $params.goal = "increase"
                        } elseif ($metric -in @("range", "variance", "std dev", "std err")) {
                            $params.goal = "decrease"
                        } elseif ($metric -in @("skewness", "kurtosis")) {
                            $params.goal = "none"
                        }

                        $table.data.$tableTitle.$innerPivot.$IPivotKey."% change".$prop.$metric = Set-CellColor @params
                    }
                    
                   

                }
                #$cellRow += 1
            } # foreach $metric
            Write-Progress -Activity "Formatting Tables" -Status "Stats Table" -Id 3 -PercentComplete (100 * (($completeIters++) / $numIters))
        }

        $table.meta.dataWidth     = Get-TreeWidth $table.cols
        $table.meta.colLabelDepth = Get-TreeDepth $table.cols
        $table.meta.dataHeight    = Get-TreeWidth $table.rows
        $table.meta.rowLabelDepth = Get-TreeDepth $table.rows 
        $table.meta.numWrites    += $table.meta.dataHeight * $table.meta.dataWidth 
        $tables += $table
    }

    if (($tables.Count -gt 0) -and (-not $NoNewWorksheets)) {
        $sheetTitle = Get-WorksheetTitle -BaseName "Stats" -OuterPivot $outerPivot -OPivotKey $OPivotKey 
        $tables     = [Array]@($sheetTitle) + $tables 
    }
    Write-Progress -Activity "Formatting Tables" -Status "Stats Table" -Id 3 -PercentComplete 100 
    return $tables
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
function Format-Quartiles {
    param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter(Mandatory=$true)] $OPivotKey, 

        [Parameter()] [String] $Tool = "",

        [Parameter()] [switch] $NoNewWorksheets
    )
    $tables = @()
    $data = $DataObj.data
    $meta = $DataObj.meta
    $innerPivot = $meta.InnerPivot
    $outerPivot = $meta.OuterPivot
    $baselineName = $meta.datasetNames["baseline"]
    $testName = $meta.datasetNames["test"]
 
    $numIters =  $meta.props.Count * $meta.innerPivotKeys.Count
    $completeIters = 0

    foreach ($prop in $meta.props) { 
        $format = $meta.format.$prop
        $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -DatasetName $baselineName -Comparison $meta.comparison -UsedCustomName $meta.usedCustomNames.baseline
        $cappedProp = (Get-Culture).TextInfo.ToTitleCase($prop)
        $table = @{
            "rows" = @{
                $prop = @{
                    $innerPivot = @{}
                }
            }
            "cols" = @{
                $tableTitle = @{
                    "min" = 0
                    "Q1"  = 1
                    "Q2"  = 2
                    "Q3"  = 3
                    "Q4"  = 4
                }
            }
            "meta" = @{ 
                "dataWidth" = 5
                "name" = "Quartiles"
                "numWrites" = 2 + 6
            }
            "data" = @{
                $tableTitle = @{
                    "min" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "Q1" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "Q2" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "Q3" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "Q4" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                }
            }
            "chartSettings" = @{ 
                "chartType"= [Excel.XlChartType]::xlColumnStacked
                "plotBy"   = [Excel.XlRowCol]::xlColumns
                "xOffset"  = 1
                "YOffset"  = 1
                "title"    = "$cappedProp Quartiles"
                "seriesSettings"= @{
                    1 = @{
                        "hide" = $true
                        "name" = " "
                    }
                    2 = @{ 
                        "color" = $ColorPalette.blue[1]
                    }
                    3 = @{ 
                        "color" = $ColorPalette.blue[3]
                    }
                    4 = @{ 
                        "color" = $ColorPalette.blue[0]
                    }
                    5 = @{ 
                        "color" = $ColorPalette.blue[2]
                    }
                }
                "axisSettings" = @{
                    1 = @{
                        "majorGridlines" = $true
                    }
                    2 = @{
                        "minorGridlines" = $true
                        "minorGridlinesColor" = $ColorPalette.LightGray[0]
                        "majorGridlinesColor" = $ColorPalette.LightGray[1]
                        "title" = $meta.units[$prop]
                    }
                }
            }
        }

        if ((-not $meta.comparison) -and $meta.usedCustomNames["baseline"]) {
            $table.chartSettings.title = "$baselineName - $cappedProp Quartiles"
        }

        if ($meta.comparison) {
            $table.cols = @{
                $tableTitle = @{
                    "<baseline>min" = 0
                    "<baseline>Q1"  = 1
                    "<baseline>Q2"  = 2
                    "<baseline>Q3"  = 3
                    "<baseline>Q4"  = 4
                    "<test>min" = 5
                    "<test>Q1"  = 6
                    "<test>Q2"  = 7
                    "<test>Q3"  = 8
                    "<test>Q4"  = 9
                }
            }
            $table.data = @{
                $tableTitle = @{
                    "<baseline>min" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "<baseline>Q1" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "<baseline>Q2" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "<baseline>Q3" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "<baseline>Q4" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "<test>min" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "<test>Q1" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "<test>Q2" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "<test>Q3" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                    "<test>Q4" = @{
                        $prop = @{
                            $innerPivot = @{}
                        }
                    }
                }
            }
            $table.chartSettings.seriesSettings[6] = @{
                "hide" = $true
                "name" = " "
            }
            $table.chartSettings.seriesSettings[7] = @{
                "color" = $ColorPalette.orange[1]
            }
            $table.chartSettings.seriesSettings[8] = @{
                "color" = $ColorPalette.orange[3]
            }
            $table.chartSettings.seriesSettings[9] = @{
                "color" = $ColorPalette.orange[0]
            }
            $table.chartSettings.seriesSettings[10] = @{
                "color" = $ColorPalette.orange[2]
            }
            $table.meta.columnFormats = @($format) * $table.cols.$tableTitle.Count; 
        }
    
        
        # Add row labels and fill data in table
        $row = 0 
        foreach ($IPivotKey in $meta.InnerPivotKeys | Sort) {
            if (-not $meta.comparison) {
                $table.meta.numWrites += 1
                $table.rows.$prop.$innerPivot.$IPivotKey = $row
                $row += 1
                $table.data.$TableTitle.min.$prop.$innerPivot.$IPivotKey = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.min }
                $table.data.$TableTitle.Q1.$prop.$innerPivot.$IPivotKey  = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["25"] - $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.min }
                $table.data.$TableTitle.Q2.$prop.$innerPivot.$IPivotKey  = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["50"] - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["25"] } 
                $table.data.$TableTitle.Q3.$prop.$innerPivot.$IPivotKey  = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["75"] - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["50"]}
                $table.data.$TableTitle.Q4.$prop.$innerPivot.$IPivotKey  = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.max - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["75"] }
            } 
            else {
                $table.meta.numWrites += 3
                $table.rows.$prop.$innerPivot.$IPivotKey = @{
                    $baselineName = $row
                    $testName     = $row + 1
                }
                $row += 2

                $table.data.$TableTitle."<baseline>min".$prop.$innerPivot.$IPivotKey = @{
                    $baselineName = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.min } 
                }
                $table.data.$TableTitle."<test>min".$prop.$innerPivot.$IPivotKey = @{ 
                    $testName     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.min}
                }
                $table.data.$TableTitle."<baseline>Q1".$prop.$innerPivot.$IPivotKey = @{
                    $baselineName = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["25"] - $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.min }
                }
                $table.data.$TableTitle."<test>Q1".$prop.$innerPivot.$IPivotKey = @{
                    $testName     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.percentiles["25"] - $data.$OPivotKey.$prop.$IPivotKey.test.stats.min }
                }
                $table.data.$TableTitle."<baseline>Q2".$prop.$innerPivot.$IPivotKey = @{
                    $baselineName = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["50"] - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["25"] } 
                }
                $table.data.$TableTitle."<test>Q2".$prop.$innerPivot.$IPivotKey = @{
                    $testName     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.percentiles["50"] - $data.$OPivotKey.$prop.$IPivotKey.test.percentiles["25"] } 
                }
                $table.data.$TableTitle."<baseline>Q3".$prop.$innerPivot.$IPivotKey = @{
                    $baselineName = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["75"] - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["50"] } 
                }
                $table.data.$TableTitle."<test>Q3".$prop.$innerPivot.$IPivotKey = @{  
                    $testName     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.percentiles["75"] - $data.$OPivotKey.$prop.$IPivotKey.test.percentiles["50"] }
                }
                $table.data.$TableTitle."<baseline>Q4".$prop.$innerPivot.$IPivotKey = @{
                    $baselineName = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.max - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles["75"] }
                }
                $table.data.$TableTitle."<test>Q4".$prop.$innerPivot.$IPivotKey = @{
                    $testName     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.max - $data.$OPivotKey.$prop.$IPivotKey.test.percentiles["75"] }
                }
            }
            Write-Progress -Activity "Formatting Tables" -Status "Quartiles Table" -Id 3 -PercentComplete (100 * (($completeIters++) / $numIters))

        }

        $table.meta.dataWidth     = Get-TreeWidth $table.cols
        $table.meta.colLabelDepth = Get-TreeDepth $table.cols
        $table.meta.dataHeight    = Get-TreeWidth $table.rows
        $table.meta.rowLabelDepth = Get-TreeDepth $table.rows  
        $table.meta.numWrites    += $table.meta.dataHeight * $table.meta.dataWidth
        
         
        $tables = $tables + $table
    }

    if (($tables.Count -gt 0) -and (-not $NoNewWorksheets)) {
        $sheetTitle = Get-WorksheetTitle -BaseName "Quartiles" -OuterPivot $outerPivot -OPivotKey $OPivotKey
        $tables = @($sheetTitle) + $tables
    }
    Write-Progress -Activity "Formatting Tables" -Status "Quartiles Table" -Id 3 -PercentComplete 100


    return $tables
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
function Format-MinMaxChart {
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter(Mandatory=$true)] $OPivotKey, 

        [Parameter()] [String] $Tool = "",

        [Parameter()] [switch] $NoNewWorksheets
    )
    
    $tables     = @()
    $data       = $DataObj.data
    $meta       = $DataObj.meta
    $innerPivot = $meta.InnerPivot
    $outerPivot = $meta.OuterPivot
    $metrics = @("min", "mean", "max")
    $baselineName = $meta.datasetNames["baseline"]
    $testName = $meta.datasetNames["test"]


    $numIters =  $meta.props.Count* $meta.innerPivotKeys.Count
    $completeIters = 0
    foreach ($prop in $meta.props) {
        $cappedProp = (Get-Culture).TextInfo.ToTitleCase($prop) 
        $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -DatasetName $baselineName -Comparison $meta.comparison -UsedCustomName $meta.usedCustomNames.baseline
        $table = @{
            "rows" = @{
                $prop = @{}
            }
            "cols" = @{
                $tableTitle = @{
                    $innerPivot = @{}
                }
            }
            "meta" = @{
                "columnFormats" = @()
                "name"          = "MinMaxCharts"
                "numWrites"     = 1 + 2
            }
            "data" = @{
                $tableTitle = @{
                    $innerPivot = @{}
                }
            }
            "chartSettings" = @{
                "chartType"    = [Excel.XlChartType]::xlLineMarkers
                "plotBy"       = [Excel.XlRowCol]::xlRows
                "title"        = $cappedProp
                "xOffset"      = 1
                "yOffset"      = 2
                "dataTable"    = $true
                "hideLegend"   = $true
                "axisSettings" = @{
                    1 = @{
                        "majorGridlines" = $true
                    }
                    2 = @{
                        "minorGridlines" = $true
                        "minorGridlinesColor" = $ColorPalette.LightGray[0]
                        "majorGridlinesColor" = $ColorPalette.LightGray[1]
                        "title" = $meta.units.$prop
                    }
                }
            }
        }
        if ($meta.comparison) {
            $table.chartSettings.seriesSettings = @{
                1 = @{
                    "color"       = $ColorPalette.blue[3]
                    "markerColor" = $ColorPalette.blue[3]
                    "markerStyle" = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                    "lineWeight"  = 3
                    "markerSize"  = 5
                }
                2 = @{
                    "color"       = $ColorPalette.orange[3]
                    "markerColor" = $ColorPalette.orange[3]
                    "markerStyle" = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                    "lineWeight"  = 3
                    "markerSize"  = 5
                }
                3 = @{
                    "color"       = $ColorPalette.blue[2]
                    "markerColor" = $ColorPalette.blue[2]
                    "markerStyle" = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                    "lineWeight"  = 3
                    "markerSize"  = 5
                }
                4 = @{
                    "color"       = $ColorPalette.orange[2]
                    "markerColor" = $ColorPalette.orange[2]
                    "markerStyle" = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                    "lineWeight"  = 3
                    "markerSize"  = 5
                }
                5 = @{
                    "color"       = $ColorPalette.blue[1]
                    "markerColor" = $ColorPalette.blue[1]
                    "markerStyle" = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                    "lineWeight"  = 3
                    "markerSize"  = 5
                }
                6 = @{
                    "color"       = $ColorPalette.orange[1]
                    "markerColor" = $ColorPalette.orange[1]
                    "markerStyle" = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                    "lineWeight"  = 3
                    "markerSize"  = 5
                }
            }
        } 
        else {
            if ($meta.usedCustomNames) {
                $table.chartSettings.title = "$baselineName - $cappedProp"
            }
            $table.chartSettings.seriesSettings = @{
                1 = @{
                    "color"       = $ColorPalette.blue[3]
                    "markerColor" = $ColorPalette.blue[3]
                    "markerStyle" = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                    "lineWeight"  = 3
                    "markerSize"  = 5
                }
                2 = @{
                    "color"       = $ColorPalette.blue[2]
                    "markerColor" = $ColorPalette.blue[2]
                    "markerStyle" = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                    "lineWeight"  = 3
                    "markerSize"  = 5
                }
                3 = @{
                    "color"       = $ColorPalette.blue[1]
                    "markerColor" = $ColorPalette.blue[1]
                    "markerStyle" = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                    "lineWeight"  = 3
                    "markerSize"  = 5
                }
            }
        }

        if (-not $innerPivot) {
            $table.chartSettings.yOffset = 3
        }

        $col = 0
        $row = 0
        foreach ($IPivotKey in $meta.innerPivotKeys | Sort) {
            # Add column labels to table
            $table.cols.$tableTitle.$innerPivot.$IPivotKey = $col
            $table.meta.numWrites += 1
            $table.data.$tableTitle.$innerPivot.$IPivotKey = @{
                $prop = @{}
            }
            $table.meta.columnFormats += $meta.format.$prop
            $col += 1
        
            # Add row labels and fill data in table
            foreach ($metric in $metrics) {
                if (-not ($table.rows.$prop.Keys -contains $metric)) { 
                    if (-not $meta.comparison) {
                        $table.rows.$prop.$metric = $row
                        $row += 1
                        $table.meta.numWrites += 1
                    } 
                    else {
                        $table.meta.numWrites += 3
                        $table.rows.$prop.$metric = @{
                            $baselineName = $row
                            $testName     = $row + 1
                        } 
                        $row += 2
                    }
                }
                if (-not ($table.data.$tableTitle.$innerPivot.$IPivotKey.$prop.Keys -contains $metric)) {
                    $table.data.$tableTitle.$innerPivot.$IPivotKey.$prop.$metric = @{}
                }

                if (-not $meta.comparison) {
                    $table.data.$tableTitle.$innerPivot.$IPivotKey.$prop.$metric = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.$metric}
                } 
                else {
                    $table.data.$tableTitle.$innerPivot.$IPivotKey.$prop.$metric.$baselineName = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.$metric}
                    $table.data.$tableTitle.$innerPivot.$IPivotKey.$prop.$metric.$testName     = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.$metric}
                } 
            }
            Write-Progress -Activity "Formatting Tables" -Status "MinMeanMax Table" -Id 3 -PercentComplete (100 * (($completeIters++) / $numIters))
        }
        $table.meta.dataWidth     = Get-TreeWidth $table.cols
        $table.meta.colLabelDepth = Get-TreeDepth $table.cols
        $table.meta.dataHeight    = Get-TreeWidth $table.rows
        $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
        $table.meta.numWrites    += $table.meta.dataHeight * $table.meta.dataWidth 
        $tables = $tables + $table
    }

    if (($tables.Count -gt 0) -and (-not $NoNewWorksheets)) {
        $sheetTitle = Get-WorksheetTitle -BaseName "MinMeanMax" -OuterPivot $outerPivot -OPivotKey $OPivotKey
        $tables = @($sheetTitle) + $tables
    }
    Write-Progress -Activity "Formatting Tables" -Status "MinMeanMax Table" -Id 3 -PercentComplete 100

    return $tables
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
function Format-Percentiles {
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter(Mandatory=$true)] $OPivotKey,

        [Parameter()] [String] $Tool = "",

        [Parameter()] [switch] $NoNewWorksheets
    )

    $tables     = @()
    $data       = $DataObj.data
    $meta       = $DataObj.meta
    $innerPivot = $meta.InnerPivot
    $outerPivot = $meta.OuterPivot 
    $baselineName = $meta.datasetNames["baseline"]
    $testName = $meta.datasetNames["test"]
  
    
    $numIters = $meta.props.Count * $meta.innerPivotKeys.Count
    $completeIters = 0


    foreach ($prop in $meta.props) {
        foreach ($IPivotKey in $meta.innerPivotKeys | Sort) { 

            if ($innerPivot) { 
                $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Percentiles - $IPivotKey $innerPivot")
                $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -DatasetName $baselineName -Comparison $meta.comparison -UsedCustomName $meta.usedCustomNames.baseline
            } 
            else {
                $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Percentiles")
                $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -DatasetName $baselineName -Comparison $meta.comparison -UsedCustomName $meta.usedCustomNames.baseline
            }
            
            if ((-not $meta.comparison) -and $meta.usedCustomNames.baseline) {
                $chartTitle = "$baselineName - $chartTitle"
            }

            $table = @{
                "rows" = @{
                    "percentiles" = @{}
                }
                "cols" = @{
                    $tableTitle = @{
                        $prop = 0
                    }
                }
                "meta" = @{
                    "columnFormats" = @($meta.format.$prop)
                    "rightAlign"    = [Array] @(2)
                    "name"          = "Percentiles"
                    "numWrites"     = 1 + 2
                }
                "data" = @{
                    $tableTitle = @{
                        $prop = @{
                            "percentiles" = @{}
                        }
                    }
                }
                "chartSettings" = @{
                    "title"     = $chartTitle
                    "yOffset"   = 1
                    "xOffset"   = 1
                    "chartType" = [Excel.XlChartType]::xlXYScatterLinesNoMarkers
                    "seriesSettings" = @{
                        1 = @{ 
                            "color"      = $ColorPalette.blue[2]
                            "lineWeight" = 3
                        }
                    }
                    "axisSettings" = @{
                        1 = @{
                            "max"            = 100
                            "title"          = "Percentiles"
                            "minorGridlines" = $true
                        }
                        2 = @{
                            "title" = $meta.units[$prop]
                        }
                    }
                }
            }

            $table.chartSettings.axisSettings[2].logarithmic = Set-Logarithmic -Data $data -OPivotKey $OPivotKey -Prop $prop -IPivotKey $IPivotKey -Meta $meta 

            if ($meta.comparison) {
                $table.meta.numWrites += 3
                $table.cols.$tableTitle.$prop = @{
                    $baselineName = 0
                    "% change" = 1
                    $testName     = 2
                }
                $table.data.$tableTitle.$prop = @{
                    $baselineName = @{
                        "percentiles" = @{}
                    }
                    "% change" = @{
                        "percentiles" = @{}
                    }
                    $testName = @{
                        "percentiles" = @{}
                    }
                }
                $table.chartSettings.seriesSettings[2] = @{
                    "delete" = $true
                }
                $table.chartSettings.seriesSettings[3] = @{
                    "color"      = $ColorPalette.orange[2]
                    "lineWeight" = 3
                }
                $table.meta.columnFormats = @($meta.format.$prop, "0.0%", $meta.format.$prop)
            }
            $row = 0

            $keys = @()
            if ($data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles.Keys.Count -gt 0) {
                $keys = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles.Keys
            } 
            else {
                $keys = $data.$OPivotKey.$prop.$IPivotKey.test.percentiles.Keys
            }

            # Add row labels and fill data in table
            
            $sortedKeys = Sort-StringsAsNumbers -Arr $keys 

            foreach ($percentile in $sortedKeys) {
                $table.rows.percentiles.$percentile = $row
                $table.meta.numWrites += 1
                if ($meta.comparison) { 
                    if ($data.$OPivotKey.$prop.$IPivotKey.ContainsKey("baseline")) {
                        $table.data.$tableTitle.$prop.$baselineName.percentiles[$percentile]   = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles.$percentile}
                    }
                    if ($data.$OPivotKey.$prop.$IPivotKey.ContainsKey("test")) {
                        $table.data.$tableTitle.$prop.$testName.percentiles[$percentile]       = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.test.percentiles.$percentile}
                    }
                    if ($data.$OPivotKey.$prop.$IPivotKey.ContainsKey("baseline") -and $data.$OPivotKey.$prop.$IPivotKey.ContainsKey("test")) {
                        $table.data.$tableTitle.$prop."% change".percentiles[$percentile] = @{"value" = "=IF([col-1][row]=0, `"--`", ([col+1][row]-[col-1][row])/ABS([col-1][row]))"}
                        $params = @{
                            "Cell"    = $table.data.$tableTitle.$prop."% change".percentiles[$percentile]
                            "TestVal" = $data.$OPivotKey.$prop.$IPivotKey.test.percentiles[$percentile]
                            "BaseVal" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[$percentile]
                            "Goal"    = $meta.goal.$prop
                        }
                        $table.data.$tableTitle.$prop."% change".percentiles[$percentile] = Set-CellColor @params
                    } 
                } 
                else {
                    $table.data.$tableTitle.$prop.percentiles[$percentile] = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles.$percentile}
                }
                $row += 1
                $nextRow += 1
                

            }
            Write-Progress -Activity "Formatting Tables" -Status "Percentiles Table" -Id 3 -PercentComplete (100 * (($completeIters++) / $numIters))
            $nextRow += $HeaderRows

            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows 
            $table.meta.numWrites    += $table.meta.dataHeight * $table.meta.dataWidth  
            $tables = $tables + $table
        }
    }

    if (($tables.Count -gt 0) -and (-not $NoNewWorksheets)) {
        $sheetTitle = Get-WorksheetTitle -BaseName "Percentiles" -OuterPivot $outerPivot -OPivotKey $OPivotKey
        $tables     = @($sheetTitle) + $tables 
    }
    Write-Progress -Activity "Formatting Tables" -Status "Done" -Id 3 -PercentComplete 100

    return $tables  
}


function Sort-ByIndex ($Arr) 
{
    $subArrs = [System.Collections.ArrayList] @()

    for ($idx = 0; $idx -lt $Arr.Count; $idx++) {
        $null = $subArrs.Add([System.Collections.ArrayList] @($idx))
    }

    while ($subArrs.Count -gt 1) {
        $newArr = [System.Collections.ArrayList] @()

        for ($i = 0; $i -lt $subArrs.Count; $i += 2) {
            if ($i -lt $subArrs.Count - 1) {
                $merged = Merge-Arrays -IdxArr1 $subArrs[$i] -IdxArr2 $subArrs[$i + 1] -ValArr $Arr
                $null = $newArr.Add($merged)
            } else {
                $null = $newArr.Add($subArrs[$i])
            }
            
        }
        $subArrs = $newArr
    }
    return $subArrs
}

function Merge-Arrays ($IdxArr1, $IdxArr2, $ValArr) {
    $i = 0
    $j = 0
    $newArr = [System.Collections.ArrayList] @()

    while (($i -lt $IdxArr1.Count) -and ($j -lt $IdxArr2.Count)) {
        $idx1 = $IdxArr1[$i]
        $idx2 = $IdxArr2[$j]

        if ( $ValArr[$idx1] -le $ValArr[$idx2] ) {
            $null = $newArr.Add($idx1)
            $i++
        } else {
            $null = $newArr.Add($idx2)
            $j++
        } 
    }

    while ($i -lt $IdxArr1.Count) {
        $idx = $IdxArr1[$i]
        $null = $newArr.Add($idx)
        $i++
    }
    while ($j -lt $IdxArr2.Count) {
        $idx = $IdxArr2[$j]
        $null = $newArr.Add($idx)
        $j++
    }

    return $newArr
}

function Sort-StringsAsNumbers ([Array] $Arr) {
    $tempArr = @()
    foreach ($item in $Arr) {
        $tempArr += [decimal] $item 
    }

    $sortedIndices = Sort-ByIndex -Arr $tempArr

    $outArr = @() 
    foreach ($idx in $sortedIndices) {
        $outArr += $Arr[$idx]
    }

    return $outArr 
}

function Set-Logarithmic ($Data, $OPivotKey, $Prop, $IPivotKey, $Meta) {
    if ($data.$OPivotKey.$Prop.$IPivotKey.baseline.stats) {
        if ($data.$OPivotKey.$Prop.$IPivotKey.baseline.stats.min -le 0) {
            return $false
        }
        if ($Meta.comparison) {
            if ($data.$OPivotKey.$Prop.$IPivotKey.test.stats.min -le 0) {
                return $false
            }
        }
        if (($data.$OPivotKey.$Prop.$IPivotKey.baseline.stats.max / ($data.$OPivotKey.$Prop.$IPivotKey.baseline.stats.min + $EPS)) -gt 10) {
            return $true
        }

        if ($Meta.comparison) {
            if (($data.$OPivotKey.$Prop.$IPivotKey.test.stats.max / ($data.$OPivotKey.$Prop.$IPivotKey.test.stats.min + $EPS)) -gt 10) {
                return $true
            }
            if (($data.$OPivotKey.$Prop.$IPivotKey.test.stats.max / ($data.$OPivotKey.$Prop.$IPivotKey.baseline.stats.min + $EPS)) -gt 10) {
                return $true
            }
            if (($data.$OPivotKey.$Prop.$IPivotKey.baseline.stats.max / ($data.$OPivotKey.$Prop.$IPivotKey.test.stats.min + $EPS)) -gt 10) {
                return $true
            }
        }
    }
    return $false
}

<#
.SYNOPSIS
    Returns a template for Format-Histogram
#>
function Get-HistogramTemplate {
    param(
        [PSObject[]] $DataObj,
        [String] $TableTitle,
        [String] $Property,
        [String] $IPivotKey
    )

    $meta = $DataObj.meta
    $baselineName = $meta.datasetNames["baseline"]
    $testName = $meta.datasetNames["test"]

    $chartTitle = if ($IPivotKey) {
        "$Property Histogram - $IPivotKey $($meta.InnerPivot)"
    } else {
        "$Property Histogram"
    }

    if ((-not $meta.comparison) -and $meta.usedCustomNames.baseline) {
        $chartTitle = "$baselineName - $chartTitle"
    }

    $table = @{
        "rows" = @{
            "histogram buckets" = @{}
        }
        "cols" = @{
            $TableTitle = @{
                $Property = 0
            }
        }
        "meta" = @{
            "rightAlign" = [Array] @(2)
            "columnFormats" = @("0.0%")
            "name"          = "Histogram"
            "numWrites"     = 1 + 2
        }
        "data" = @{
            $TableTitle = @{
                $Property = @{
                    "histogram buckets" = @{}
                }
            }
        }
        "chartSettings"= @{
            "title"   = $TextInfo.ToTitleCase($chartTitle)
            "yOffset" = 1
            "xOffset" = 1
            "seriesSettings" = @{
                1 = @{ 
                    "color" = $ColorPalette.blue[2]
                    "lineWeight" = 1
                    "name" = "Frequency"
                }
            }
            "axisSettings" = @{
                1 = @{
                    "title" = "$Property ($($meta.units[$Property]))"
                    "tickLabelSpacing" = 5
                }
                2 = @{
                    "title" = "Frequency"
                }
            }
        } # chartSettings
    }

    # Support base/test comparison mode
    if ($meta.comparison) {
        $table.cols.$TableTitle.$Property = @{
            $baselineName = 0
            "% change" = 1
            $testName     = 2
        }
        
        $table.meta.numWrites += 3
        $table.data.$TableTitle.$Property = @{
            $baselineName = @{
                "histogram buckets" = @{}
            }
            "% change" = @{
                "histogram buckets" = @{}
            }
            $testName = @{
                "histogram buckets" = @{}
            }
        }

        $table.chartSettings.seriesSettings[1].name = $baselineName
        $table.chartSettings.seriesSettings[2] = @{
            "delete" = $true # don't plot % change
        }
        $table.chartSettings.seriesSettings[3] = @{
            "color"      = $ColorPalette.orange[2]
            "name"       = $testName
            "lineWeight" = 1
        }

        $table.meta.columnFormats = @("0.0%", "0.0%", "0.0%")
    }

    return $table
} # Get-HistogramTemplate

<#
.SYNOPSIS
    Outputs a table with a histogram and chart.
#>
function Format-Histogram {
    param(
        [Parameter(Mandatory=$true)]
        [PSObject[]] $DataObj,

        [Parameter(Mandatory=$true)]
        $OPivotKey,

        [Parameter(Mandatory=$true)]
        [String] $Tool
    )

    $tables = @()
    $meta = $DataObj.meta
    $baselineName = $meta.datasetNames["baseline"]
    $testName = $meta.datasetNames["test"]

    foreach ($prop in $DataObj.data.$OPivotKey.Keys) {
        foreach ($iPivotKey in $DataObj.data.$OPivotKey.$prop.Keys | sort) {
            $data = $DataObj.data.$OPivotKey.$prop.$iPivotKey

            if (-not $data.baseline.histogram -and -not $data.test.histogram) {
                continue
            }

            $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $meta.OuterPivot -OPivotKey $OPivotKey -InnerPivot $meta.InnerPivot -IPivotKey $iPivotKey -Comparison $meta.comparison -UsedCustomName $meta.usedCustomNames.baseline -DatasetName $baselineName 
            $table = Get-HistogramTemplate -DataObj $DataObj -TableTitle $tableTitle -Property $prop -IPivotKey $iPivotKey

            if ($data.baseline.histogram) {
                $baseSum = ($data.baseline.histogram.Values | measure -Sum).Sum
            }
            
            if ($data.test.histogram) {
                $testSum = ($data.test.histogram.Values | measure -Sum).Sum
            }

            # Add row labels and fill data in table
            $row = 0
            $buckets = if ($data.baseline.histogram.Keys.Count -gt 0) {$data.baseline.histogram.Keys} else {$data.test.histogram.Keys}
            foreach ($bucket in ($buckets | sort)) {
                $table.rows."histogram buckets".$bucket = $row
                $table.meta.numWrites += 1
                
                

                if (-not $meta.comparison) {
                    $baseVal = $data.baseline.histogram.$bucket / $baseSum
                    $table.data.$tableTitle.$prop."histogram buckets"[$bucket] = @{"value" = $baseVal}
                } else {
                    if ($data.baseline.histogram) {
                        $baseVal = $data.baseline.histogram.$bucket / $baseSum
                        $table.data.$tableTitle.$prop.$baselineName."histogram buckets"[$bucket]   = @{"value" = $baseVal}
                    }
                    if ($data.test.histogram) { 
                        $testVal = $data.test.histogram.$bucket / $testSum
                        $table.data.$tableTitle.$prop.$testName."histogram buckets"[$bucket]       = @{"value" = $testVal}
                    }
                    if ($data.baseline.histogram -and $data.test.histogram) {
                        $baseCell = "C$($row + $HeaderRows)"
                        $testCell = "E$($row + $HeaderRows)"
                        $table.data.$tableTitle.$prop."% change"."histogram buckets"[$bucket] = @{"value" = "=IF([col-1][row]=0, `"--`", ([col+1][row]-[col-1][row])/ABS([col-1][row]))"}
                        $table.data.$tableTitle.$prop."% change"."histogram buckets"[$bucket] = Set-CellColor -Cell $table.data.$tableTitle.$prop."% change"."histogram buckets"[$bucket] -BaseVal $baseVal -TestVal $testVal -Goal "increase"
                    }
                }

                $row += 1
            }

            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            
            $table.meta.numWrites += $table.meta.dataHeight * $table.meta.dataWidth  
            $tables = $tables + $table
        }
    }

    if ($table.Count -gt 0) {
        $worksheetTitle = Get-WorksheetTitle -BaseName "Histogram" -OuterPivot $meta.OuterPivot -OPivotKey $OPivotKey
        $tables = @($worksheetTitle) + $tables
    }

    return $tables
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
function Format-Distribution {
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter()] [string] $OPivotKey,

        [Parameter()] [String] $Tool = "",

        [Parameter()] [String] $Prop,

        [Parameter()] [Int] $SubSampleRate = -1,

        [Parameter()] [switch] $NoNewWorksheets
        
    )

    $DEFALT_SEGMENTS_TARGET = 200

    $meta  = $DataObj.meta 
    $modes = if ($meta.comparison) { @("baseline", "test") } else { @(,"baseline") } 
    $tables     = @()
    $innerPivot = $meta.InnerPivot
    $outerPivot = $meta.OuterPivot

    $baselineName = $meta.datasetNames["baseline"]
    $testName = $meta.datasetNames["test"]

    $NumSamples = Calculate-MaxNumSamples -RawData $DataObj.rawData -Modes $modes -Prop $Prop
    if ($SubSampleRate -eq -1) {
        $SubSampleRate = [Int] ($NumSamples/$DEFALT_SEGMENTS_TARGET)
    } 
    $numIters = Calculate-NumIterations -Distribution -DataObj $dataObj -Prop $Prop -SubSampleRate $SubSampleRate
    $j = 0

    foreach ($IPivotKey in $DataObj.data.$OPivotKey.$Prop.Keys) { 
        foreach ($mode in $modes) { 
            if (-Not $DataObj.data.$OPivotKey.$Prop.$IPivotKey.$mode.stats) {
                continue
            } 

            
            $logarithmic = Set-Logarithmic -Data $dataObj.data -OPivotKey $OPivotKey -Prop $Prop -IPivotKey $IPivotKey `
                                            -Meta $meta
            $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey -DatasetName $baselineName -Comparison $meta.comparison -UsedCustomName $meta.usedCustomNames.baseline
            $data       = $dataObj.rawData.$mode 
            $table = @{
                "meta" = @{
                    "name" = "Distribution"
                    "numWrites" = 1 + 3
                }
                "rows" = @{
                    "Data Point" = @{}
                }
                "cols" = @{
                    $tableTitle = @{
                        "Time Segment" = 0
                        $Prop          = 1
                    }
                }
                "data" = @{
                    $tableTitle = @{
                        "Time Segment" = @{
                            "Data Point" = @{}
                        }
                        $Prop = @{
                            "Data Point" = @{}
                        }
                    }
                }
                "chartSettings" = @{
                    "chartType" = [Excel.XlChartType]::xlXYScatter
                    "yOffset"   = 2
                    "xOffset"   = 2
                    "title"     = "Temporal $prop Distribution"
                    "axisSettings" = @{
                        1 = @{
                            "title"          = "Time Series"
                            "minorGridlines" = $true
                            "majorGridlines" = $true
                            "max"            = $NumSamples
                        }
                        2 = @{
                            "title"       = $meta.units.$Prop
                            "logarithmic" = $logarithmic
                            "min"         = 10
                        }
                    }
                }
            }

            if ($mode -eq "baseline") {
                $table.chartSettings.seriesSettings = @{
                    1 = @{
                            "markerStyle"           = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                            "markerBackgroundColor" = $ColorPalette.blue[3]
                            "markerForegroundColor" = $ColorPalette.blue[2]
                            "name"                  = "$Prop Sample" 
                        }
                }
            } else {
                $table.chartSettings.seriesSettings = @{
                    1 = @{
                            "markerStyle"           = [Excel.XlMarkerStyle]::xlMarkerStyleCircle
                            "markerBackgroundColor" = $ColorPalette.blue[3]
                            "markerForegroundColor" = $ColorPalette.blue[2]
                            "name"                  = "$Prop Sample"
                        }
                }
            }

            # Add row labels and fill data in table
            $i   = 0
            $row = 0

            if ($SubSampleRate -gt 0) { 
                $finished = $false
                while (-Not $finished) {
                    [Array]$segmentData = @()
                    foreach ($entry in $data) {
                        if ($entry.$Prop.GetType().Name -ne "Object[]") {
                            continue
                        }
                        if (((-not $innerPivot) -or ($entry.$innerPivot -eq $IPivotKey)) -and `
                                ((-not $outerPivot) -or ($entry.$outerPivot -eq $OPivotKey)) -and `
                                    ($i * $SubSampleRate -lt $entry.$Prop.Count)) {
                            $finalIdx = (($i + 1) * $SubSampleRate) - 1
                            if (((($i + 1) * $SubSampleRate) - 1) -ge $entry.$Prop.Count) {
                                $finalIdx = $entry.$Prop.Count - 1
                            }
                            $segmentData += $entry.$Prop[($i * $SubSampleRate) .. $finalIdx]
                        }
                    }
                    $segmentData = $segmentData | Sort
                    $time        = $i * $subSampleRate
                    if ($segmentData.Count -ge 5) {
                        $table.rows."Data Point".$row       = $row
                        $table.rows."Data Point".($row + 1) = $row + 1
                        $table.rows."Data Point".($row + 2) = $row + 2
                        $table.rows."Data Point".($row + 3) = $row + 3
                        $table.rows."Data Point".($row + 4) = $row + 4
                        $table.data.$tableTitle."Time Segment"."Data Point".$row       = @{"value" = $time}
                        $table.data.$tableTitle."Time Segment"."Data Point".($row + 1) = @{"value" = $time}
                        $table.data.$tableTitle."Time Segment"."Data Point".($row + 2) = @{"value" = $time}
                        $table.data.$tableTitle."Time Segment"."Data Point".($row + 3) = @{"value" = $time}
                        $table.data.$tableTitle."Time Segment"."Data Point".($row + 4) = @{"value" = $time}
                        $table.data.$tableTitle.$Prop."Data Point".$row = @{"value"       = $segmentData[0]}
                        $table.data.$tableTitle.$Prop."Data Point".($row + 1) = @{"value" = $segmentData[[int]($segmentData.Count / 4)]}
                        $table.data.$tableTitle.$Prop."Data Point".($row + 2) = @{"value" = $segmentData[[int]($segmentData.Count / 2)]}
                        $table.data.$tableTitle.$Prop."Data Point".($row + 3) = @{"value" = $segmentData[[int]((3 * $segmentData.Count) / 4)]}
                        $table.data.$tableTitle.$Prop."Data Point".($row + 4) = @{"value" = $segmentData[-1]}
                        $row += 5
                        $table.meta.numWrites += 5
                    } 
                    elseif ($segmentData.Count -ge 1){
                        foreach ($sample in $segmentData) {
                            $table.rows."Data Point".$row = $row
                            $table.data.$tableTitle."Time Segment"."Data Point".$row = @{"value" = $time}
                            $table.data.$tableTitle.$Prop."Data Point".$row          = @{"value" = $sample}
                            $row++
                            $table.meta.numWrites += 1
                        }
                    } else {
                        $finished = $true
                    }
                    $i++

                    Write-Progress -Activity "Formatting Tables" -Status "Distribution Table" -Id 3 -PercentComplete (100 * (($j++) / $numIters))

                }
            } else {
                $finished = $false
                while (-not $finished) { 
                    [Array]$segmentData = @()
                    foreach ($entry in $data) {
                        if ($entry.$prop.GetType().Name -ne "Object[]") {
                            continue
                        }
                        if (((-not $innerPivot) -or ($entry.$innerPivot -eq $IPivotKey)) -and ((-not $outerPivot) -or ($entry.$outerPivot -eq $OPivotKey))) {
                            if ($null -eq $entry[$Prop][$i]) {
                                continue
                            } 
                            $segmentData += $entry[$Prop][$i]
                        }
                    }
                    
                    $finished = ($segmentData.Count -eq 0) 
                    foreach ($sample in $segmentData) {
                        $table.rows."Data Point".$row = $row
                        $table.data.$tableTitle."Time Segment"."Data Point".$row = @{"value" = $i}
                        $table.data.$tableTitle.$Prop."Data Point".$row          = @{"value" = $sample}
                        $row++
                        $table.meta.numWrites += 1
                    }
                    $i++
                    Write-Progress -Activity "Formatting Tables" -Status "Distribution Table" -Id 3 -PercentComplete (100 * (($j++) / $numIters))

                }
            }
            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $table.meta.numWrites += $table.meta.dataHeight * $table.meta.dataWidth 
            if (-not $NoNewWorksheets) {
                if ($modes.Count -gt 1) {
                    if ($mode -eq "baseline") {
                        $worksheetName = Get-WorksheetTitle -BaseName "Base Distr." -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey -Prop $Prop
                    } 
                    else {
                        $worksheetName = Get-WorksheetTitle -BaseName "Test Distr." -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey -Prop $Prop
                    } 
                } 
                else {
                    $worksheetName = Get-WorksheetTitle -BaseName "Distr." -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey -Prop $Prop
                } 
                $tables += $worksheetName
            }

            $tables += $table
        }
    }
    
    Write-Progress -Activity "Formatting Tables" -Status "Distribution Table" -Id 3 -PercentComplete 100

    return $tables
}


<#
.SYNOPSIS 
    Calculates the maximum numer of samples of a given property provided
    by a single data file
#>
function Calculate-MaxNumSamples ($RawData, $Modes, $Prop) {
    $max = 0
    foreach ($mode in $Modes) {
        foreach ($fileEntry in $RawData.$mode) {
            if ($fileEntry.$Prop.Count -gt $max) {
                $max = $fileEntry.$Prop.Count
            }
        } 
    }
    $max
}
function Calculate-NumIterations {
    param (
        [Parameter(Mandatory=$true, ParameterSetName="distribution")]
        [Switch] $Distribution,

        [Parameter(Mandatory=$true)]
        $DataObj, 

        [Parameter(Mandatory=$true, ParameterSetName = "distribution")]
        [String] $Prop, 

        [Parameter(Mandatory=$true, ParameterSetName = "distribution")]
        [Int] $SubSampleRate

         
    )

    if ($Distribution) { 
        $innerLoopIters = 0
        $maxSamples = Calculate-MaxNumSamples -RawData $DataObj.rawData -Modes @("baseline") -Prop $Prop
        if ($SubsampleRate -gt 0) { 
            $innerLoopIters += 1 + [Int]( ($maxSamples / $SubsampleRate) + 0.5) 
        } else {
            $innerLoopIters += $maxSamples
        }

        
        if ($dataObj.meta.comparison) {
            $maxSamples = Calculate-MaxNumSamples -RawData $DataObj.rawData -Modes @("test") -Prop $Prop
            if ($SubsampleRate -gt 0) {  
                $innerLoopIters += 1 + [Int](($maxSamples/ $SubSampleRate) + 0.5) 
            } else {
                $innerLoopIters += $maxSamples + 1
            }
        } 
        return $DataObj.meta.innerPivotKeys.Count * $innerLoopIters
    }
}


<#
.SYNOPSIS
    Sets the colors of a cell, indicating whether a test value shows
    an improvement when compared to a baseline value. Improvement is
    defined by the goal (increase/decrease) for the given value.
.PARAMETER Cell
    Object containg a cell's value and other settings.
.PARAMETER TestVal
    Test metric value.
.PARAMETER BaseVal
    Baseline metric value.
.PARAMETER Goal
    Defines metric improvement direction. "increase", "decrease", or "none".
#>
function Set-CellColor ($Cell, [Decimal] $TestVal, [Decimal] $BaseVal, $Goal) {
    if (($Goal -ne "none") -and ($TestVal -ne $BaseVal)) {
        if (($Goal -eq "increase") -eq ($TestVal -gt $BaseVal)) {
            $Cell["fontColor"] = $ColorPalette.Green
            $Cell["cellColor"] = $ColorPalette.LightGreen
        } else {
            $Cell["fontColor"] = $ColorPalette.Red
            $Cell["cellColor"] = $ColorPalette.LightRed
        }
    }

    return $Cell
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
function Get-TreeWidth ($Tree) {
    if ($Tree.GetType().Name -eq "Int32") {
        return 1
    }
    $width = 0
    foreach ($key in $Tree.Keys) {
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
function Get-TreeDepth ($Tree) {
    if ($Tree.GetType().Name -eq "Int32") {
        return 0
    }
    $depths = @()
    foreach ($key in $Tree.Keys) {
        $depths = $depths + [int](Get-TreeDepth -Tree $Tree[$key])
    }
    return ($depths | Measure -Maximum).Maximum + 1
}
