$LIGHTGREEN = 10416289
$GREEN = 1268766
$LIGHTRED = 10396159
$RED = 2108032
$BLUES = @(10249511, 14058822, 16758932)
$ORANGES = @(294092, 1681916, 6014716)

$EPS = 0.0001

$WorksheetMaxLen = 31

$ABBREVIATIONS = @{
    "sessions" = "sess."
    "bufferLen" = "bufLen."
    "bufferCount" = "bufCt."
    "protocol" = "protocol"
    "sendMethod" = "sndMthd" 
}

$THROUGHPUTS = @(1, 10, 25, 40, 50, 100, 200, 400)


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
function Format-RawData {
    param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter(Mandatory=$true)] $OPivotKey,

        [Parameter()] [String] $Tool = "",

        [Parameter()] [switch] $NoNewWorksheets
    )

    try {
        $legend = @{
            "meta" = @{
                "colLabelDepth" = 1
                "rowLabelDepth" = 1
                "dataWidth"     = 2
                "dataHeight"    = 3 
            }
            "rows" = @{
                " "   = 0
                "  "  = 1
                "   " = 2
            }
            "cols" = @{
                "legend" = @{
                    " "  = 0
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
                            "value"     = "Improvement"
                            "fontColor" = $GREEN
                            "cellColor" = $LIGHTGREEN
                        }
                        "   " = @{
                            "value"     = "Regression"
                            "fontColor" = $RED
                            "cellColor" = $LIGHTRED
                        }
                    } 
                }
            }
        }

        $meta       = $DataObj.meta
        $innerPivot = $meta.InnerPivot
        $outerPivot = $meta.OuterPivot
        $tables     = @()

        if (-not $NoNewWorksheets) {
            $tables += Get-WorksheetTitle -BaseName "Raw Data" -OuterPivot $outerPivot -OPivotKey $OPivotKey
        }
        if ($meta.comparison) {
            $tables += $legend
        }

        # Fill single array with all data and sort, label data as baseline/test if necessary
        [Array] $data = @() 
        foreach ($entry in $DataObj.rawData.baseline) {
            if ($meta.comparison) {
                $entry.baseline = $true
            } 
            if (($OPivotKey -eq $NoPivot) -or ($entry.$outerPivot -eq $OPivotKey)) {
                $data += $entry
            }
        }

        if ($meta.comparison) {
            foreach ($entry in $DataObj.rawData.test) {
                if (($OPivotKey -eq $NoPivot) -or ($entry.$outerPivot -eq $OPivotKey)) {
                    $data += $entry
                }
            }
        }

        if ($innerPivot) {
            $data = Sort-ByProp -Data $data -Prop $innerPivot
        }
        

        foreach ($prop in $dataObj.data.$OPivotKey.Keys) {
            $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey

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
                    "leftAlign"     = [Array] @(2)
                }
                "data"  = @{
                    $tableTitle = @{
                        $innerPivot = @{}
                    }
                }
            }
            $col = 0
            $row = 0

            foreach ($entry in $data) {
                $iPivotKey = $NoPivot
                if ($innerPivot) {
                    $iPivotKey = $entry.$innerPivot
                }

                # Add column labels to table
                if (-not ($table.cols.$tableTitle.$innerPivot.Keys -contains $iPivotKey)) {
                    if ($meta.comparison) {
                        $table.cols.$tableTitle.$innerPivot.$iPivotKey = @{
                            "baseline" = $col
                            "test"     = $col + 1
                        }
                        $table.meta.columnFormats += @($meta.format.$prop, $meta.format.$prop) 
                        $col += 2
                        $table.data.$tableTitle.$innerPivot.$iPivotKey = @{
                            "baseline" = @{
                                $prop = @{}
                            }
                            "test" = @{
                                $prop = @{}
                            }
                        }
                    } 
                    else {
                        $table.meta.columnFormats += $meta.format.$prop
                        $table.cols.$tableTitle.$innerPivot.$iPivotKey = $col
                        $table.data.$tableTitle.$innerPivot.$iPivotKey = @{
                            $prop = @{}
                        }
                        $col += 1
                    }
                }

                # Add row labels and fill data in table
                $filename = $entry.fileName.Split('\')[-2] + "\" + $entry.fileName.Split('\')[-1] 
                while ($table.rows.$prop.keys -contains $filename) {
                    $filename += "*"
                }
                $table.rows.$prop.$filename = $row
                
                $row += 1
                if ($meta.comparison) {
                    if ($entry.baseline) {
                        $table.data.$tableTitle.$innerPivot.$iPivotKey.baseline.$prop.$filename = @{
                            "value" = $entry.$prop
                        }
                    }
                    else {
                        $table.data.$tableTitle.$innerPivot.$iPivotKey.test.$prop.$filename = @{
                            "value" = $entry.$prop
                        }
                        $params = @{
                            "Cell"    = $table.data.$tableTitle.$innerPivot.$iPivotKey.test.$prop.$filename
                            "TestVal" = $entry.$prop
                            "BaseVal" = $DataObj.data.$OPivotKey.$prop.$iPivotKey.baseline.stats.mean
                            "Goal"    = $meta.goal.$prop
                        }
                        
                        $table.data.$tableTitle.$innerPivot.$iPivotKey.test.$prop.$filename = Select-Color @params
                    }
                } 
                else {
                    $table.data.$tableTitle.$innerPivot.$iPivotKey.$prop.$filename = @{
                        "value" = $entry.$prop
                    }
                }
            }
            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
        }

        foreach ($entry in $data) {
            if ($entry.baseline) {
                $entry.Remove("baseline")
            }
        }
        return $tables
    } 
    catch {
        Write-Warning "Error at Format-RawData"
        Write-Error $_.Exception.Message
    }
}

function Get-WorksheetTitle ($BaseName, $OuterPivot, $OPivotKey, $InnerPivot, $IPivotKey) {
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
        return "$BaseName"
    }
}

function Get-TableTitle ($Tool, $OuterPivot, $OPivotKey, $InnerPivot, $IPivotKey) { 
    if ($OuterPivot -and $InnerPivot) {
        $OAbv = $ABBREVIATIONS[$OuterPivot]
        $IAbv = $ABBREVIATIONS[$InnerPivot]

        return "$Tool - $OPivotKey $OAbv - $IPivotKey $IAbv"
    } 
    elseif ($OuterPivot) {
        $OAbv = $ABBREVIATIONS[$OuterPivot]

        return "$Tool - $OPivotKey $OAbv"
    } 
    elseif ($InnerPivot) {
        $IAbv = $ABBREVIATIONS[$InnerPivot]

        return "$Tool - $IPivotKey $IAbv"
    }
    else {
        return "$Tool"
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
function Format-Stats {
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter(Mandatory=$true)] $OPivotKey,

        [Parameter()] $Metrics = $null,

        [Parameter()] [String] $Tool = "",

        [Parameter()] [switch] $NoNewWorksheets
    )
    
    try {
        $tables = @()
        $data = $DataObj.data
        $meta = $DataObj.meta
        $innerPivot = $meta.InnerPivot
        $outerPivot = $meta.OuterPivot

        foreach ($prop in $data.$OPivotKey.keys) {
            $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey

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
                }
                "data" = @{
                    $tableTitle = @{
                        $innerPivot = @{}
                    }
                }
            }

            $col = 0
            $row = 0
            $noStats = $false
            foreach ($IPivotKey in $data.$OPivotKey.$prop.Keys | Sort) {
                if (-Not ($data.$OPivotKey.$prop.$IPivotKey.baseline.Keys -contains "stats")) {
                    $noStats = $true
                    break
                }

                # Add column labels to table
                if (-not $meta.comparison) {
                    $table.cols.$tableTitle.$innerPivot.$IPivotKey  = $col 
                    $table.data.$tableTitle.$innerPivot.$IPivotKey  = @{
                        $prop = @{}
                    }
                    $col += 1
                    $table.meta.columnFormats += $meta.format.$prop
                } 
                else {
                    $table.cols.$tableTitle.$innerPivot.$IPivotKey = @{
                        "baseline" = $col
                        "% Change" = $col + 1
                        "test"     = $col + 2
                    }
                    $table.meta.columnFormats += $meta.format.$prop
                    $table.meta.columnFormats += $meta.format."% change"
                    $table.meta.columnFormats += $meta.format.$prop
                    $col += 3
                    $table.data.$tableTitle.$innerPivot.$IPivotKey = @{
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

                if (-not $Metrics) {
                    $Metrics = ($data.$OPivotKey.$prop.$IPivotKey.baseline.stats.Keys | Sort)
                }

                # Add row labels and fill data in table
                foreach ($metric in $Metrics) {
                    if (-not ($table.rows.$prop.Keys -contains $metric)) {
                        $table.rows.$prop.$metric = $row
                        $row += 1
                    }

                    if (-not $meta.comparison) {
                        $table.data.$tableTitle.$innerPivot.$IPivotKey.$prop.$metric = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.$metric}
                    } 
                    else {
                        $table.data.$tableTitle.$innerPivot.$IPivotKey.baseline.$prop.$metric = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.$metric}
                        $table.data.$tableTitle.$innerPivot.$IPivotKey.test.$prop.$metric     = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.$metric}
                    
                        $percentChange = $data.$OPivotKey.$prop.$IPivotKey."% change".stats.$metric

                        $table.data.$tableTitle.$innerPivot.$IPivotKey."% change".$prop.$metric = @{"value" = "$percentChange %"}

                        $params = @{
                            "Cell"    = $table.data.$tableTitle.$innerPivot.$IPivotKey."% change".$prop.$metric
                            "TestVal" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.$metric
                            "BaseVal" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.$metric
                            "Goal"    = $meta.goal.$prop
                        }

                        # Color % change cell if necessary
                        if (@("std dev", "variance", "std err", "range") -contains $metric) {
                            $params.goal = "decrease"
                            $table.data.$tableTitle.$innerPivot.$IPivotKey."% change".$prop.$metric = Select-Color @params
                        } 
                        elseif ( -not (@("sum", "count", "kurtosis", "skewness") -contains $metric)) {
                            $table.data.$tableTitle.$innerPivot.$IPivotKey."% change".$prop.$metric = Select-Color @params
                        }
                    }
                }
            }

            if ($noStats) {
                continue
            }

            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
        }
        if (($tables.Count -gt 0) -and (-not $NoNewWorksheets)) {
            $sheetTitle = Get-WorksheetTitle -BaseName "Stats" -OuterPivot $outerPivot -OPivotKey $OPivotKey
            $tables     = @($sheetTitle) + $tables 
        }

        return $tables
    } 
    catch {
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
function Format-Quartiles {
    param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter(Mandatory=$true)] $OPivotKey, 

        [Parameter()] [String] $Tool = "",

        [Parameter()] [switch] $NoNewWorksheets
    )
    try {
        $tables = @()
        $data = $DataObj.data
        $meta = $DataObj.meta
        $innerPivot = $meta.InnerPivot
        $outerPivot = $meta.OuterPivot

        foreach ($prop in $data.$OPivotKey.Keys) { 
            $format = $meta.format.$prop
            $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey
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
                    "columnFormats" = @($format, $format, $format, $format, $format )
                    "dataWidth" = 5
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
                    "chartType"= $XLENUM.xlColumnStacked
                    "plotBy"   = $XLENUM.xlColumns
                    "xOffset"  = 1
                    "YOffset"  = 1
                    "title"    = "$cappedProp Quartiles"
                    "seriesSettings"= @{
                        1 = @{
                            "hide" = $true
                            "name" = " "
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
        
            
            # Add row labels and fill data in table
            $row = 0
            foreach ($IPivotKey in $data.$OPivotKey.$prop.Keys | Sort) {
                if (-not $meta.comparison) {
                    $table.rows.$prop.$innerPivot.$IPivotKey = $row
                    $row += 1
                    $table.data.$TableTitle.min.$prop.$innerPivot.$IPivotKey = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.min }
                    $table.data.$TableTitle.Q1.$prop.$innerPivot.$IPivotKey  = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[25] - $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.min }
                    $table.data.$TableTitle.Q2.$prop.$innerPivot.$IPivotKey  = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[50] - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[25] } 
                    $table.data.$TableTitle.Q3.$prop.$innerPivot.$IPivotKey  = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[75] - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[50]}
                    $table.data.$TableTitle.Q4.$prop.$innerPivot.$IPivotKey  = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.max - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[75] }
                } 
                else {
                    $table.rows.$prop.$innerPivot.$IPivotKey = @{
                        "baseline" = $row
                        "test"     = $row + 1
                    }
                    $row += 2
                    $table.data.$TableTitle.min.$prop.$innerPivot.$IPivotKey = @{
                        "baseline" = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.min }
                        "test"     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.min}
                    }
                    $table.data.$TableTitle.Q1.$prop.$innerPivot.$IPivotKey = @{
                        "baseline" = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[25] - $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.min }
                        "test"     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.percentiles[25] - $data.$OPivotKey.$prop.$IPivotKey.test.stats.min }
                    }
                    $table.data.$TableTitle.Q2.$prop.$innerPivot.$IPivotKey = @{
                        "baseline" = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[50] - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[25] } 
                        "test"     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.percentiles[50] - $data.$OPivotKey.$prop.$IPivotKey.test.percentiles[25] } 
                    }
                    $table.data.$TableTitle.Q3.$prop.$innerPivot.$IPivotKey = @{
                        "baseline" = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[75] - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[50] } 
                        "test"     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.percentiles[75] - $data.$OPivotKey.$prop.$IPivotKey.test.percentiles[50] }
                    }
                    $table.data.$TableTitle.Q4.$prop.$innerPivot.$IPivotKey = @{
                        "baseline" = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.max - $data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles[75] }
                        "test"     = @{ "value" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.max - $data.$OPivotKey.$prop.$IPivotKey.test.percentiles[75] }
                    }
                }
            }

            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
        }

        if (($tables.Count -gt 0) -and (-not $NoNewWorksheets)) {
            $sheetTitle = Get-WorksheetTitle -BaseName "Quartiles" -OuterPivot $outerPivot -OPivotKey $OPivotKey
            $tables = @($sheetTitle) + $tables
        }

        return $tables
    } 
    catch {
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
function Format-MinMaxChart {
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter(Mandatory=$true)] $OPivotKey, 

        [Parameter()] [String] $Tool = "",

        [Parameter()] [switch] $NoNewWorksheets
    )
    
    try {
        $tables     = @()
        $data       = $DataObj.data
        $meta       = $DataObj.meta
        $innerPivot = $meta.InnerPivot
        $outerPivot = $meta.OuterPivot

        foreach ($prop in $data.$OPivotKey.keys) {
            $cappedProp = (Get-Culture).TextInfo.ToTitleCase($prop) 
            $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey
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
                }
                "data" = @{
                    $tableTitle = @{
                        $innerPivot = @{}
                    }
                }
                "chartSettings" = @{
                    "chartType"    = $XLENUM.xlLineMarkers
                    "plotBy"       = $XLENUM.xlRows
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
                            "title" = $meta.units.$prop
                        }
                    }
                }
            }
            if ($meta.comparison) {
                $table.chartSettings.seriesSettings = @{
                    1 = @{
                        "color"       = $BLUES[2]
                        "markerColor" = $BLUES[2]
                        "markerStyle" = $XLENUM.xlMarkerStyleCircle
                        "lineWeight"  = 3
                        "markerSize"  = 5
                    }
                    2 = @{
                        "color"       = $ORANGES[2]
                        "markerColor" = $ORANGES[2]
                        "markerStyle" = $XLENUM.xlMarkerStyleCircle
                        "lineWeight"  = 3
                        "markerSize"  = 5
                    }
                    3 = @{
                        "color"       = $BLUES[1]
                        "markerColor" = $BLUES[1]
                        "markerStyle" = $XLENUM.xlMarkerStyleCircle
                        "lineWeight"  = 3
                        "markerSize"  = 5
                    }
                    4 = @{
                        "color"       = $ORANGES[1]
                        "markerColor" = $ORANGES[1]
                        "markerStyle" = $XLENUM.xlMarkerStyleCircle
                        "lineWeight"  = 3
                        "markerSize"  = 5
                    }
                    5 = @{
                        "color"       = $BLUES[0]
                        "markerColor" = $BLUES[0]
                        "markerStyle" = $XLENUM.xlMarkerStyleCircle
                        "lineWeight"  = 3
                        "markerSize"  = 5
                    }
                    6 = @{
                        "color"       = $ORANGES[0]
                        "markerColor" = $ORANGES[0]
                        "markerStyle" = $XLENUM.xlMarkerStyleCircle
                        "lineWeight"  = 3
                        "markerSize"  = 5
                    }
                }
            } 
            else {
                $table.chartSettings.seriesSettings = @{
                    1 = @{
                        "color"       = $BLUES[2]
                        "markerColor" = $BLUES[2]
                        "markerStyle" = $XLENUM.xlMarkerStyleCircle
                        "lineWeight"  = 3
                        "markerSize"  = 5
                    }
                    2 = @{
                        "color"       = $BLUES[1]
                        "markerColor" = $BLUES[1]
                        "markerStyle" = $XLENUM.xlMarkerStyleCircle
                        "lineWeight"  = 3
                        "markerSize"  = 5
                    }
                    3 = @{
                        "color"       = $BLUES[0]
                        "markerColor" = $BLUES[0]
                        "markerStyle" = $XLENUM.xlMarkerStyleCircle
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
            foreach ($IPivotKey in $data.$OPivotKey.$prop.Keys | Sort) {
                # Add column labels to table
                $table.cols.$tableTitle.$innerPivot.$IPivotKey = $col
                $table.data.$tableTitle.$innerPivot.$IPivotKey = @{
                    $prop = @{}
                }
                $table.meta.columnFormats += $meta.format.$prop
                $col += 1
            
                # Add row labels and fill data in table
                foreach ($metric in @("min", "mean", "max")) {
                    if (-not ($table.rows.$prop.Keys -contains $metric)) { 
                        if (-not $meta.comparison) {
                            $table.rows.$prop.$metric = $row
                            $row += 1
                        } 
                        else {
                            $table.rows.$prop.$metric = @{
                                "baseline" = $row
                                "test"     = $row + 1
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
                        $table.data.$tableTitle.$innerPivot.$IPivotKey.$prop.$metric.baseline = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.stats.$metric}
                        $table.data.$tableTitle.$innerPivot.$IPivotKey.$prop.$metric.test     = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.test.stats.$metric}
                    }
                }

            }
            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
        }

        if (($tables.Count -gt 0) -and (-not $NoNewWorksheets)) {
            $sheetTitle = Get-WorksheetTitle -BaseName "MinMeanMax" -OuterPivot $outerPivot -OPivotKey $OPivotKey
            $tables = @($sheetTitle) + $tables
        }

        return $tables
    } 
    catch {
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
function Format-Percentiles {
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter(Mandatory=$true)] $OPivotKey,

        [Parameter()] [String] $Tool = "",

        [Parameter()] [switch] $NoNewWorksheets
    )
    try {
        $tables     = @()
        $data       = $DataObj.data
        $meta       = $DataObj.meta
        $innerPivot = $meta.InnerPivot
        $outerPivot = $meta.OuterPivot

        
        foreach ($prop in $data.$OPivotKey.Keys) {
            foreach ($IPivotKey in $data.$OPivotKey.$prop.Keys | Sort) {
                if ($data.$OPivotKey.$prop.$IPivotKey.baseline.percentiles) {
                    $metricName = "percentiles"
                } 
                elseif ($data.$OPivotKey.$prop.$IPivotKey.baseline.percentilesHist) {
                    $metricName = "percentilesHist"
                } 
                else {
                    continue
                }

                if ($innerPivot) {
                    if ($metricName -eq "percentilesHist") {
                        $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Appx. Percentiles - $IPivotKey $innerPivot")
                    } 
                    else {
                        $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Percentiles - $IPivotKey $innerPivot")
                    }
                        
                    $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey
                } 
                else {
                    if ($metricName -eq "percentilesHist") {
                        $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Appx. Percentiles")
                    } 
                    else {
                        $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Percentiles")
                    }
                        
                    $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey
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
                        "chartType" = $XLENUM.xlXYScatterLinesNoMarkers
                        "seriesSettings" = @{
                            1 = @{ 
                                "color"      = $BLUES[1]
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

                if ($prop -eq "throughput") {
                    $max = Get-ThroughputMax -Data $data -OPivotKey $OPivotKey -IPivotKey $IPivotKey -Meta $meta
                    if ($max -ne -1) {
                        $table.chartSettings.axisSettings[2].max = $max
                    }
                }
                    
                $table.chartSettings.axisSettings[2].logarithmic = Set-Logarithmic -Data $data -OPivotKey $OPivotKey -Prop $prop -IPivotKey $IPivotKey -Meta $meta 

                if ($meta.comparison) {
                    $table.cols.$tableTitle.$prop = @{
                        "baseline" = 0
                        "% change" = 1
                        "test"     = 2
                    }
                    $table.data.$tableTitle.$prop = @{
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
                        "color"      = $ORANGES[1]
                        "lineWeight" = 3
                    }
                    $table.meta.columnFormats = @($meta.format.$prop, $meta.format."% change", $meta.format.$prop)
                }
                $row = 0

                $keys = @()
                if ($data.$OPivotKey.$prop.$IPivotKey.baseline.$metricName.Keys.Count -gt 0) {
                    $keys = $data.$OPivotKey.$prop.$IPivotKey.baseline.$metricName.Keys
                } 
                else {
                    $keys = $data.$OPivotKey.$prop.$IPivotKey.test.$metricName.Keys
                }

                # Add row labels and fill data in table
                foreach ($percentile in $keys | Sort) {
                    $table.rows.percentiles.$percentile = $row
                    if ($meta.comparison) {
                        $percentage = $data.$OPivotKey.$prop.$IPivotKey."% change".$metricName[$percentile]
                        $percentage = "$percentage %"

                        $table.data.$tableTitle.$prop.baseline.percentiles[$percentile]   = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.$metricName.$percentile}
                        $table.data.$tableTitle.$prop."% change".percentiles[$percentile] = @{"value" = $percentage}
                        $table.data.$tableTitle.$prop.test.percentiles[$percentile]       = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.test.$metricName.$percentile}
                        $params = @{
                            "Cell"    = $table.data.$tableTitle.$prop."% change".percentiles[$percentile]
                            "TestVal" = $data.$OPivotKey.$prop.$IPivotKey.test.$MetricName[$percentile]
                            "BaseVal" = $data.$OPivotKey.$prop.$IPivotKey.baseline.$MetricName[$percentile]
                            "Goal"    = $meta.goal.$prop
                        }
                        $table.data.$tableTitle.$prop."% change".percentiles[$percentile] = Select-Color @params
                    } 
                    else {
                        $table.data.$tableTitle.$prop.percentiles[$percentile] = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.$metricName.$percentile}
                    }
                    $row += 1
                
                }
                $table.meta.dataWidth     = Get-TreeWidth $table.cols
                $table.meta.colLabelDepth = Get-TreeDepth $table.cols
                $table.meta.dataHeight    = Get-TreeWidth $table.rows
                $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
                $tables = $tables + $table
            }
        }

        if (($tables.Count -gt 0) -and (-not $NoNewWorksheets)) {
            $sheetTitle = Get-WorksheetTitle -BaseName "Percentiles" -OuterPivot $outerPivot -OPivotKey $OPivotKey
            $tables     = @($sheetTitle) + $tables 
        }

        return $tables  
    } 
    catch {
        Write-Warning "Error at Format-Percentiles"
        Write-Error $_.Exception.Message
    }
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

function Get-ThroughputMax ($Data, $OPivotKey, $IPivotKey, $Meta) {
    $max = -1
    foreach ($val in $THROUGHPUTS) {
        if ($val -gt $Data.$OPivotKey.throughput.$IPivotKey.baseline.stats.max) {
            if ($Meta.comparison) {
                if ($val -gt $data.$OPivotKey.throughput.$IPivotKey.test.stats.max) {
                    $max = $val
                    break
                }
            } 
            else {
                $max = $val 
                break
            }
        }
    }
    return $max
}

function Format-Histogram {
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter(Mandatory=$true)] $OPivotKey,

        [Parameter()] [String] $Tool = "",

        [Parameter()] [switch] $NoNewWorksheets
    )
    try {
        $legend = @{
            "meta" = @{
                "colLabelDepth" = 2
                "rowLabelDepth" = 2
                "dataWidth"     = 2
                "dataHeight"    = 3 
            }
            "rows" = @{
                " "   = 0
                "  "  = 1
                "   " = 2
            }
            "cols" = @{
                "legend" = @{
                    " "  = 0
                    "  " = 1
                }
            }
            "data" = @{
                "legend" = @{
                    " " = @{
                        " " = @{
                            "value" = "For side by side comparison,`n we compare each buckets share of its `ndataset's total latency samples.  
"
                        }
                        "  " = @{
                            "value" = "buckets whose share of total `nsamples increased are colored: "
                        }
                        "   " = @{
                            "value" = "buckets whose share of total `nsamples decreased are colored:"
                        }
                    }
                    "  " = @{
                        "  " = @{
                            "value"     = "increase"
                            "fontColor" = $GREEN
                            "cellColor" = $LIGHTGREEN
                        }
                        "   " = @{
                            "value"     = "decrease"
                            "fontColor" = $RED
                            "cellColor" = $LIGHTRED
                        }
                    } 
                }
            }
        }

        $tables     = @()
        $data       = $DataObj.data
        $meta       = $DataObj.meta
        $innerPivot = $meta.InnerPivot
        $outerPivot = $meta.OuterPivot 

        foreach ($prop in $data.$OPivotKey.Keys) {
            foreach ($IPivotKey in $data.$OPivotKey.$prop.Keys | Sort) {
                if (-Not $data.$OPivotKey.$prop.$IPivotKey.baseline.Histogram) {
                    continue
                }

                $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey

                if ($IPivotKEy) {
                    $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Histogram - $IPivotKey $innerPivot") 
                } else {
                    $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Histogram") 
                }

                $units = $meta.units.$prop
                $table = @{
                    "rows" = @{
                        "histogram buckets" = @{}
                    }
                    "cols" = @{
                        $tableTitle = @{
                            $prop = 0
                        }
                    }
                    "meta" = @{
                        "rightAlign" = [Array] @(2)
                    }
                    "data" = @{
                        $tableTitle = @{
                            $prop = @{
                                "histogram buckets" = @{}
                            }
                        }
                    }
                    "chartSettings"= @{
                        "title"   = $chartTitle
                        "yOffset" = 1
                        "xOffset" = 1
                        "seriesSettings" = @{
                            1 = @{ 
                                "color"      = $BLUES[1]
                                "lineWeight" = 3
                                "name" = "Sample Count"
                            }
                        }
                        "axisSettings" = @{
                            1 = @{
                                "title" = "$prop ($units)"
                                "tickLabelSpacing" = 5
                            }
                            2 = @{
                                "title" = "Count"
                            }
                        }
                    }
                }
                
                if ($meta.comparison) {
                    $table.cols.$tableTitle.$prop = @{
                        "baseline" = 0
                        "% change" = 1
                        "test"     = 2
                    }
                    $table.data.$tableTitle.$prop = @{
                        "baseline" = @{
                            "histogram buckets" = @{}
                        }
                        "% change" = @{
                            "histogram buckets" = @{}
                        }
                        "test" = @{
                            "histogram buckets" = @{}
                        }
                    }
                    $table.chartSettings.seriesSettings[2] = @{
                        "delete" = $true
                    }
                    $table.chartSettings.seriesSettings[1].name = "Baseline Sample Count"
                     
                    $table.chartSettings.seriesSettings[3] = @{
                        "color"      = $ORANGES[1]
                        "name"       = "Test Sample Count"
                        "lineWeight" = 3
                    }
                    $table.meta.columnFormats = @($null, $meta.format."% change", $null)
                }
                $row = 0

                $keys = @()
                if ($data.$OPivotKey.$prop.$IPivotKey.baseline.histogram.Keys.Count -gt 0) {
                    $keys = $data.$OPivotKey.$prop.$IPivotKey.baseline.histogram.Keys
                } else {
                    $keys = $data.$OPivotKey.$prop.$IPivotKey.test.histogram.Keys
                }

                # Add row labels and fill data in table
                foreach ($bucket in ($keys | Sort)) {
                    $table.rows."histogram buckets".$bucket = $row
                    if ($meta.comparison) {
                        $percentage = $data.$OPivotKey.$prop.$IPivotKey."% change".histogram[$bucket]
                        $percentage = "$percentage %"

                        $table.data.$tableTitle.$prop.baseline."histogram buckets"[$bucket]   = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.histogram.$bucket}
                        $table.data.$tableTitle.$prop."% change"."histogram buckets"[$bucket] = @{"value" = $percentage}
                        $table.data.$tableTitle.$prop.test."histogram buckets"[$bucket]       = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.test.histogram.$bucket}
                        $params = @{
                            "Cell"    = $table.data.$tableTitle.$prop."% change"."histogram buckets"[$bucket]
                            "TestVal" = $data.$OPivotKey.$prop.$IPivotKey.test.histogram[$bucket]
                            "BaseVal" = $data.$OPivotKey.$prop.$IPivotKey.baseline.histogram[$bucket]
                            "Goal"    = "increase"
                        }
                        $table.data.$tableTitle.$prop."% change"."histogram buckets"[$bucket] = Select-Color @params
                    } 
                    else {
                        $table.data.$tableTitle.$prop."histogram buckets"[$bucket] = @{"value" = $data.$OPivotKey.$prop.$IPivotKey.baseline.histogram.$bucket}
                    }
                    $row += 1
                
                }
                $table.meta.dataWidth     = Get-TreeWidth $table.cols
                $table.meta.colLabelDepth = Get-TreeDepth $table.cols
                $table.meta.dataHeight    = Get-TreeWidth $table.rows
                $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
                $tables = $tables + $table
            }
        }
        if ($table.Count -gt 0) {
            if ($meta.comparison) {
                $tables = @($legend) + $tables
            }
            $worksheetTitle = Get-WorksheetTitle -BaseName "Histogram" -OuterPivot $outerPivot -OPivotKey $OPivotKey
            $tables = @($worksheetTitle) + $tables
        }
        return $tables  
    } 
    catch {
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
function Format-Distribution {
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter()] [string] $OPivotKey,

        [Parameter()] [String] $Tool = "",

        [Parameter()] [String] $Prop = "latency",

        [Parameter()] [String] $NumSamples = 10000,
    
        [Parameter()] [Int] $SubSampleRate = 50,

        [Parameter()] [switch] $NoNewWorksheets
        
    )
    try {
        $meta  = $DataObj.meta 
        $modes = @("baseline")
        if ($meta.comparison) {
            $modes += "test"
        }
        $tables     = @()
        $innerPivot = $meta.InnerPivot
        $outerPivot = $meta.OuterPivot

        foreach ($IPivotKey in $DataObj.data.$OPivotKey.$Prop.Keys) {
            foreach ($mode in $modes) { 
                if (-Not $DataObj.data.$OPivotKey.$Prop.$IPivotKey.$mode.stats) {
                    continue
                }
                $tableTitle = Get-TableTitle -Tool $Tool -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey
                $data       = $dataObj.rawData.$mode
                $table = @{
                    "meta" = @{}
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
                            "latency" = @{
                                "Data Point" = @{}
                            }
                        }
                    }
                    "chartSettings" = @{
                        "chartType" = $XLENUM.xlXYScatter
                        "yOffset"   = 2
                        "xOffset"   = 2
                        "title"     = "Temporal Latency Distribution"
                        "axisSettings" = @{
                            1 = @{
                                "title"          = "Time Series"
                                "max"            = $NumSamples
                                "minorGridlines" = $true
                                "majorGridlines" = $true
                            }
                            2 = @{
                                "title"       = "us"
                                "logarithmic" = $true
                                "min"         = 10
                            }
                        }
                    }
                }

                if ($mode -eq "baseline") {
                    $table.chartSettings.seriesSettings = @{
                        1 = @{
                                "markerStyle"           = $XLENUM.xlMarkerStyleCircle
                                "markerBackgroundColor" = $BLUES[2]
                                "markerForegroundColor" = $BLUES[1]
                                "name"                  = "$prop Sample" 
                            }
                    }
                } else {
                    $table.chartSettings.seriesSettings = @{
                        1 = @{
                                "markerStyle"           = $XLENUM.xlMarkerStyleCircle
                                "markerBackgroundColor" = $ORANGES[2]
                                "markerForegroundColor" = $ORANGES[1]
                                "name"                  = "$prop Sample"
                            }
                    }
                }

                # Add row labels and fill data in table
                $i   = 0
                $row = 0
                if ($SubSampleRate -gt 0) {
                    $NumSegments = $NumSamples / $SubSampleRate
                    while ($i -lt $NumSegments) {
                        [Array]$segmentData = @()
                        foreach ($entry in $data) {
                            if ($entry.$prop.GetType().Name -ne "Object[]") {
                                continue
                            }
                            if (((-not $innerPivot) -or ($entry.$innerPivot -eq $IPivotKey)) -and ((-not $outerPivot) -or ($entry.$outerPivot -eq $OPivotKey))) {
                                $segmentData += $entry[$Prop][($i * $SubSampleRate) .. ((($i + 1) * $SubSampleRate) - 1)]
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
                        } 
                        else {
                            foreach ($sample in $segmentData) {
                                $table.rows."Data Point".$row = $row
                                $table.data.$tableTitle."Time Segment"."Data Point".$row = @{"value" = $time}
                                $table.data.$tableTitle.$Prop."Data Point".$row          = @{"value" = $sample}
                                $row++
                            }
                        }
                        $i++
                    }
                } else {
                    while ($i -lt $NumSamples) {
                        [Array]$segmentData = @()
                        foreach ($entry in $data) {
                            if ($entry.$prop.GetType().Name -ne "Object[]") {
                                continue
                            }
                            if (((-not $innerPivot) -or ($entry.$innerPivot -eq $IPivotKey)) -and ((-not $outerPivot) -or ($entry.$outerPivot -eq $OPivotKey))) {
                                $segmentData += $entry[$Prop][$i]
                            }
                        }

                        foreach ($sample in $segmentData) {
                            $table.rows."Data Point".$row = $row
                            $table.data.$tableTitle."Time Segment"."Data Point".$row = @{"value" = $i}
                            $table.data.$tableTitle.$Prop."Data Point".$row          = @{"value" = $sample}
                            $row++
                        }
                        $i++
                    }
                }
                $table.meta.dataWidth     = Get-TreeWidth $table.cols
                $table.meta.colLabelDepth = Get-TreeDepth $table.cols
                $table.meta.dataHeight    = Get-TreeWidth $table.rows
                $table.meta.rowLabelDepth = Get-TreeDepth $table.rows

                if (-not $NoNewWorksheets) {
                    if ($modes.Count -gt 1) {
                        if ($mode -eq "baseline") {
                            $worksheetName = Get-WorksheetTitle -BaseName "Base Distr." -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey
                        } 
                        else {
                            $worksheetName = Get-WorksheetTitle -BaseName "Test Distr." -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey
                        } 
                    } 
                    else {
                        $worksheetName = Get-WorksheetTitle -BaseName "Distr." -OuterPivot $outerPivot -OPivotKey $OPivotKey -InnerPivot $innerPivot -IPivotKey $IPivotKey
                    }
                    
                    $tables += $worksheetName
                }

                $tables += $table
            }
        }
        return $tables
    } 
    catch {
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
function Select-Color ($Cell, $TestVal, $BaseVal, $Goal) {
    if ( $Goal -eq "increase") {
        if ($TestVal -ge $BaseVal) {
            $Cell["fontColor"] = $GREEN
            $Cell["cellColor"] = $LIGHTGREEN
        } 
        else {
            $Cell["fontColor"] = $RED
            $Cell["cellColor"] = $LIGHTRED
        }
    } 
    else {
        if ($TestVal -le $BaseVal) {
            $Cell["fontColor"] = $GREEN
            $Cell["cellColor"] = $LIGHTGREEN
        } 
        else {
            $Cell["fontColor"] = $RED
            $Cell["cellColor"] = $LIGHTRED
        }
    }
    return $cell
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
function Sort-ByProp {
    param(
        [Parameter()]
        [PSObject] $Data,

        [Parameter()]
        [string] $Prop
    )

    if ($Data.length -eq 1) {
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
    
    while ($idx1 -lt $arr1.length -and $idx2 -lt $arr2.length) {
        if ($arr1[$idx1].$prop -le $arr2[$idx2].$prop) {
            $sorted  = $sorted + $arr1[$idx1]
            $idx1   += 1
        } 
        else {
            $sorted  = $sorted + $arr2[$idx2]
            $idx2   += 1
        }
    }

    while ($idx1 -lt $arr1.length) {
        $sorted  = $sorted + $arr1[$idx1]
        $idx1   += 1
    }

    while ($idx2 -lt $arr2.length) {
        $sorted  = $sorted + $arr2[$idx2]
        $idx2   += 1
    }
    return $sorted
}