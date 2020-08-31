$LIGHTGREEN = 10416289
$GREEN = 1268766
$LIGHTRED = 10396159
$RED = 2108032
$BLUES = @(10249511, 14058822, 16758932)
$ORANGES = @(294092, 1681916, 6014716)

$EPS = 0.0001

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

        [Parameter()] [String] $TableTitle = ""
    )

    try {
        $tables = @() 
        $meta = $DataObj.meta
        $sortProp = $meta.sortProp 
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

        if ($meta.comparison) {
            $tables += $legend
        }

        # Fill single array with all data and sort, label data as baseline/test if necessary
        [Array] $data = @()
        $baseData = $DataObj.rawData.baseline
        foreach ($entry in $baseData) {
            if ($meta.comparison) {
                $entry.baseline = $true
            }
            $data += $entry
        }

        if ($meta.comparison) {
            $testData = $DataObj.rawData.test
            foreach ($entry in $testData) {
                $data += $entry
            }
        }

        $data = Sort-ByProp -Data $data -Prop $sortProp

        foreach ($prop in $dataObj.data.Keys) {
            $table = @{
                "rows" = @{
                    $prop = @{}
                }
                "cols" = @{
                    $TableTitle = @{}
                }
                "meta" = @{
                    "columnFormats" = @()
                    "leftAlign"     = [Array] @(2)
                }
                "data"  = @{
                    $TableTitle = @{}
                }
            }
            $col = 0
            $row = 0

            foreach ($entry in $data) {
                $sortKey = $entry.$sortProp

                # Add column labels to table
                if (-not ($table.cols.$TableTitle.Keys -contains $sortKey)) {
                    if ($meta.comparison) {
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
                    else {
                        $table.meta.columnFormats        += $meta.format.$prop
                        $table.cols.$TableTitle.$sortKey  = $col
                        $table.data.$TableTitle.$sortKey  = @{
                            $prop = @{}
                        }
                        $col += 1
                    }
                }

                # Add row labels and fill data in table
                $filename = $entry.fileName.Split('\')[-2] + "\" + $entry.fileName.Split('\')[-1] 
                $table.rows.$prop.$filename = $row
                
                $row += 1
                if ($meta.comparison) {
                    if ($entry.baseline) {
                        $table.data.$TableTitle.$sortKey.baseline.$prop.$filename = @{
                            "value" = $entry.$prop
                        }
                    }
                    else {
                        $table.data.$TableTitle.$sortKey.test.$prop.$filename = @{
                            "value" = $entry.$prop
                        }
                        $params = @{
                            "Cell"    = $table.data.$TableTitle.$sortKey.test.$prop.$filename
                            "TestVal" = $entry.$prop
                            "BaseVal" = $DataObj.data.$prop.$sortKey.baseline.stats.mean
                            "Goal"    = $meta.goal.$prop
                        }
                        
                        $table.data.$TableTitle.$sortKey.test.$prop.$filename = Select-Color @params
                    }
                } 
                else {
                    $table.data.$TableTitle.$sortKey.$prop.$filename = @{
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

        [Parameter()] [String] $TableTitle = "",

        [Parameter()] [Array] $Metrics=$null
    )
    
    try {
        $tables = @()
        $data = $DataObj.data
        $meta = $DataObj.meta
        foreach ($prop in $data.keys) { 
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
            $noStats = $false
            foreach ($sortKey in $data.$prop.Keys | Sort) { 
                if (-not $data.$prop.$sortKey.baseline.stats) {
                    $noStats = $true
                    break
                }
                # Add column labels to table
                if (-not $meta.comparison) {
                    $table.cols.$TableTitle.$sortKey  = $col
                    $table.meta.columnFormats        += $meta.format.$prop 
                    $table.data.$TableTitle.$sortKey  = @{
                        $prop = @{}
                    }
                    $col += 1
                } 
                else {
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

                if (-not $Metrics) {
                    $Metrics = ($data.$prop.$sortKey.baseline.stats.Keys | Sort)
                }

                # Add row labels and fill data in table
                foreach ($metric in $Metrics) {
                    if (-not ($table.rows.$prop.Keys -contains $metric)) {
                        $table.rows.$prop.$metric = $row
                        $row += 1
                    }

                    if (-not $meta.comparison) {
                        $table.data.$TableTitle.$sortKey.$prop.$metric = @{"value" = $data.$prop.$sortKey.baseline.stats.$metric}
                    } 
                    else {
                        $table.data.$TableTitle.$sortKey.baseline.$prop.$metric = @{"value" = $data.$prop.$sortKey.baseline.stats.$metric}
                        $table.data.$TableTitle.$sortKey.test.$prop.$metric     = @{"value" = $data.$prop.$sortKey.test.stats.$metric}
                    
                        $percentChange = $data.$prop.$sortKey."% change".stats.$metric

                        $table.data.$TableTitle.$sortKey."% change".$prop.$metric = @{"value" = "$percentChange %"}

                        $params = @{
                            "Cell"    = $table.data.$TableTitle.$sortKey."% change".$prop.$metric
                            "TestVal" = $data.$prop.$sortKey.test.stats.$metric
                            "BaseVal" = $data.$prop.$sortKey.baseline.stats.$metric
                            "Goal"    = $meta.goal.$prop
                        }

                        # Color % change cell if necessary
                        if (@("std dev", "variance", "std err", "range") -contains $metric) {
                            $params.goal = "decrease"
                            $table.data.$TableTitle.$sortKey."% change".$prop.$metric = Select-Color @params
                        } 
                        elseif ( -not (@("sum", "count", "kurtosis", "skewness") -contains $metric)) {
                            $table.data.$TableTitle.$sortKey."% change".$prop.$metric = Select-Color @params
                        }
                    }
                }
            }
            if ($noStats) {
                $noStats = $false
                continue
            }

            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
        }
        if ($tables.Count -gt 0) {
            $tables = @("Stats") + $tables 
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

        [Parameter()] [String] $TableTitle = ""
    )
    try {
        $tables = @()
        $data = $DataObj.data
        $meta = $DataObj.meta
        $sortProp = $meta.sortProp

        foreach ($prop in $data.Keys) { 
            $format = $meta.format.$prop
            $cappedProp = (Get-Culture).TextInfo.ToTitleCase($prop)
            $table = @{
                "rows" = @{
                    $prop = @{
                        $sortProp = @{}
                    }
                }
                "cols" = @{
                    $TableTitle = @{
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
            foreach ($sortKey in $data.$prop.Keys | Sort) {
                if (-not $meta.comparison) {
                    $table.rows.$prop.$sortProp.$sortKey = $row
                    $row += 1
                    $table.data.$TableTitle.min.$prop.$sortProp.$sortKey = @{ "value" = $data.$prop.$sortKey.baseline.stats.min }
                    $table.data.$TableTitle.Q1.$prop.$sortProp.$sortKey  = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[25] - $data.$prop.$sortKey.baseline.stats.min }
                    $table.data.$TableTitle.Q2.$prop.$sortProp.$sortKey  = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[50] - $data.$prop.$sortKey.baseline.percentiles[25] } 
                    $table.data.$TableTitle.Q3.$prop.$sortProp.$sortKey  = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[75] - $data.$prop.$sortKey.baseline.percentiles[50]}
                    $table.data.$TableTitle.Q4.$prop.$sortProp.$sortKey  = @{ "value" = $data.$prop.$sortKey.baseline.stats.max - $data.$prop.$sortKey.baseline.percentiles[75] }
                } 
                else {
                    $table.rows.$prop.$sortProp.$sortKey = @{
                        "baseline" = $row
                        "test"     = $row + 1
                    }
                    $row += 2
                    $table.data.$TableTitle.min.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.stats.min }
                        "test"     = @{ "value" = $data.$prop.$sortKey.test.stats.min}
                    }
                    $table.data.$TableTitle.Q1.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[25] - $data.$prop.$sortKey.baseline.stats.min }
                        "test"     = @{ "value" = $data.$prop.$sortKey.test.percentiles[25] - $data.$prop.$sortKey.test.stats.min }
                    }
                    $table.data.$TableTitle.Q2.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[50] - $data.$prop.$sortKey.baseline.percentiles[25] } 
                        "test"     = @{ "value" = $data.$prop.$sortKey.test.percentiles[50] - $data.$prop.$sortKey.test.percentiles[25] } 
                    }
                    $table.data.$TableTitle.Q3.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.percentiles[75] - $data.$prop.$sortKey.baseline.percentiles[50] } 
                        "test"     = @{ "value" = $data.$prop.$sortKey.test.percentiles[75] - $data.$prop.$sortKey.test.percentiles[50] }
                    }
                    $table.data.$TableTitle.Q4.$prop.$sortProp.$sortKey = @{
                        "baseline" = @{ "value" = $data.$prop.$sortKey.baseline.stats.max - $data.$prop.$sortKey.baseline.percentiles[75] }
                        "test"     = @{ "value" = $data.$prop.$sortKey.test.stats.max - $data.$prop.$sortKey.test.percentiles[75] }
                    }
                }

            }
            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
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

        [Parameter()] [String] $TableTitle = ""
    )
    
    try {
        $tables   = @()
        $data     = $DataObj.data
        $meta     = $DataObj.meta
        $sortProp = $meta.sortProp

        foreach ($prop in $data.keys) {
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
                    "chartType"    = $XLENUM.xlLineMarkers
                    "plotBy"       = $XLENUM.xlRows
                    "title"        = $cappedProp
                    "xOffset"      = 1
                    "yOffset"      = 1
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
            $col = 0
            $row = 0
            foreach ($sortKey in $data.$prop.Keys | Sort) {
                # Add column labels to table
                $table.cols.$TableTitle.$sortProp.$sortKey = $col
                $table.data.$TableTitle.$sortProp.$sortKey = @{
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
                    if (-not ($table.data.$TableTitle.$sortProp.$sortKey.$prop.Keys -contains $metric)) {
                        $table.data.$TableTitle.$sortProp.$sortKey.$prop.$metric = @{}
                    }

                    if (-not $meta.comparison) {
                        $table.data.$TableTitle.$sortProp.$sortKey.$prop.$metric = @{"value" = $data.$prop.$sortKey.baseline.stats.$metric}
                    } 
                    else {
                        $table.data.$TableTitle.$sortProp.$sortKey.$prop.$metric.baseline = @{"value" = $data.$prop.$sortKey.baseline.stats.$metric}
                        $table.data.$TableTitle.$sortProp.$sortKey.$prop.$metric.test     = @{"value" = $data.$prop.$sortKey.test.stats.$metric}
                    }
                }

            }
            $table.meta.dataWidth     = Get-TreeWidth $table.cols
            $table.meta.colLabelDepth = Get-TreeDepth $table.cols
            $table.meta.dataHeight    = Get-TreeWidth $table.rows
            $table.meta.rowLabelDepth = Get-TreeDepth $table.rows
            $tables = $tables + $table
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

        [Parameter()] [String] $TableTitle = ""
    )
    try {
        $tables    = @()
        $data      = $DataObj.data
        $meta      = $DataObj.meta
        $sortProp  = $meta.sortProp
        $baseTitle = $TableTitle
        foreach ($MetricName in @("percentiles", "percentilesHist" )){
            foreach ($prop in $data.Keys) {
                foreach ($sortKey in $data.$prop.Keys | Sort) {
                    if (-not $data.$prop.$sortKey.baseline.$MetricName) {
                        continue
                    }
                    if ($sortProp) {
                        $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Percentiles - $sortKey $sortProp")
                        $TableTitle = "$baseTitle - $sortKey $sortProp"
                    } else {
                        $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Percentiles")
                        $TableTitle = "$baseTitle"
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
                            "rightAlign"    = [Array] @(2)
                        }
                        "data" = @{
                            $TableTitle = @{
                                $prop = @{
                                    "percentiles" = @{}
                                }
                            }
                        }
                        "chartSettings"= @{
                            "title"     = $chartTitle
                            "yOffset"   = 1
                            "xOffset"   = 1
                            "chartType" = $XLENUM.xlXYScatterLinesNoMarkers
                            "seriesSettings" = @{
                                1 = @{ 
                                    "color" = $BLUES[1]
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
                        $max = -1
                        foreach ($val in $THROUGHPUTS) {
                            if ($val -gt $data.$prop.$sortKey.baseline.stats.max) {
                                if ($meta.comparison) {
                                    if ($val -gt $data.$prop.$sortKey.test.stats.max) {
                                        $max = $val
                                        break
                                    }
                                } else {
                                    $max = $val 
                                    break
                                }
                            }
                        }
                        if ($max -ne -1) {
                            $table.chartSettings.axisSettings[2].max = $max
                        }
                    }
                    if ($data.$prop.$sortKey.baseline.stats) {
                        if (($data.$prop.$sortKey.baseline.stats.max / ($data.$prop.$sortKey.baseline.stats.min + $EPS)) -gt 10) {
                            $table.chartSettings.axisSettings[2].logarithmic = $true
                        }
                        if ($meta.comparison -and (($data.$prop.$sortKey.test.stats.max / ($data.$prop.$sortKey.test.stats.min + $EPS)) -gt 10)) {
                            $table.chartSettings.axisSettings[2].logarithmic = $true
                        }
                    }

                    if ($meta.comparison) {
                        $table.cols.$TableTitle.$prop = @{
                            "baseline" = 0
                            "% change" = 1
                            "test"     = 2
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
                            "color"      = $ORANGES[1]
                            "lineWeight" = 3
                        }
                        $table.meta.columnFormats = @($meta.format.$prop, $meta.format."% change", $meta.format.$prop)
                    }
                    $row = 0

                    $keys = @()
                    if ($data.$prop.$sortKey.baseline.$MetricName.Keys.Count -gt 0) {
                        $keys = $data.$prop.$sortKey.baseline.$MetricName.Keys
                    } else {
                        $keys = $data.$prop.$sortKey.test.$MetricName.Keys
                    }

                    # Add row labels and fill data in table
                    foreach ($percentile in $keys | Sort) {
                        $table.rows.percentiles.$percentile = $row
                        if ($meta.comparison) {
                            $percentage = $data.$prop.$sortKey."% change".$MetricName[$percentile]
                            $percentage = "$percentage %"

                            $table.data.$TableTitle.$prop.baseline.percentiles[$percentile]   = @{"value" = $data.$prop.$sortKey.baseline.$MetricName.$percentile}
                            $table.data.$TableTitle.$prop."% change".percentiles[$percentile] = @{"value" = $percentage}
                            $table.data.$TableTitle.$prop.test.percentiles[$percentile]       = @{"value" = $data.$prop.$sortKey.test.$MetricName.$percentile}
                            $params = @{
                                "Cell"    = $table.data.$TableTitle.$prop."% change".percentiles[$percentile]
                                "TestVal" = $data.$prop.$sortKey.test.$MetricName[$percentile]
                                "BaseVal" = $data.$prop.$sortKey.baseline.$MetricName[$percentile]
                                "Goal"    = $meta.goal.$prop
                            }
                            $table.data.$TableTitle.$prop."% change".percentiles[$percentile] = Select-Color @params
                        } 
                        else {
                            $table.data.$TableTitle.$prop.percentiles[$percentile] = @{"value" = $data.$prop.$sortKey.baseline.$MetricName.$percentile}
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
        }
        return $tables  
    } 
    catch {
        Write-Warning "Error at Format-Percentiles"
        Write-Error $_.Exception.Message
    }
}

function Format-Histogram {
    Param (
        [Parameter(Mandatory=$true)] [PSobject[]] $DataObj,

        [Parameter()] [String] $TableTitle = ""
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

        $tables    = @()
        $data      = $DataObj.data
        $meta      = $DataObj.meta
        $sortProp  = $meta.sortProp
        $baseTitle = $TableTitle
        foreach ($prop in $data.Keys) {
            foreach ($sortKey in $data.$prop.Keys | Sort) {
                if (-Not $data.$prop.$sortKey.baseline.Histogram) {
                    continue
                }

                if ($sortProp) {
                    $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Histogram - $sortKey $sortProp")
                    $TableTitle = "$baseTitle - $sortKey $sortProp"
                } else {
                    $chartTitle = (Get-Culture).TextInfo.ToTitleCase("$prop Histogram")
                    $TableTitle = "$baseTitle"
                }
                $units = $meta.units.$prop
                $table = @{
                    "rows" = @{
                        "histogram buckets" = @{}
                    }
                    "cols" = @{
                        $TableTitle = @{
                            $prop = 0
                        }
                    }
                    "meta" = @{
                        "rightAlign" = [Array] @(2)
                    }
                    "data" = @{
                        $TableTitle = @{
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
                    $table.cols.$TableTitle.$prop = @{
                        "baseline" = 0
                        "% change" = 1
                        "test"     = 2
                    }
                    $table.data.$TableTitle.$prop = @{
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
                if ($data.$prop.$sortKey.baseline.histogram.Keys.Count -gt 0) {
                    $keys = $data.$prop.$sortKey.baseline.histogram.Keys
                } else {
                    $keys = $data.$prop.$sortKey.test.histogram.Keys
                }

                # Add row labels and fill data in table
                foreach ($bucket in ($keys | Sort)) {
                    $table.rows."histogram buckets".$bucket = $row
                    if ($meta.comparison) {
                        $percentage = $data.$prop.$sortKey."% change".histogram[$bucket]
                        $percentage = "$percentage %"

                        $table.data.$TableTitle.$prop.baseline."histogram buckets"[$bucket]   = @{"value" = $data.$prop.$sortKey.baseline.histogram.$bucket}
                        $table.data.$TableTitle.$prop."% change"."histogram buckets"[$bucket] = @{"value" = $percentage}
                        $table.data.$TableTitle.$prop.test."histogram buckets"[$bucket]       = @{"value" = $data.$prop.$sortKey.test.histogram.$bucket}
                        $params = @{
                            "Cell"    = $table.data.$TableTitle.$prop."% change"."histogram buckets"[$bucket]
                            "TestVal" = $data.$prop.$sortKey.test.histogram[$bucket]
                            "BaseVal" = $data.$prop.$sortKey.baseline.histogram[$bucket]
                            "Goal"    = "increase"
                        }
                        $table.data.$TableTitle.$prop."% change"."histogram buckets"[$bucket] = Select-Color @params
                    } 
                    else {
                        $table.data.$TableTitle.$prop."histogram buckets"[$bucket] = @{"value" = $data.$prop.$sortKey.baseline.histogram.$bucket}
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
            $tables = @("Summary Histogram") + $tables
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

        [Parameter()] [String] $Title = "",

        [Parameter()] [String] $Prop = "latency",

        [Parameter()] [String] $NumSamples = 10000,
    
        [Parameter()] [Int] $SubSampleRate = 50
        
    )
    try {
        $meta  = $DataObj.meta
        $modes = @("baseline")
        if ($meta.comparison) {
            $modes += "test"
        }
        $tables = @()
        $sortProp = $DataObj.meta.sortProp
        foreach ($sortKey in $DataObj.data.$Prop.Keys) {
            foreach ($mode in $modes) { 
                if (-Not $DataObj.data.$Prop.$sortKey.$mode.stats) {
                    continue
                }
                $tables += (Get-Culture).TextInfo.ToTitleCase("$mode Raw Dist. - $sortKey")
                $TableTitle = (Get-Culture).TextInfo.ToTitleCase("$mode Dist. - $sortKey $sortProp")
                $data = $dataObj.rawData.$mode
                $table = @{
                    "meta" = @{}
                    "rows" = @{
                        "Data Point" = @{}
                    }
                    "cols" = @{
                        $TableTitle = @{
                            "Time Segment" = 0
                            $Prop          = 1
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
                $NumSegments = $NumSamples / $SubSampleRate
                while ($i -lt $NumSegments) {
                    [Array]$segmentData = @()
                    foreach ($entry in $data) {
                        if ($entry.$prop.GetType().Name -ne "Object[]") {
                            continue
                        }
                        if ($entry.$sortProp -ne $sortKey) {
                            continue
                        }
                        $segmentData += $entry[$Prop][($i * $SubSampleRate) .. ((($i + 1) * $SubSampleRate) - 1)]
                    }
                    $segmentData = $segmentData | Sort
                    $time        = $i * $subSampleRate
                    if ($segmentData.Count -ge 5) {
                        $table.rows."Data Point".$row       = $row
                        $table.rows."Data Point".($row + 1) = $row + 1
                        $table.rows."Data Point".($row + 2) = $row + 2
                        $table.rows."Data Point".($row + 3) = $row + 3
                        $table.rows."Data Point".($row + 4) = $row + 4
                        $table.data.$TableTitle."Time Segment"."Data Point".$row       = @{"value" = $time}
                        $table.data.$TableTitle."Time Segment"."Data Point".($row + 1) = @{"value" = $time}
                        $table.data.$TableTitle."Time Segment"."Data Point".($row + 2) = @{"value" = $time}
                        $table.data.$TableTitle."Time Segment"."Data Point".($row + 3) = @{"value" = $time}
                        $table.data.$TableTitle."Time Segment"."Data Point".($row + 4) = @{"value" = $time}
                        $table.data.$TableTitle.$Prop."Data Point".$row = @{"value"       = $segmentData[0]}
                        $table.data.$TableTitle.$Prop."Data Point".($row + 1) = @{"value" = $segmentData[[int]($segmentData.Count / 4)]}
                        $table.data.$TableTitle.$Prop."Data Point".($row + 2) = @{"value" = $segmentData[[int]($segmentData.Count / 2)]}
                        $table.data.$TableTitle.$Prop."Data Point".($row + 3) = @{"value" = $segmentData[[int]((3 * $segmentData.Count) / 4)]}
                        $table.data.$TableTitle.$Prop."Data Point".($row + 4) = @{"value" = $segmentData[-1]}
                        $row += 5
                    } 
                    else {
                        foreach ($sample in $segmentData) {
                            $table.rows."Data Point".$row = $row
                            $table.data.$TableTitle."Time Segment"."Data Point".$row = @{"value" = $time}
                            $table.data.$TableTitle.$Prop."Data Point".$row          = @{"value" = $sample}
                            $row++
                        }
                    }
                    $i++
                }
                $table.meta.dataWidth     = Get-TreeWidth $table.cols
                $table.meta.colLabelDepth = Get-TreeDepth $table.cols
                $table.meta.dataHeight    = Get-TreeWidth $table.rows
                $table.meta.rowLabelDepth = Get-TreeDepth $table.rows

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