$Percentiles = @(0, 1, 5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90, 95, 96, 97, 98,`
                                         99, 99.9, 99.99, 99.999, 99.9999, 99.99999, 100)
$NoPivot = " "

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
function Process-Data {
    param (
        [Parameter(Mandatory=$true)] [PSobject[]] $BaselineRawData,
        [Parameter()] [PSobject[]] $TestRawData,
        [Parameter()] [String] $InnerPivot,
        [Parameter()] [String] $OuterPivot
    )
    try {
        $processedDataObj = @{
            "meta"    = $BaselineRawData.meta
            "data"    = @{}
            "rawData" = @{
                "baseline" = $BaselineRawData.data
            }
        }

        $processedDataObj.meta.InnerPivot = $InnerPivot
        $processedDataObj.meta.OuterPivot = $OuterPivot

        if ($TestRawData) {
            $processedDataObj.meta.comparison = $true
            $processedDataObj.rawData.test    = $TestRawData.data
        }

        $meta = $processedDataObj.meta 

        $modes = @("baseline")
        foreach ($prop in ([Array]$BaselineRawData.data)[0].Keys) {
            if ($BaselineRawData.meta.noTable -contains $prop) {
                continue
            }

            # Extract property values from dataEntry objects and place values in the correct spot within the processedData object
            foreach($item in $BaselineRawData.data) {
                Place-DataEntry -DataObj $processedDataObj -Item $item -Property $prop -InnerPivot $InnerPivot -OuterPivot $OuterPivot -Mode "baseline"
            }

            if ($TestRawData) {
                $modes += "test"
                foreach ($item in $TestRawData.data) {
                    Place-DataEntry -DataObj $processedDataObj -Item $item -Property $prop -InnerPivot $InnerPivot -OuterPivot $OuterPivot -Mode "test"
                }
            }

            foreach ($oPivotKey in $processedDataObj.data.Keys) {
                foreach ($iPivotKey in $processedDataObj.data.$oPivotKey.$prop.Keys) {
                    foreach ($mode in $modes) {
                        if ($processedDataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.orderedData) {
                            Fill-Metrics -DataObj $processedDataObj -Property $prop -IPivotKey $iPivotKey -OPivotKey $oPivotKey -Mode $mode
                        }
                        if ($processedDataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.histogram) {
                            Percentiles-FromHistogram -DataObj $processedDataObj -Property $prop -IPivotKey $iPivotKey -OPivotKey $oPivotKey -Mode $mode
                        }
                    }
                    if ($TestRawData) {
                        Calculate-PercentChange -DataObj $processedDataObj -Property $prop -IPivotKey $iPivotKey -OPivotKey $oPivotKey
                    }
                }
            }
        }
        return $processedDataObj
    } 
    catch {
        Write-Warning "Error in Process-Data"
        Write-Error $_.Exception.Message
    }
}

function Percentiles-FromHistogram ($DataObj, $Property, $IPivotKey, $OPivotKey, $Mode) {
    # Calculate cumulative density function
    $cdf = @{}
    $sumSoFar = 0
    foreach ($bucket in ($DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram.Keys | Sort)) {
        $sumSoFar    += $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram.$bucket
        $cdf.$bucket  = $sumSoFar 
    }

    $buckets = ([Array] $cdf.Keys | Sort )

    foreach ($bucket in $buckets) {
        $cdf.$bucket = 100 * ($cdf.$bucket / $sumSoFar)
    }

    $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.percentilesHist = @{}
    
    $j = 0
    $i = 0

    while ($j -lt $Percentiles.Count) {
        while (($cdf.($buckets[$i]) -le $Percentiles[$j]) -and ($i -lt ($buckets.Count - 1))) {
            $i++
        }

        if ($i -eq 0) {
            $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.percentilesHist.($Percentiles[$j]) = $buckets[$i]
        } 
        else {
            $lowerVal    = $cdf.($buckets[$i - 1])
            $lowerBucket = $buckets[$i - 1]
            $upperVal    = $cdf.($buckets[$i])
            $upperBucket = $buckets[$i] 
            
            $dist = ($Percentiles[$j] - $lowerVal) / ($upperVal - $lowerVal)
            
            $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.percentilesHist.($Percentiles[$j]) = ($dist * $upperBucket) + ((1 - $dist) * $lowerBucket)
        }

        $j++ 
    }

}

function Calculate-PercentChange ($DataObj, $Property, $IPivotKey, $OPivotKey) {
    $DataObj.data.$OPivotKey.$Property.$IPivotKey."% change" = @{}
    foreach ($metricSet in @("stats", "percentiles", "percentilesHist")) {
        if (-not $DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.$metricSet) {
            continue
        }

        $DataObj.data.$OPivotKey.$Property.$IPivotKey."% change".$metricSet = @{}
        foreach ($metric in $DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.$metricSet.Keys) {
            $diff          =  $DataObj.data.$OPivotKey.$Property.$IPivotKey.test.$metricSet.$metric - $DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.$metricSet.$metric
            if ($DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.$metricSet.$metric) {
                $percentChange = 100 * ($diff / [math]::Abs( $DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.$metricSet.$metric))
                $DataObj.data.$OPivotKey.$Property.$IPivotKey."% change".$metricSet.$metric = $percentChange
            }
                  
        }
    }

    if ($DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.histogram) {
        $DataObj.data.$OPivotKey.$Property.$IPivotKey."% change".histogram = @{}
        $baseTotal = 0
        $testTotal = 0
        foreach ($bucket in $DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.histogram.Keys) {
            $baseTotal += $DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.histogram.$bucket
            $testTotal += $DataObj.data.$OPivotKey.$Property.$IPivotKey.test.histogram.$bucket
        }

        foreach ($bucket in $DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.histogram.Keys) {
            $basePercent = $DataObj.data.$OPivotKey.$Property.$IPivotKey.baseline.histogram.$bucket / $baseTotal
            $testPercent = $DataObj.data.$OPivotKey.$Property.$IPivotKey.test.histogram.$bucket / $testTotal
            if ($basePercent -ne 0) {
                $diff          = $testPercent - $basePercent
                $percentChange = 100 * ($diff / $basePercent)
                
                $DataObj.data.$OPivotKey.$Property.$IPivotKey."% change".histogram.$bucket = $percentChange  
            }    
        }
    }
}

function Place-DataEntry ($DataObj, $Item, $Property, $InnerPivot, $OuterPivot, $Mode) {
    $iPivotKey = $NoPivot
    $oPivotKey = $NoPivot

    if ($InnerPivot) {
        $iPivotKey = $item.$InnerPivot 
    } 
    if ($OuterPivot) {
        $oPivotKey = $item.$OuterPivot    
    }

    if (-not ($DataObj.data.Keys -contains $oPivotKey)) {
        $DataObj.data.$oPivotKey = @{}
    }
    if (-not ($DataObj.data.$oPivotKey.Keys -contains $Property)) {
        $DataObj.data.$oPivotKey.$Property = @{}
    }
    if (-not ($DataObj.data.$oPivotKey.$Property.Keys -contains $iPivotKey)) {
        $DataObj.data.$oPivotKey.$Property.$iPivotKey = @{}
    }
    if (-not ($DataObj.data.$oPivotKey.$Property.$iPivotKey.Keys -contains $Mode)) {
        $DataObj.data.$oPivotKey.$Property.$iPivotKey.$Mode = @{}
    }


    if ($Item.$Property.GetType().Name -eq "Hashtable") {
        Merge-Histograms -DataObj $DataObj -Histogram $Item.$Property -Property $Property -IPivotKey $iPivotKey -OPivotKey $oPivotKey -Mode $Mode
    } 
    else {
        if (-not ($DataObj.data.$oPivotKey.$Property.$iPivotKey.$Mode.Keys -contains "orderedData")) {
            $DataObj.data.$oPivotKey.$Property.$iPivotKey.$Mode.orderedData = [Array] @()
        }
        $DataObj.data.$oPivotKey.$Property.$iPivotKey.$Mode.orderedData += $Item.$Property
    }
}


function Fill-Metrics ($DataObj, $Property, $IPivotKey, $OPivotKey, $Mode) {
    $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData = $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData | Sort
    if ($DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData.Count -gt 0) {
        $stats = Measure-Stats -arr $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData
    } 
    else {
        $stats = @{}
    } 
    $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.stats       = $stats
    $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.percentiles = @{}

    if ($DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData.Count -gt 0) {
        foreach ($percentile in $Percentiles) {
            $idx   = [int] (($percentile / 100) * ($DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData.Count - 1))
            $value = $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData[$idx]

            $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.percentiles.$percentile = $value
        }
    }
}


##
#
##
function Merge-Histograms ($DataObj, $Histogram, $Property, $IPivotKey, $OPivotKey, $Mode) {
    if (-Not ($DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.Keys -contains "histogram")) {
        $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram = @{}
    }

    foreach ($bucket in $Histogram.Keys) {
        if (-not ($DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram.Keys -contains $bucket)) {
            $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram.$bucket = $Histogram.$bucket
        } else {
            $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram.$bucket += $Histogram.$bucket
        }
    }
}


##
# Measure-Stats
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
function Measure-Stats ($Arr) {
    $measures = ($Arr | Measure -Average -Maximum -Minimum -Sum)
    $stats = @{
        "count" = $measures.Count
        "sum"   = $measures.Sum
        "min"   = $measures.Minimum
        "mean"  = $measures.Average
        "max"   = $measures.Maximum
    }
    $N   = $measures.Count
    $Arr = $Arr | Sort

    $squareDiffSum = 0
    $cubeDiffSum   = 0
    $quadDiffSum   = 0
    $curCount      = 0
    $curVal        = $null
    $mode          = $null
    $modeCount     = 0

    foreach ($val in $Arr) {
        if ($val -ne $curVal) {
            $curVal   = $val
            $curCount = 1
        } 
        else {
            $curCount++ 
        }

        if ($curCount -gt $modeCount) {
            $mode      = $val
            $modeCount = $curCount
        }

        $squareDiffSum += [Math]::Pow(($val - $measures.Average), 2)
        $quadDiffSum   += [Math]::Pow(($val - $measures.Average), 4)
    }
    $stats["median"]   = $Arr[[int]($N / 2)]
    $stats["mode"]     = $mode
    $stats["range"]    = $stats["max"] - $stats["min"]
    $stats["std dev"]  = [Math]::Sqrt(($squareDiffSum / ($N - 1)))
    $stats["variance"] = $squareDiffSum / ($N - 1)
    $stats["std err"]  = $stats["std dev"] / [math]::Sqrt($N)

    if ($N -gt 3) {
        $stats["kurtosis"] = (($N * ($N + 1))/( ($N - 1) * ($N - 2) * ($N - 3))) * ($quadDiffSum / [Math]::Pow($stats["variance"], 2)) - 3 * ([Math]::Pow($N - 1, 2) / (($N - 2) * ($N - 3)) )
        foreach ($val in $Arr | Sort) { 
            $cubeDiffSum += [Math]::Pow(($val - $measures.Average) / $stats["std dev"], 3) 
        }
        $stats["skewness"] = ($N / (($N - 1) * ($N - 2))) * $cubeDiffSum
    }
    return $stats
}

