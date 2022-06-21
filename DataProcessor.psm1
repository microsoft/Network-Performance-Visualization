$Percentiles = @(0, 1, 5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90, 95, 96, 97, 98, 99, 99.9, 99.99, 99.999, 99.9999, 99.99999, 100)

##
# Process-Data
# ------------
# This function organizes raw data into subsets, and then calculates statistics and percentiles 
# over each sub-category of data. Subsets are delineated by three values: a Property, an 
# innerPivot value (inner pivot key), and an outerPivot value (outer pivot key). Hence, values are extracted from
# raw dataEntry objects and placed into subsets based on the property name whose value is being extracted and the values 
# of pivot properties of the same dataEntry object. Processed data, the original raw data, and some meta data are 
# then stored together in an object and returned. 
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
        [Parameter(Mandatory=$true)] 
        [PSobject[]] $BaselineRawData,

        [Parameter()] 
        [PSobject[]] $TestRawData,

        [Parameter()] 
        [String] $InnerPivot,

        [Parameter()] 
        [String] $OuterPivot,

        [Parameter()]  
        [AllowNull()] 
        [Int] $NumHistogramBuckets
    )
    if ($NumHistogramBuckets -eq $null) {
        $NumHistogramBuckets = 50
    }
    $meta = $BaselineRawData.meta
    if ($TestRawData) {
        $meta = Merge-MetaData -BaselineMeta $BaselineRawData.meta -TestMeta $TestRawData.meta
    }

    $processedDataObj = @{
        "meta"    = $meta
        "data"    = @{}
        "rawData" = @{
            "baseline" = $BaselineRawData.data
        }
    }

    $processedDataObj.meta.innerPivot = $InnerPivot
    $processedDataObj.meta.outerPivot = $OuterPivot

    if ($processedDataObj.meta) { 
        $processedDataObj.rawData.test = $TestRawData.data
        
    }

    $modes = if ($TestRawData) { "baseline", "test" } else { ,"baseline" }


    $numProps = $processedDataObj.meta.props.Count
    $baselineDataCount = $BaselineRawData.data.Count
    $testDataCount = if ($TestRawData) {$TestRawData.data.Count} else {0}
    $totalIters = $numProps * ($baselineDataCount + $testDataCount + `
                    ($processedDataObj.meta.innerPivotKeys.Count * $processedDataObj.meta.outerPivotKeys.Count * $modes.Count))

    $i = 0
    foreach ($prop in $processedDataObj.meta.props) {
        if ($BaselineRawData.meta.noTable -contains $prop) {
            continue
        }

        # Extract property values from dataEntry objects and place values in the correct spot within the processedData object
        foreach($item in $BaselineRawData.data) {
            Place-DataEntry -DataObj $processedDataObj -DataEntry $item -Property $prop -InnerPivot $InnerPivot -OuterPivot $OuterPivot -Mode "baseline"
            Write-Progress -Activity "Processing Raw Data..." -Status "Processing..." -Id 2 -PercentComplete (100 * (($i++) / $totalIters))
        }

        if ($meta.comparison){
            foreach ($item in $TestRawData.data) {
                Place-DataEntry -DataObj $processedDataObj -DataEntry $item -Property $prop -InnerPivot $InnerPivot -OuterPivot $OuterPivot -Mode "test"
                Write-Progress -Activity "Processing Raw Data..." -Status "Processing..." -Id 2 -PercentComplete (100 * (($i++) / $totalIters))
            }
        }


        foreach ($oPivotKey in $processedDataObj.data.Keys) {
            foreach ($iPivotKey in $processedDataObj.data.$oPivotKey.$prop.Keys) {
                foreach ($mode in $modes) {
                    if ($processedDataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.orderedData) {
                        Add-OrderedDataStats -DataObj $processedDataObj -Property $prop -IPivotKey $iPivotKey -OPivotKey $oPivotKey -Mode $mode
                    }
                }

                foreach ($mode in $modes) {
                    if (-not $processedDataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.percentiles -and `
                            $processedDataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.histogram) {
                        Percentiles-FromHistogram -DataObj $processedDataObj -Property $prop -IPivotKey $iPivotKey -OPivotKey $oPivotKey -Mode $mode 
                    } 
                    
                    if (-not $processedDataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.histogram -and `
                            $processedDataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.orderedData ) {
                        Histogram-FromOrderedData -DataObj $processedDataObj -Property $prop -IPivotKey $iPivotKey -OPivotKey $oPivotKey -Mode $mode -NumBuckets $NumHistogramBuckets
                    }

                    Write-Progress -Activity "Processing Raw Data..." -Status "Processing..." -Id 2 -PercentComplete (100 * (($i++) / $totalIters))
                }
            }
        }
    }
    Write-Progress -Activity "Processing Raw Data..." -Status "Done"-Id 2 -PercentComplete 100
    return $processedDataObj
}

function Merge-MetaData ($BaselineMeta, $TestMeta) {
    $BaselineMeta.comparison = $true

    $BaselineMeta.props = $BaselineMeta.props | ?{$TestMeta.props -contains $_} 

    # Keep only the intersection of pivot keys between the two datasets 
    $BaselineMeta.innerPivotKeys = $BaselineMeta.innerPivotKeys | ?{$TestMeta.innerPivotKeys -contains $_} 
    $BaselineMeta.outerPivotKeys = $BaselineMeta.outerPivotKeys | ?{$TestMeta.outerPivotKeys -contains $_}  
    return $BaselineMeta
}
function Invert-Mode($Mode) {
    if ($Mode -eq "baseline") {
        return "test"
    } else {
        return "baseline"
    }
}

function Histogram-FromOrderedData ($DataObj, $Property, $IPivotKey, $OPivotKey, $Mode, $NumBuckets = 50) {
    $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram = @{}

    
    $otherMode = Invert-Mode($Mode)
    $buckets = [Array] @()
    if (-not $DataObj.data.$OPivotKey.$Property.$IPivotKey.$otherMode.histogram) {
        
        $min = $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData[0] 
        $max = $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData[-1]

        if ($DataObj.data.$OPivotKey.$Property.$IPivotKey.$otherMode.orderedData) {

            if ($DataObj.data.$OPivotKey.$Property.$IPivotKey.test.orderedData[0] -lt $min) {
                $min = $DataObj.data.$OPivotKey.$Property.$IPivotKey.test.orderedData[0]
            }

            if ($DataObj.data.$OPivotKey.$Property.$IPivotKey.test.orderedData[-1] -gt $max) {
                $max = $DataObj.data.$OPivotKey.$Property.$IPivotKey.test.orderedData[-1]
            }
        }

        $bucketSize = ($max - $min) / $NumBuckets
        $curBucket = $min
        $buckets = [Array]@()


        for($i = 0; $i -lt $NumBuckets; $i++) { 
            $buckets += $curBucket
            $curBucket += $bucketSize
        }
    } 
    else {
        $buckets = $DataObj.data.$OPivotKey.$Property.$IPivotKey.$otherMode.histogram.Keys | Sort-Object
    }

    $bucketIdx = 0
    $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram[$buckets[$bucketIdx]] = 0
    foreach ($value in $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.orderedData) {  
        while ($bucketIdx -lt ($buckets.Count - 1) -and $value -ge $buckets[$bucketIdx + 1]) {   
            $bucketIdx++ 
            $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram[$buckets[$bucketIdx]] = 0 
        }
        $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode.histogram[$buckets[$bucketIdx]]++ 
    }
}



##
# Place-DataEntry
# ---------------
# This function extracts raw data values from dataEntry objects, and places them in the correct
# position within the processed data object.
#
# Parameters
# ----------
# DataObj (HashTable) - Processed data object
# DataEntry (HashTable) - DataEntry object whose data is being added to processed data object
# Property (String) - Name of the property whose value should be extracted from the dataEntry
# InnerPivot (String) - Name of the property to use as an inner pivot
# OuterPivot (String) - Name of the property to use as an outer pivot
# Mode (String) - Mode (baseline/test) of the given dataEntry object
# 
# Return
# ------
# None
#
function Place-DataEntry ($DataObj, $DataEntry, $Property, $InnerPivot, $OuterPivot, $Mode) {
    $iPivotKey = if ($InnerPivot) {$DataEntry.$InnerPivot} else {""}
    $oPivotKey = if ($OuterPivot) {$DataEntry.$OuterPivot} else {""}

    if (-not $DataEntry.ContainsKey($Property)) { return }
    if ($null -eq $DataEntry.$Property) { return }

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


    if ($DataEntry.$Property.GetType().Name -eq "Hashtable") { # $Item.$Property should be $DataEntry.$Property?
        Merge-Histograms -DataObj $DataObj -Histogram $DataEntry.$Property -Property $Property -IPivotKey $iPivotKey -OPivotKey $oPivotKey -Mode $Mode
    } else {
        if (-not ($DataObj.data.$oPivotKey.$Property.$iPivotKey.$Mode.ContainsKey("orderedData"))) {
            $DataObj.data.$oPivotKey.$Property.$iPivotKey.$Mode.orderedData = [Array] @()
        }
        $DataObj.data.$oPivotKey.$Property.$iPivotKey.$Mode.orderedData += $DataEntry.$Property
    }
    
    
}

<#
.SYNOPSIS
    Calculate metrics from the ordered data of a given data subset,
    adding them to the data object.
.PARAMETER DataObj
    Processed data object.
.PARAMETER Property
    Name of the property of the data subset.
.PARAMETER IPivotKey
    Inner pivot of the data subset.
.PARAMETER OPivotKey
    Outer pivot of the data subset.
.PARAMETER Mode
    baseline or test.
#>
function Add-OrderedDataStats($DataObj, $Property, $IPivotKey, $OPivotKey, $Mode) {
    $dataModel = $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode

    $dataModel.stats = [Ordered]@{}
    $dataModel.percentiles = @{} # TODO ordered

    if ($dataModel.orderedData.Count -eq 0) {
        return
    }

    $dataModel.orderedData = $dataModel.orderedData | sort
    $stat = ($dataModel.orderedData | measure -Sum -Average -Maximum -Minimum)
    $n = $stat.Count

    if ($n -eq 0) {return}

    $variance = ($dataModel.orderedData | foreach {[Math]::Pow($_ - $stat.Average, 2)} | measure -Average).Average
    $stdDev = [Math]::Sqrt($variance)



    $dataModel.stats = [Ordered]@{
        "n"        = $n
        "sum"      = $stat.Sum
        "mean"     = $stat.Average
        "median"   = if ($n % 2) {$dataModel.orderedData[[Math]::Floor($n / 2)]} else {0.50 * ($dataModel.orderedData[$n / 2] + $dataModel.orderedData[($n / 2) - 1])}
        "mode"     = ($dataModel.orderedData | group -NoElement | sort -Property Count)[-1].Name
        "min"      = $stat.Minimum
        "max"      = $stat.Maximum
        "range"    = $stat.Maximum - $stat.Minimum
        "variance" = $variance
        "std dev"  = $stdDev
        "std err"  = $stdDev / [Math]::Sqrt($n)
    }

    if (($n -gt 3) -and ($stdDev -ne 0)) {
        $s1 = $n / (($n - 1) * ($n - 2))
        $k1 = $s1 * (($n + 1) / ($n - 3))
        $k2 = 3 * ((($n - 1) * ($n - 1)) / (($n - 2) * ($n - 3)))

        $cubeDiffs = $dataModel.orderedData | foreach {[Math]::Pow(($_ - $stat.Average) / $stdDev, 3)}
        $quadDiffs = $dataModel.orderedData | foreach {[Math]::Pow($_ - $stat.Average, 4)} 

        $dataModel.stats["skewness"] = $s1 * ($cubeDiffs | measure -Sum).Sum
        $dataModel.stats["kurtosis"] = $k1 * (($quadDiffs | measure -Sum).Sum / [Math]::Pow($stdDev, 4)) - $k2
    }

    # Fill out percentiles
    foreach ($percentile in $Percentiles) {
        $i = [Int](($percentile / 100) * ($dataModel.orderedData.Count - 1))
        $dataModel.percentiles["$percentile"] = $dataModel.orderedData[$i]
    }
}


##
# Percentiles-FromHistogram
# -------------------------
# This function uses a histogram stored in the processed data object to calculate approximate
# percentiles for a subset of data.
# 
# Parameters
# ----------
# DataObj (HashTable) - Processed data object
# Property (String) - Name of the property of the data subset whose histogram is being used (ex: latency)
# IPivotKey (String) - Value of the inner pivot of the data subset whose histogram is being used
# OPivotKey (String) - Value of the outer pivot of the data subset whose histogram is being used
# Mode (String) - Mode (baseline/test) of the data subset whose histogram is being used
#
# Return
# ------
# None
#
##
function Percentiles-FromHistogram ($DataObj, $Property, $IPivotKey, $OPivotKey, $Mode) {
    $dataModel = $DataObj.data.$OPivotKey.$Property.$IPivotKey.$Mode
    # Calculate cumulative density function
    $cdf = @{}
    $sumSoFar = 0
    foreach ($bucket in ($dataModel.histogram.Keys | Sort)) {
        $sumSoFar += $dataModel.histogram.$bucket
        $cdf.$bucket = $sumSoFar 
    }

    # Convert to pecentages
    $buckets = [System.Collections.Queue]@($cdf.Keys | Sort)
    foreach ($bucket in $buckets) {
        $cdf.$bucket = 100 * ($cdf.$bucket / $sumSoFar)
    }
    
    $dataModel.percentiles = @{}

    $prevBucket = $null
    $bucket = $buckets.Dequeue()
    foreach ($percentile in $Percentiles) {
        # Skip buckets irrevalent to current percentile calculation
        while ($cdf.$bucket -lt $percentile) {
            $prevBucket = $bucket
            $bucket = $buckets.Dequeue()
        }

        if ($null -eq $prevBucket) {
            $dataModel.percentiles["$percentile"] = $bucket
        } 
        else {
            # Approx. the desired percentile via linear interpolation
            $interp = ($percentile - $cdf.$prevBucket) / ($cdf.$bucket - $cdf.$prevBucket)
            $approxPercentile = ($interp * $bucket) + ((1 - $interp) * $prevBucket)

            $dataModel.percentiles["$percentile"] = $approxPercentile
        }
    }
}

##
# Merge-Histograms
# ----------------
# This function merges a given histogram with a specified data subset's histogram in the processed data object. 
# 
# Parameters
# ----------
# DataObj (HashTable) - Processed data object
# Histogram (HashTable) - New histogram to merge with the specified data subset's histogram
# Property (String) - Name of the property of the data subset for which histograms are being merged
# IPivotKey (String) - Value of the inner pivot of the data subset for which histograms are being merged
# OPivotKey (String) - Value of the outer pivot of the data subset for which histograms are being merged
# Mode (String) - Mode (baseline/test) of the data subset whose histograms are being merged
#
# Return
# ------
# None 
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


