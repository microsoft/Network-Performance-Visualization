
function Restructure-Data ($DataObj) {
    $meta = $DataObj.meta 
    $output = @{} 
    $depthStack = [System.Collections.Stack]@()

    foreach ($oPivotKey in $meta.OuterPivotKeys) {
        $useOPivotLabel = $false
        if ($meta.OuterPivot -ne "") {
            $useOPivotLabel = $true
            $oPivotLabel = "$($oPivotKey)-$($meta.OuterPivot)"
        }

        if ($useOPivotLabel) {
            $output[$oPivotLabel] = @{}
            $null = $depthStack.push($output.$oPivotLabel)
        } 
        else {
            
            $null = $depthStack.push($output)
        }
         

        foreach ($prop in $meta.props) {

            $prevDepth = $depthStack.Peek()
            if ($prop -notin $prevDepth) { 
                $prevDepth[$prop] = @{}
            }
            $null = $depthStack.Push($prevDepth[$prop])

            foreach ($iPivotKey in $meta.InnerPivotKeys) {
                $useIPivotLabel = $false
                if ($meta.InnerPivot -ne "") {
                    $useIPivotLabel = $true
                    $iPivotLabel = "$($iPivotKey)-$($meta.InnerPivot)"
                }
                
                $prevDepth = $depthStack.Peek()
                if ($useIPivotLabel) {
                    $prevDepth[$iPivotLabel] = @{}
                    $null = $depthStack.Push($prevDepth[$iPivotLabel])
                } else {
                    $null = $depthStack.Push($prevDepth)
                } 

                foreach ($mode in $DataObj.data.$oPivotKey.$prop.$iPivotKey.keys) {
                    $prevDepth = $depthStack.Peek()
                    if ($meta.comparison) {
                        $prevDepth[$mode] = @{}
                        $null = $depthStack.Push($prevDepth[$mode])
                    } else {
                        $null = $depthStack.Push($prevDepth)
                    }

                    foreach ($metric in $DataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.keys) {
                        if ($metric -in @("orderedData", "histogram")) { continue }
                        $prevDepth = $depthStack.Peek()
                        $prevDepth.$metric = @{} 
                        foreach ($measure in $DataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.$metric.keys) {
                            $prevDepth.$metric[[String]$measure] = [String]$DataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.$metric.$measure
                        } 
                    } 
                    $null = $depthStack.Pop()
                }
                $null = $depthStack.Pop()
            }
            $null = $depthStack.Pop()
        }
        $null = $depthStack.Pop()
    }





    return $output # ConvertNestedHashTableTo-Json($output)
}


function ConvertNestedHashTableTo-Json ($object) {
    if ($object.GetType().Name -ne "Hashtable") {
        return $object
    }

    $newObject = @{}
    foreach ($key in $object.keys) {
        $newObject["$key"] = ConvertNestedHashTableTo-Json($object[$key])
    }

    return ($newObject | ConvertTo-Json)
} 