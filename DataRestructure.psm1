
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
            if ($prop -notin $prevDepth.keys) { 
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
                        $prevDepth["$mode"] = @{}
                        $null = $depthStack.Push($prevDepth["$mode"])
                    } else {
                        $null = $depthStack.Push($prevDepth)
                    }

                    foreach ($metric in $DataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.keys) {
                        if ($metric -in @("orderedData", "histogram")) { continue }
                        $prevDepth = $depthStack.Peek()
                        $prevDepth.$metric = @{} 
                        foreach ($measure in $DataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.$metric.keys) { 
                            $prevDepth.$metric[$measure] = [String]$DataObj.data.$oPivotKey.$prop.$iPivotKey.$mode.$metric.$measure
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




    $output.meta = $meta
    return $output # ConvertNestedHashTableTo-Json($output)
}

function Parse-RestructedData ($dataObj) {
    $processedDataObj = @{
        "meta" = $dataObj.meta
        "data" = @{}
    }

    $meta = $dataObj.meta 
    
    $depthStack = [System.Collections.Stack]@() 

    foreach ($oPivotKey in $meta.OuterPivotKeys) {
        $usedOPivotLabel = ($meta.OuterPivot -ne "")
        $oPivotLabel = "$($oPivotKey)-$($meta.OuterPivot)"
        
        $processedDataObj.data[$oPivotKey] = @{}
        $null = $depthStack.push($processedDataObj.data.$oPivotKey)
         
        foreach ($prop in $meta.props) {

            $prevDepth = $depthStack.Peek()
            if ($prop -notin $prevDepth.keys) {
                $prevDepth[$prop] = @{}
            }
            $null = $depthStack.Push($prevDepth[$prop])

            foreach ($iPivotKey in $meta.InnerPivotKeys) {
                $usedIPivotLabel = ($meta.InnerPivot -ne "") 
                $iPivotLabel = "$($iPivotKey)-$($meta.InnerPivot)"
                
                $prevDepth = $depthStack.Peek() 
                $prevDepth[$iPivotKey] = @{}
                $null = $depthStack.Push($prevDepth[$iPivotKey]) 

                
                if ($usedOPivotLabel) {
                    $inputObjDepth = $dataObj.$oPivotLabel.$prop
                } 
                else {
                    $inputObjDepth = $dataObj.$prop 
                }

                if ($usedIPivotLabel) {
                    $inputObjDepth = $inputObjDepth.$iPivotLabel
                }
                
                $modes = @("baseline")
                if ($meta.comparison) {
                    $modes += "test"
                }

                foreach ($mode in $modes) { 

                    $prevDepth = $depthStack.Peek()
                    $prevDepth[$mode] = @{}
                    $depthStack.push($prevDepth[$mode]) 

                    
                    $inputFinalDepth = $inputObjDepth
                    if ($usedModeLabel) {
                        $inputFinalDepth = $inputFinalDepth["$mode"]
                    }

                    foreach ($metric in $inputFinalDepth.keys) { 
                        $prevDepth = $depthStack.Peek()
                        $prevDepth.$metric = @{} 
                        foreach ($measure in $inputFinalDepth.$metric.keys) {
                            $prevDepth.$metric[[String]$measure] = [Decimal] $inputFinalDepth.$metric.$measure
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

    return $processedDataObj
}