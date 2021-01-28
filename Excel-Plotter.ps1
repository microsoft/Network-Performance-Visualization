﻿$MINCOLUMNWIDTH = 7.0


##
# Create-ExcelFile
# -------------
# This function plots tables in excel and can use those tables to generate charts. The function
# receives data in a custom format which it uses to create tables. 
# 
# Parameters
# ----------
# Tables (HashTable[]) - Array of table objects 
# SavePath (String) - Path and filename for where to save excel workbook
#
# Returns
# -------
# (String) - Name of saved file
#
##
function Create-ExcelFile {
    param (
        [Parameter(Mandatory=$true)] 
        [PSObject[]]$Tables, 

        [Parameter(Mandatory)]
        [string]$SavePath
    )

    $excelObject = New-Object -ComObject Excel.Application -ErrorAction Stop
    
    # Can be set to true for debugging purposes
    $excelObject.Visible = $false

    $workbookObject = $excelObject.Workbooks.Add()
    $worksheetObject = $workbookObject.Worksheets.Item(1)
        
    $rowOffset = 1
    $chartNum  = 1
    $first = $true
    foreach ($table in $Tables) {
        if ($table.GetType().Name -eq "string") {
            if ($first) {
                $first = $false
            } 
            else {
                Fit-Cells -Worksheet $worksheetObject 
                $worksheetObject = $workbookObject.worksheets.Add()
            }
            $worksheetObject.Name = $table
            $chartNum = 1
            $rowOffset = 1
            continue
        }

        Fill-ColLabels -Worksheet $worksheetObject -cols $table.cols -startCol ($table.meta.rowLabelDepth + 1) -row $rowOffset | Out-Null
        Fill-RowLabels -Worksheet $worksheetObject -rows $table.rows -startRow ($table.meta.colLabelDepth + $rowOffset) -col 1 | Out-Null
        Fill-Data -Worksheet $worksheetObject -Data $table.data -Cols $table.cols -Rows $table.rows -StartCol ($table.meta.rowLabelDepth + 1) -StartRow ($table.meta.colLabelDepth + $rowOffset) | Out-Null
        if ($table.chartSettings) {
            Create-Chart -Worksheet $worksheetObject -Table $table -StartCol 1 -StartRow $rowOffset -chartNum $chartNum | Out-Null
            $chartNum += 1
        }
            Format-ExcelSheet -Worksheet $worksheetObject -Table $table -RowOffset $rowOffset
        $rowOffset += $table.meta.colLabelDepth + $table.meta.dataHeight + 1
    }
    
    Fit-Cells -Worksheet $worksheetObject

    $workbookObject.SaveAs($savePath, $XLENUM.xlOpenXMLWorkbook) | Out-Null  
    $workbookObject.Saved = $true 
    $workbookObject.Close() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookObject) | Out-Null  

    $excelObject.Quit() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObject) | Out-Null
    [System.GC]::Collect() | Out-Null
    [System.GC]::WaitForPendingFinalizers() | Out-Null

    return [string]$savePath
}


##
# Fit-Cells
# ---------
# This function fits the width and height of cells for a given worksheet
# 
# Parameters
# ----------
# Worksheet (ComObject) - Object representing an excel worksheet
#
# Return
# ------
# None
##
function Fit-Cells ($Worksheet) {
    $Worksheet.UsedRange.Columns.Autofit() | Out-Null
    $Worksheet.UsedRange.Rows.Autofit() | Out-Null
    foreach ($column in $Worksheet.UsedRange.Columns) {
        if ($column.ColumnWidth -lt $MINCOLUMNWIDTH) {
            $column.ColumnWidth = $MINCOLUMNWIDTH
        }
    }
}


##
# Format-ExcelSheet
# -----------------
# This function makes cell formatting changes to the provided table in Excel.
#
# Parameters
# ----------
# WorkSheet (ComObject) - Object containing the current worksheet's internal state
# Table (HashTable) - Object containing formatted data and chart settings
# RowOffset (int) - The row index at which the current table begins (top edge)
#
# Return
# ------
# None
#
##
function Format-ExcelSheet ($Worksheet, $Table, $RowOffset) {
    if ($Table.meta.columnFormats) {
        for ($i = 0; $i -lt $Table.meta.columnFormats.Count; $i++) {
            if ($Table.meta.columnFormats[$i]) {
                $column = $worksheetObject.Range($Worksheet.Cells($RowOffset + $Table.meta.colLabelDepth, 1 + $Table.meta.rowLabelDepth + $i), $Worksheet.Cells($RowOffset + $Table.meta.colLabelDepth + $Table.meta.dataHeight - 1, 1 + $Table.meta.rowLabelDepth + $i))
                $column.select() | Out-Null
                $column.NumberFormat = $Table.meta.columnFormats[$i]
            }
        }
    }
    if ($Table.meta.leftAlign) {
        foreach ($col in $Table.meta.leftAlign) {
            $selection = $Worksheet.Range($Worksheet.Cells($RowOffset, $col), $Worksheet.Cells($RowOffset + $Table.meta.colLabelDepth + $Table.meta.dataHeight - 1, $col))
            $selection.select() | Out-Null
            $selection.HorizontalAlignment = $XLENUM.xlHAlignLeft
        }
    }
    if ($Table.meta.rightAlign) {
        foreach ($col in $Table.meta.rightAlign) {
            $selection = $Worksheet.Range($Worksheet.Cells($RowOffset, $col), $Worksheet.Cells($RowOffset + $Table.meta.colLabelDepth + $Table.meta.dataHeight - 1, $col))
            $selection.select() | Out-Null
            $selection.HorizontalAlignment = $XLENUM.xlHAlignRight
        }
    }
    $selection = $Worksheet.Range($Worksheet.Cells($RowOffset, 1), $Worksheet.Cells($RowOffset + $Table.meta.colLabelDepth + $Table.meta.dataHeight - 1, $Table.meta.rowLabelDepth + $Table.meta.dataWidth))
    $selection.select() | Out-Null
    $selection.BorderAround($XLENUM.xlContinuous, $XLENUM.xlThick) | Out-Null
}


##
# Create-Chart
# ------------
# This function uses a table's chartSettings to create and customize a chart
# that visualizes the table's data. 
#
# Parameters
# ----------
# WorkSheet (ComObject) - Object containing the current worksheet's internal state
# Table (HashTable) - Object containing formatted data and chart settings
# StartRow (int) - The row number on which the top of the already-plotted table begins
# StartCol (int) - The column number on which the left side of the already-plotted table begins
# ChartNum (int) - The index this chart will occupy in the worksheet's internally-stored lisrt of charts
#
# Return
# ------
# None
#
##
function Create-Chart ($Worksheet, $Table, $StartRow, $StartCol, $ChartNum) {
    $chart = $Worksheet.Shapes.AddChart().Chart 

    $width = $Table.meta.dataWidth + $Table.meta.rowLabelDepth
    $height = $Table.meta.dataHeight + $Table.meta.colLabelDepth
    if ($Table.chartSettings.yOffset) {
        $height -= $Table.chartSettings.yOffset
        $StartRow += $Table.chartSettings.yOffset
    }
    if ($Table.chartSettings.xOffset) {
        $width -= $Table.chartSettings.xOffset
        $StartCol += $Table.chartSettings.xOffset
    }
    if ($Table.chartSettings.chartType) {
        $chart.ChartType = $Table.chartSettings.chartType
    }
    $chart.SetSourceData($Worksheet.Range($Worksheet.Cells($StartRow, $StartCol), $Worksheet.Cells($StartRow + $height - 1, $StartCol + $width - 1)))
    
    if ($Table.chartSettings.plotBy) {
        $chart.PlotBy = $Table.chartSettings.plotBy
    }
     
    if ($Table.chartSettings.seriesSettings) {
        foreach($seriesNum in $Table.chartSettings.seriesSettings.Keys) {
            if ($Table.chartSettings.seriesSettings.$seriesNum.hide) {
                $chart.SeriesCollection($seriesNum).format.fill.ForeColor.TintAndShade = 1
                $chart.SeriesCollection($seriesNum).format.fill.Transparency = 1
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.lineWeight) {
                $chart.SeriesCollection($seriesNum).format.Line.weight = $Table.chartSettings.seriesSettings.$seriesNum.lineWeight
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.markerSize) {
                $chart.SeriesCollection($seriesNum).markerSize = $Table.chartSettings.seriesSettings.$seriesNum.markerSize
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.color) {
                $chart.SeriesCollection($seriesNum).Border.Color = $Table.chartSettings.seriesSettings.$seriesNum.color
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.name) {
                $chart.SeriesCollection($seriesNum).Name = $Table.chartSettings.seriesSettings.$seriesNum.name
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.delete) {
                $chart.SeriesCollection($seriesNum).Delete()
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.markerStyle) {
                $chart.SeriesCollection($seriesNum).MarkerStyle = $Table.chartSettings.seriesSettings.$seriesNum.markerStyle
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.markerBackgroundColor) {
                $chart.SeriesCollection($seriesNum).MarkerBackgroundColor = $Table.chartSettings.seriesSettings.$seriesNum.markerBackgroundColor
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.markerForegroundColor) {
                $chart.SeriesCollection($seriesNum).MarkerForegroundColor = $Table.chartSettings.seriesSettings.$seriesNum.markerForegroundColor
            }
            if ($Table.chartSettings.seriesSettings.$seriesNum.markerColor) {
                $chart.SeriesCollection($seriesNum).MarkerBackgroundColor = $Table.chartSettings.seriesSettings.$seriesNum.markerColor
                $chart.SeriesCollection($seriesNum).MarkerForegroundColor = $Table.chartSettings.seriesSettings.$seriesNum.markerColor
            }
        }
    }

    if ($Table.chartSettings.axisSettings) {
        foreach($axisNum in $Table.chartSettings.axisSettings.Keys) {
            if ($Table.chartSettings.axisSettings.$axisNum.min) { 
                $Worksheet.chartobjects($ChartNum).chart.Axes($axisNum).MinimumScale = [decimal] $Table.chartSettings.axisSettings.$axisNum.min
            }
            if ($Table.chartSettings.axisSettings.$axisNum.tickLabelSpacing) {
                $Worksheet.chartobjects($ChartNum).chart.Axes($axisNum).TickLabelSpacing = $Table.chartSettings.axisSettings.$axisNum.tickLabelSpacing
            }
            if ($Table.chartSettings.axisSettings.$axisNum.max) { 
                $Worksheet.chartobjects($ChartNum).chart.Axes($axisNum).MaximumScale = [decimal] $Table.chartSettings.axisSettings.$axisNum.max
            }
            if ($Table.chartSettings.axisSettings.$axisNum.logarithmic) {
                $Worksheet.chartobjects($ChartNum).chart.Axes($axisNum).scaleType = $XLENUM.xlScaleLogarithmic
            }
            if ($Table.chartSettings.axisSettings.$axisNum.title) {
                $Worksheet.chartobjects($ChartNum).chart.Axes($axisNum).HasTitle = $true
                $Worksheet.chartobjects($ChartNum).chart.Axes($axisNum).AxisTitle.Caption = $Table.chartSettings.axisSettings.$axisNum.title
            }
            if ($Table.chartSettings.axisSettings.$axisNum.minorGridlines) {
                $Worksheet.chartobjects($ChartNum).chart.Axes($axisNum).HasMinorGridlines = $true
            }
            if ($Table.chartSettings.axisSettings.$axisNum.majorGridlines) {
                $Worksheet.chartobjects($ChartNum).chart.Axes($axisNum).HasMajorGridlines = $true
            }
        }
    }

    if ($Table.chartSettings.title) {
        $chart.HasTitle = $true
        $chart.ChartTitle.Caption = [string]$Table.chartSettings.title
    }
    if ($Table.chartSettings.dataTable) {
        $chart.HasDataTable = $true
        $chart.HasLegend = $false
    }

    $Worksheet.Shapes.Item("Chart " + $ChartNum ).top = $Worksheet.Cells($StartRow, $StartCol + $width + 1).top
    $Worksheet.Shapes.Item("Chart " + $ChartNum ).left = $Worksheet.Cells($StartRow, $StartCol + $width + 1).left
}


##
# Fill-Cell
# ---------
# This function fills an excel cell with a value, and optionally also customizes the cell style.
# 
# Parameters
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Row (int) - Row index of the cell to fill
# Col (int) - Column index of the cell to fill
# CellSettings (HashTable) - Object containing the value and style settings for the cell
#
# Return
# ------
# None
#
##
function Fill-Cell ($Worksheet, $Row, $Col, $CellSettings) {
    $Worksheet.Cells($Row, $Col).Borders.LineStyle = $XLENUM.xlContinuous
    if ($CellSettings.fontColor) {
        $Worksheet.Cells($Row, $Col).Font.Color = $CellSettings.fontColor
    }

    if ($CellSettings.cellColor) {
        $Worksheet.Cells($Row, $Col).Interior.Color = $CellSettings.cellColor
    }

    if ($CellSettings.bold) {
        $Worksheet.Cells($Row, $Col).Font.Bold = $true
    }

    if ($CellSettings.center) {
        $Worksheet.Cells($Row, $Col).HorizontalAlignment = $XLENUM.xlHAlignCenter
        $Worksheet.Cells($Row, $Col).VerticalAlignment = $XLENUM.xlHAlignCenter
    }
    
    if ($null -ne $CellSettings.value) {

        $Worksheet.Cells($Row, $Col) = $CellSettings.value
    }
}

##
# Merge-Cells
# -----------
# This function merges a range of cells into a single cell and adds a border
#
# Parameters
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Row1 (int) - Row index of top left cell of range to merge
# Col1 (int) - Column index of top left cell of range to merge
# Row2 (int) - Row index of bottom right cell of range to merge
# Col2 (int) - Column index of bottom right cell of range to merge
#
# Return
# ------
# None
#
##
function Merge-Cells ($Worksheet, $Row1, $Col1, $Row2, $Col2) {
    $cells = $Worksheet.Range($Worksheet.Cells($Row1, $Col1), $Worksheet.Cells($Row2, $Col2))
    $cells.Select()
    $cells.MergeCells = $true
    $cells.Borders.LineStyle = $XLENUM.xlContinuous
}


##
# Fill-ColLabels
# --------------
# This function consumes the cols field of a table object, and plots the column labels by recursing 
# through the object. 
#
# Parameters 
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Cols (HashTable) - Object storing column label structure and column indices of labels. 
# StartCol (int) - The column index on which the labels should start being drawn (left edge)
# Row (int) - The row at which the current level of labels should be drawn
#
# Return 
# ------
# (int[]) - Tuple of integers capturing the column index range across which the just-drawn label spans
#
##
function Fill-ColLabels ($Worksheet, $Cols, $StartCol, $Row) {
    $range = @(-1, -1)
    foreach ($label in $Cols.Keys) {
        if ($Cols.$label.GetType().Name -ne "Int32") {
            $subRange = Fill-ColLabels -Worksheet $Worksheet -Cols $Cols.$label -StartCol $StartCol -Row ($Row + 1)
            Merge-Cells -Worksheet $Worksheet -Row1 $Row -Col1 $subRange[0] -Row2 $Row -Col2 $subRange[1] | Out-Null
            $cellSettings = @{
                "value" = $label
                "bold" = $true
                "center" = $true
            }
            Fill-Cell -Worksheet $Worksheet -Row $Row -Col $subRange[0] -CellSettings $cellSettings | Out-Null
            if (($subRange[0] -lt $range[0]) -or ($range[0] -eq -1)) {
                $range[0] = $subRange[0]
            } 
            if (($subRange[1] -gt $range[1]) -or ($range[0] -eq -1)) {
                $range[1] = $subRange[1]
            }
        } 
        else {
            $cellSettings = @{
                "value" = $label
                "bold" = $true
                "center" = $true
            }
            Fill-Cell $Worksheet -Row $Row -Col ($StartCol + $Cols.$label) -CellSettings $cellSettings | Out-Null
            if (($StartCol + $Cols.$label -lt $range[0]) -or ($range[0] -eq -1)) {
                $range[0] = $StartCol + $Cols.$label
            }
            if (($StartCol + $Cols.$label -gt $range[1]) -or ($range[1] -eq -1)) {
                $range[1] = $StartCol + $Cols.$label
            }
        }    
    }
    return $range
}


##
# Fill-RowLabels
# --------------
# This function consumes the rows field of a table object, and plots the row labels by recursing 
# through the object. 
#
# Parameters 
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Rows (HashTable) - Object storing row label structure and row indices of labels. 
# StartRow (int) - The row index on which the labels should start being drawn (top edge)
# Col (int) - The column at which the current level of labels should be drawn
#
# Return 
# ------
# (int[]) - Tuple of integers capturing the row index range across which the just-drawn label spans
#
##
function Fill-RowLabels ($Worksheet, $Rows, $StartRow, $Col) {
    $range = @(-1, -1)
    foreach ($label in $rows.Keys) {
        if ($Rows.$label.GetType().Name -ne "Int32") {
            $subRange = Fill-RowLabels -Worksheet $Worksheet -Rows $Rows.$label -StartRow $StartRow -Col ($Col + 1)
            Merge-Cells -Worksheet $Worksheet -Row1 $subRange[0] -Col1 $Col -Row2 $subRange[1] -Col2 $Col | Out-Null
            $cellSettings = @{
                "value" = $label
                "bold" = $true
                "center" = $true
            }
            Fill-Cell -Worksheet $Worksheet -Row $subRange[0] -Col $Col -CellSettings $cellSettings | Out-Null
            if (($subRange[0] -lt $range[0]) -or ($range[0] -eq -1)) {
                $range[0] = $subRange[0]
            } 
            if (($subRange[1] -gt $range[1]) -or ($range[0] -eq -1)) {
                $range[1] = $subRange[1]
            }
        } 
        else {
            $cellSettings = @{
                "value" = $label
                "bold" = $true
                "center" = $true
            }
            Fill-Cell $Worksheet -Row ($StartRow + $Rows.$label) -Col $Col -CellSettings $cellSettings | Out-Null
            if (($StartRow + $Rows.$label -lt $range[0]) -or ($range[0] -eq -1)) {
                $range[0] = $StartRow + $Rows.$label
            }
            if (($StartRow + $Rows.$label -gt $range[1]) -or ($range[1] -eq -1)) {
                $range[1] = $StartRow + $Rows.$label
            }
        }    
    }
    return $range
}


##
# Fill-Data
# ---------
# This function uses the data, rows, and cols fields of a table object to fill in the data
# values of a table. The objects are recursed through depth-first, and the path followed is used to retreive 
# row and column indices from the Rows and Cols objects while data values and cell formatting are retreived from the
# data object. 
#
# Parameters
# ----------
# Worksheet (ComObject) - Object containing the current worksheet's internal state
# Data (HashTable) - Object containing data values and cell formatting
# Cols (HashTable) - Object storing column label structure and column indices of the labels. 
# Rows (HashTable) - Object storing row label structure and row indices of the labels.
# StartCol (int) - Column index where the data range of the table begins (left edge) 
# StartRow (int) - Row index where the data range of the table begins (top edge)
#
# Return
# ------
# None
#
## 
function Fill-Data ($Worksheet, $Data, $Cols, $Rows, $StartCol, $StartRow) {
    if($Cols.GetType().Name -eq "Int32" -and $Rows.GetType().Name -eq "Int32") {
        Fill-Cell -Worksheet $Worksheet -Row ($StartRow + $Rows) -Col ($StartCol + $Cols) -CellSettings $Data
        return
    }  
    foreach ($label in $Data.Keys) {
        if ($Cols.getType().Name -ne "Int32") {
            Fill-Data -Worksheet $Worksheet -Data $Data.$label -Cols $Cols.$label -Rows $Rows -StartCol $StartCol -StartRow $StartRow
        } 
        else {
            Fill-Data -Worksheet $Worksheet -Data $Data.$label -Cols $Cols -Rows $Rows.$label -StartCol $StartCol -StartRow $StartRow
        }
    }
}