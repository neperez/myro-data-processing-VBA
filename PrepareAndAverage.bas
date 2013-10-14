Attribute VB_Name = "PrepareAndAverage"
Option Explicit

' Module:  PrepareAndAverage
' Author:  Nicolas Perez
' Purpose:  relabels averaging column for the processed signal.  Clears old averages and linear distance columns.
'           Sets up new linear distance scale, performs new averaging, and writes result to worksheet.
'
Private Sub PrepareAndAverage()
    
    ' This is an attempt to speed up the execution.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Worksheet iterator
    Dim sheetIter As Integer
    
    ' Cycle iterator
    Dim cycleIter As Integer
    
    ' Cycle key cell
    Dim curKeyCell As Range
    
    ' The beginning cell of the current range.
    Dim beginRange As Range
    
    ' The end cell of the current range.
    Dim endRange As Range
    
    ' Total distance range
    Dim totalRange As Range
    
    ' Step value for filling the linear distance column.
    Dim stepValue As Double
    stepValue = 0.02
    
    ' Distance traveled/stop value for dataseries fill.
    Dim actualDistanceCell As Range
    Dim stopValue As Double
    Dim endRangeNumber As Integer
    
    ' Window size for averaging.
    Dim windowSize As Integer
    windowSize = 300
    
    ' Holds range to average
    Dim rangeToAverage As Range
    
    ' holds write location for averaged values.
    Dim rangeToWrite As Range
    
    ' offset for averages
    Dim rangeOffset As Integer
    rangeOffset = (0 - (windowSize - 1))
    Dim beginAverage As Range
    Dim endAverage As Range
    Dim totalAverage As Range
    
    ' loop through first nine sheets
    For sheetIter = 1 To 9
        
        ' activate current sheet
        Sheets(sheetIter).Activate
        
        ' set key cell to H7 at the beginning of a new sheet
        Set curKeyCell = Range("H7")
        curKeyCell.Activate
        
        ' all references will be relatively the same from each key cell in all cycles
        For cycleIter = 1 To 30
            
            'Keep key cell as A1 if first cycle,
            ' else go ten columns to the right of last
            ' key cell to get next key cell.
            If cycleIter <> 1 Then
                Set curKeyCell = curKeyCell.Offset(0, 10)
                curKeyCell.Activate
            End If
            
            ' get actual distance travelled in the current cycle.
            Set actualDistanceCell = curKeyCell.Offset(-4, -1)
            stopValue = actualDistanceCell.Value
            
            ' change heading of keycell to 300 value average
            curKeyCell.Value = "300 value average:"
            endRangeNumber = stopValue / stepValue
            
            ' clear current averaged signal range values
            ' set begin range to beginning of signal column
            Set beginRange = curKeyCell.Offset(2, 0)
            
            ' set endrange to ending cell for clearing
            Set endRange = beginRange.Offset(1214, 0)
            
            ' set total range
            Set totalRange = Range(beginRange, endRange)
            totalRange.Clear
            
            ' clear standarized distance column.
            Set beginRange = beginRange.Offset(0, 1)
            Set endRange = beginRange.Offset(1214, 0)
            Set totalRange = Range(beginRange, endRange)
            totalRange.Clear
            beginRange.Value = 6
            
            ' refill series on std. distance range values.
            ' set endrange to correct cell
            Set endRange = beginRange.Offset(endRangeNumber, 0)
            
            ' place beginning and end into one range
            Set totalRange = Range(beginRange, endRange)
            totalRange.DataSeries Rowcol:=xlColumns, Type:=xlDataSeriesLinear, Step:=stepValue, Stop:=stopValue
            
            ' Set up parameters to average signal values using 300 value window.
            Set beginAverage = curKeyCell.Offset(2, -3)
            beginAverage.Activate
            Set endAverage = ActiveCell.End(xlDown)
            Set endAverage = endAverage.Offset(rangeOffset, 0)
            Set rangeToAverage = Range(beginAverage, endAverage)
            Set beginRange = beginAverage.Offset(0, 3)
            Set endRange = endAverage.Offset(0, 3)
            Set rangeToWrite = Range(beginRange, endRange)
            
            ' call window average subroutine in AverageAndInterpolate module.
            Call AverageAndInterpolate.windowAverage(windowSize, rangeToAverage, rangeToWrite)
        
        Next cycleIter
    
    Next sheetIter
    
    ' restore speed related excel events.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

