Attribute VB_Name = "InterpolateAndAverage"
Option Explicit

' Module: InterpolateAndAverage
'
' Author: Nicolas Perez
'
' Purpose:  Takes the unprocessed values and interpolates the signal so it will fit a linear scale.
'           The processed signal values are then averaged using a multi-value window averaging method.

'Subroutine: DataProcess()
'
'Purpose:  main driving sub for this module.  Collects the information needed to perform linear interpolation with
'            respect to a linear scale and for averaging those processed signals using a multi-value window
'            averaging method.  Calls two other subroutines to help accomplish this goal.


Sub dataProcess()
    
    ' To slightly speedup execution of this module. The nested for loops are the real culprit.
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Worksheet iterator
    Dim sheetIter As Integer
    
    ' Cycle iterator
    Dim cycleIter As Integer
    
    ' Cycle key cell
    Dim curKeyCell As Range
    
    ' Range of standarized distance column
    Dim stdDisRange As Range
    
    ' Range of unprocessed signal column
    Dim uSignalRange As Range
   
    ' Range of unprocessed distance column
    Dim uDistanceRange As Range
    
    ' Range to write interpolated signal values to.
    Dim interWriteRange As Range
    
    ' Size of window for averaging.
    Dim windowSize As Integer
    
    ' Adjusted range to average (range Of Interpolated Values - (window size -1))
    Dim rangeToAverage As Range
    
    ' Range to write averaged values to.
    Dim writeAvgRange As Range
    
    ' Begining of current range we're working with
    Dim beginRange As Range
    
    ' End of current range we're working with
    Dim endRange As Range
    
    ' Number of cells to offset end of range we're averaging
    Dim rangeOffset As Integer
    
    ' Beginning of range to average column
    Dim beginRangeToAverage As Range
    
    ' End of average range to average column
    Dim endRangeToAverage As Range
    
    ' Loop through data sheets
    For sheetIter = 1 To 1
        
        ' Activate current sheet
        Sheets(sheetIter).Activate
        
        ' Set key cell to A1 at the beginning of a new sheet
        
        Set curKeyCell = Range("A1")
        curKeyCell.Activate
                
        ' All references will be relatively the same from each key cell in all cycles
        For cycleIter = 1 To 1
            
            ' Get window average size
            windowSize = 300
            rangeOffset = 0 - (windowSize - 1)
            
            'Keep key cell as A1 if first cycle,
            ' else go ten columns to the right of last
            ' key cell to get next key cell.
            If cycleIter <> 1 Then
                Set curKeyCell = curKeyCell.Offset(0, 10)
                curKeyCell.Activate
            End If
            
            ' Get range of standarized distance column
            Set beginRange = curKeyCell.Offset(8, 5)
            beginRange.Activate
            Set endRange = ActiveCell.End(xlDown)
            Set stdDisRange = Range(beginRange.Address(0, 0), endRange.Address(0, 0))
            
            ' Get range where non-averaged interpolated signal will be written to
            Set beginRange = beginRange.Offset(0, -1)
            
            ' Get beginning range number for averaging purposes
            Set beginRangeToAverage = beginRange
            Set endRange = endRange.Offset(0, -1)
            
            ' Get adjusted range end for averaging
            Set endRangeToAverage = endRange.Offset(rangeOffset, 0)
            Set interWriteRange = Range(beginRange.Address(0, 0), endRange.Address(0, 0))
            
            ' Get adjusted range of interpolated values for averaging
            Set rangeToAverage = Range(beginRangeToAverage, endRangeToAverage)
            
            ' Get range to write averaged values to
            Set beginRange = beginRange.Offset(0, 3)
            Set endRange = endRange.Offset(0, 3)
            Set endRange = endRange.Offset(rangeOffset, 0)
            Set writeAvgRange = Range(beginRange, endRange)
            
            ' Get range of unprocessed signal
            Set beginRange = curKeyCell.Offset(8, 0)
            beginRange.Activate
            Set endRange = ActiveCell.End(xlDown)
            Set uSignalRange = Range(beginRange.Address(0, 0), endRange.Address(0, 0))
            
            ' Get range of unprocessed distance
            Set beginRange = curKeyCell.Offset(8, 2)
            beginRange.Activate
            Set endRange = ActiveCell.End(xlDown)
            Set uDistanceRange = Range(beginRange.Address(0, 0), endRange.Address(0, 0))
            'MsgBox "about to call interpolateNoAverage"
            
            ' Call interpolateNoAverage with above 4 ranges.
            Call interpolateNoAverage(stdDisRange, uSignalRange, uDistanceRange, interWriteRange)
           'MsgBox "back from interpolate, about to call windowAverage"
           
           ' Call windowAverage with above 3 parameters
            Call windowAverage(windowSize, rangeToAverage, writeAvgRange)
            'MsgBox "back from window average"
        Next cycleIter
        
    Next sheetIter
    'MsgBox "sub about to end"
    ' Restore default application settings.
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

Public Sub interpolateNoAverage(stdDistance As Range, uSignalRange As Range, uDisRange As Range, interToWrite As Range)
    
    ' Set range of standarized distance column
    Dim theRange As Range
    Set theRange = stdDistance
    
    ' Range of unprocessed signal values.
    Dim uArrayRange As Range
    Set uArrayRange = uSignalRange
       
    ' For length of UDistance and USignal arrays
    Dim uArrayLength As Integer
    uArrayLength = Application.WorksheetFunction.CountA(uArrayRange)
       
     ' Get ranges of unprocessed data:
    Dim uSigRange As Range
    Set uSigRange = uArrayRange
    Dim uDistanceRange As Range
    Set uDistanceRange = uDisRange
    
    ' Range to write interpolated signal values.
    Dim rangeToWrite As Range
    Set rangeToWrite = interToWrite
         
    ' Iterator for cells in processed distance series
    Dim d As Range
    
    ' For length of Signal array
    Dim signalArrayLength As Integer
    
    ' Get length of entire standardized distance range
    Dim rangeLength As Integer
    rangeLength = Application.WorksheetFunction.CountA(theRange)
    
    ' Signal array Length is equal to length of standardized distance range
    signalArrayLength = rangeLength
    
    ' Initialized arrays for unprocessed signal and distance values.
    Dim UDistanceArray() As Double
    Dim USignalArray() As Integer
    
    ' Must redimension the arrays,
    ' cannot use non-constant for length parameter in initial dimensioning
    ReDim UDistanceArray(0 To uArrayLength - 1) As Double
    ReDim USignalArray(0 To uArrayLength - 1) As Integer
            
    ' Iterators for range
    Dim signal As Range
    Dim distance As Range
        
    ' Iterator for array population:
    Dim x As Integer
    x = 0
    Dim y As Integer
    y = 0
    Dim jStart As Integer
    jStart = 0
       
    ' Fill signal array:
    For Each signal In uSigRange
        USignalArray(x) = signal.Value
        x = x + 1
    Next signal
    
    ' Fill distance array
    For Each distance In uDistanceRange
        UDistanceArray(y) = distance.Value
        y = y + 1
    Next distance
       
    ' Array that will hold signal interpolation results.
    Dim signalArray() As Integer
    ReDim signalArray(0 To signalArrayLength - 1) As Integer
    
    Dim i As Integer
    i = 0
    Dim j As Integer
    
    ' Holds part of calculation
    Dim valueFromProportion As Double
    valueFromProportion = 0
               
    ' Iterate through standarized distance column
    For Each d In theRange.Cells
                    
        ' For computation
        j = jStart
               
        If UDistanceArray(j) > d.Value Then
            signalArray(i) = USignalArray(j)
            
        Else
            ' Get index of upper bound (tsubh)
            Do While UDistanceArray(j) < d.Value And j < uArrayLength - 1
                j = j + 1
            Loop

            ' Linear interplolation of Signal:
            If UDistanceArray(j) < d.Value Then
                signalArray(i) = USignalArray(j)
            ElseIf UDistanceArray(j) = d.Value Then
                signalArray(i) = USignalArray(j)
            ElseIf (USignalArray(j - 1) > USignalArray(j)) Then
                valueFromProportion = ((d.Value - UDistanceArray(j - 1)) * (USignalArray(j - 1) - USignalArray(j))) / (UDistanceArray(j) - UDistanceArray(j - 1))
                signalArray(i) = USignalArray(j - 1) - valueFromProportion
            Else
                valueFromProportion = ((d.Value - UDistanceArray(j - 1)) * (USignalArray(j) - USignalArray(j - 1))) / (UDistanceArray(j) - UDistanceArray(j - 1))
                signalArray(i) = USignalArray(j - 1) + valueFromProportion
            End If
        End If
                   
        i = i + 1
             
    Next d
       
    ' Write out averaged values from array to current sheet
    rangeToWrite.Value = Application.WorksheetFunction.Transpose(signalArray)

End Sub


' Takes entire range and collects averages of groups of values.
Public Sub windowAverage(winSize As Integer, rangeToAverage As Range, avgToWrite As Range)
    
    ' Set size of window
    Dim windowSize As Integer
    windowSize = winSize
    
    ' Set range we will be averaging
    Dim theRange As Range
    
    ' Range is set from first cell to 8/32 or 16/64 from last cell
    Set theRange = rangeToAverage
    
    ' The range where averages will be written.
    Dim rangeToWrite As Range
    Set rangeToWrite = avgToWrite
           
    ' Iterator for cells in window Range
    Dim c As Range
    
    ' For length of array
    Dim arrayLength As Integer
    
    ' Length of entire range
    Dim rangeLength As Integer
    rangeLength = Application.WorksheetFunction.CountA(theRange)
    
    ' Arraylength is the length of the passed in range
    arrayLength = rangeLength
    
    ' Declare to hold window values
    Dim sum As Double
    Dim counter As Integer
    Dim curAverage As Integer
        
    ' Array that will hold window average results.
    Dim curArray() As Integer
    ReDim curArray(0 To arrayLength) As Integer
    
    ' Iterator for curArray
    Dim i As Integer
    i = 0
        
    ' Iterate through entire range
    For Each c In theRange.Cells
        
        ' Comp for window.
        counter = 0
        sum = 0
    
        ' Iterate through window, sum values
        For counter = 0 To (windowSize - 1)
            sum = sum + c.Offset(counter, 0).Value
        Next counter
        
        ' Determine then place average into curArray
        curAverage = sum / windowSize
        curArray(i) = curAverage
        
        ' Set up curArray index for next loop.
        i = i + 1
             
    Next c
    
    ' pad series with zeros
    For counter = i To 1010
        curArray(i) = 0
    Next counter
    
    ' write out to sheet
    rangeToWrite.Value = Application.WorksheetFunction.Transpose(curArray)
        
End Sub











