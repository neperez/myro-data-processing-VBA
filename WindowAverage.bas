Attribute VB_Name = "WindowAverage"
Option Explicit

' Module:  WindowAverage
'
' Authors:  Nicolas Perez
'
' Purpose:  Averages signal values in groups for smoothing purposes.
'

Public Sub WindowAverage()
        
    ' Set size of window
    Dim windowSize As Integer
    windowSize = 400 '<-------------------- Enter window size
    
    ' Set range we will be averaging
    Dim theRange As Range
    
    ' Range is set from first cell to 8/32 or 16/64 from last cell
    Set theRange = Range("M8:M558") '<-------------------- Enter adjusted range.
    
    ' The range where averages will be written.
    Dim rangeToWrite As Range
    Set rangeToWrite = Range("AH8", "AH1010") '<----------------------- Enter write range
           
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
    
    ' Pad series with zeros
    For counter = i To 1010
        curArray(i) = 0
    Next counter
    
    ' Write out to sheet
    rangeToWrite.Value = Application.WorksheetFunction.Transpose(curArray)
        
End Sub

