Attribute VB_Name = "Module1"
Option Explicit

' takes entire range and collects averages of groups of values.
Public Sub windowAverage()
    
    ' set size of window
    Dim windowSize As Integer
    windowSize = 32 '--------------------
    
    ' set range we will be averaging
    Dim theRange As Range
    
    ' range is set from first cell to 8/32 or 16/64 from last cell
    Set theRange = Range("M36:M157") '--------------------
       
    ' iterator for cells in window Range
    Dim c As Range
    
    ' for length of array
    Dim arrayLength As Integer
    
    ' length of entire range
    Dim rangeLength As Integer
    rangeLength = Application.WorksheetFunction.CountA(theRange)
    
    'arraylength is the length of the passed in range
    arrayLength = rangeLength
    
    ' declare to hold window values
    Dim sum As Double
    Dim counter As Integer
    Dim curAverage As Integer
        
    ' array that will hold window average results.
    Dim curArray() As Integer
    ReDim curArray(0 To arrayLength) As Integer
    
    'iterator for curArray
    Dim i As Integer
    i = 0
           
    ' iterate through entire range
    For Each c In theRange.Cells
        ' comp for window.
        counter = 0
        sum = 0
        
        ' iterate through window, sum values
        For counter = 0 To 31 '--------------------
            sum = sum + c.Offset(counter, 0).Value
        Next counter
        
        ' determine then place average into curArray
        curAverage = sum / windowSize
        curArray(i) = curAverage
        ' set up curArray index for next loop.
        i = i + 1
             
    Next c
    
    ' pad series with zeros
    For counter = i To 929
        curArray(i) = 0
    Next counter
    
    ' write out to sheet
    Range("P36", "P930").Value = Application.WorksheetFunction.Transpose(curArray) '--------------------


        
End Sub


