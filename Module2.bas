Attribute VB_Name = "Module2"
Option Explicit


Public Sub interpolateNoAverage()
    
    ' set range of standarized time column
    Dim theRange As Range
    Set theRange = Range("N36:N188") '--------------------
       
    ' iterator for cells in processed time series
    Dim t As Range
    
    ' for length of Signal array
    Dim signalArrayLength As Integer
    
    ' for length of UTime and USignal arrays
    Dim uArrayLength As Integer
    uArrayLength = Application.WorksheetFunction.CountA(Range("B5:B28")) '--------------------
        
    ' get length of entire standardized time range
    Dim rangeLength As Integer
    rangeLength = Application.WorksheetFunction.CountA(theRange)
    
    ' signal array Length is equal to length of standardized time range
    signalArrayLength = rangeLength
    
    ' initialized arrays for unprocessed signal and time values.
    Dim UTimeArray() As Double
    Dim USignalArray() As Integer
    
    ' must redimension the arrays,
    ' cannot use non-constant for length parameter in initial dimensioning
    ReDim UTimeArray(0 To uArrayLength - 1) As Double
    ReDim USignalArray(0 To uArrayLength - 1) As Integer
            
    ' iterators for range
    Dim signal As Range
    Dim time As Range
        
    ' iterator for array population:
    Dim x As Integer
    x = 0
    Dim y As Integer
    y = 0
    Dim jStart As Integer
    jStart = 0
     
    ' get ranges of unprocessed data:
    Dim uSigRange As Range
    Set uSigRange = Range("B5:B28") '--------------------
    Dim uTimeRange As Range
    Set uTimeRange = Range("C5:C28") '--------------------
    
    ' fill signal array:
    For Each signal In uSigRange
        USignalArray(x) = signal.Value
        x = x + 1
    Next signal
    
    ' fill time array
    For Each time In uTimeRange
        UTimeArray(y) = time.Value
        y = y + 1
    Next time
       
    ' array that will hold signal interpolation results.
    Dim signalArray() As Integer
    ReDim signalArray(0 To signalArrayLength - 1) As Integer
    
    Dim i As Integer
    i = 0
    Dim j As Integer
    
    ' holds part of calculation
    Dim valueFromProportion As Double
    valueFromProportion = 0
               
    ' iterate through standarized time column
    For Each t In theRange.Cells
                    
        ' for computation
        j = jStart
        
        'this test is repeated with different consequences below (REMOVE COMMENT)
        If UTimeArray(j) > t.Value Then
            signalArray(i) = USignalArray(j)
        Else
        
            ' get index of upper bound (tsubh)
            Do While UTimeArray(j) < t.Value And j < uArrayLength - 1
                j = j + 1
            Loop
        
                 
            ' linear interplolation:
            If UTimeArray(j) < t.Value Then
                signalArray(i) = USignalArray(j)
            ElseIf UTimeArray(j) = t.Value Then
                signalArray(i) = USignalArray(j)
            ElseIf (USignalArray(j - 1) > USignalArray(j)) Then
                valueFromProportion = ((t.Value - UTimeArray(j - 1)) * (USignalArray(j - 1) - USignalArray(j))) / (UTimeArray(j) - UTimeArray(j - 1))
                signalArray(i) = USignalArray(j - 1) - valueFromProportion
            Else
                valueFromProportion = ((t.Value - UTimeArray(j - 1)) * (USignalArray(j) - USignalArray(j - 1))) / (UTimeArray(j) - UTimeArray(j - 1))
                signalArray(i) = USignalArray(j - 1) + valueFromProportion
            End If
        
        End If
       
           
        i = i + 1
             
    Next t
       
    ' write out values next to processed time column.
    Range("M36", "M438").Value = Application.WorksheetFunction.Transpose(signalArray) '--------------------


        
End Sub



