Attribute VB_Name = "InterpolateNoAverage"
Option Explicit

' Module: InterpolateNoAverage
'
' Author: Nicolas Perez
'
' Purpose: Performs linear interpolation with respect to a standarized scale.
'

Public Sub InterpolateNoAverage()
    
    ' Set range of standarized distance column
    Dim theRange As Range
    Set theRange = Range("K8:K558") '<-------Enter range of std. distance column.
    
    ' Set range of unprocessed signals.
    Dim uArrayRange As Range
    Set uArrayRange = Range("B7:B30") '<----------Enter range of unprocessed signals.
       
    ' For length of UDistance and USignal arrays
    Dim uArrayLength As Integer
    uArrayLength = Application.WorksheetFunction.CountA(uArrayRange)
       
    ' Get ranges of unprocessed data:
    Dim uSigRange As Range
    Set uSigRange = uArrayRange
    Dim uDistanceRange As Range
    Set uDistanceRange = Range("D7:D30") '--------------------
    
    ' The range values will be written to.
    Dim rangeToWrite As Range
    Set rangeToWrite = Range("J8", "J1020") '-----------------------
         
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
        
                 
            ' Perform linear interplolation on Signals:
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




