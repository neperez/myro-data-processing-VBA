Attribute VB_Name = "FixDistances"
Public Sub fixDistances()

    Dim sheetIter As Integer
    Dim cycleIter As Integer
    Dim curKeyCell As Range
    Dim beginRange As Range
    Dim endRange As Range
    Dim totalRange As Range
    Dim num As Double
    
    For sheetIter = 1 To 2
         
         Sheets(sheetIter).Activate
         Set curKeyCell = Range("F10")
         curKeyCell.Activate
        
        
        For cycleIter = 1 To 30
                
            
            If cycleIter <> 1 Then
                Set curKeyCell = curKeyCell.Offset(0, 10)
                curKeyCell.Activate
            End If
            
            Set beginRange = curKeyCell
            Set endRange = curKeyCell.Offset(1000, 0)
            Set totalRange = Range(beginRange, endRange)
            totalRange.Clear
            
            
            
                        
                       
                
        
        Next cycleIter
    
    
    
    Next sheetIter
    




End Sub
