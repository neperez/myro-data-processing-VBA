Attribute VB_Name = "GetLongestDistance"
Option Explicit

' Module:  GetLongestDistance
' Author:  Nicolas Perez
' Purpose: Gets longest distance travelled by the robot for a specified speed.
'           For setup of the linear distance column on the averaging worksheets.

Public Sub GetLongestDistance()
    
    ' Slightly speed up execution
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Iterator for sheets
    Dim sheetIter As Integer
    
    ' Iterator for cycle sections
    Dim cycleIter As Integer
    
    ' Holds current largest distance for a speed.
    Dim largestDistance As Double
    largestDistance = 0
    
    ' Distance being tested.
    Dim curDistance As Double
    curDistance = 0
    Dim curKeyCell As Range
    
    For sheetIter = 2 To 2 '<----------------Enter sheets to check
            
            ' Activate current sheet
            Sheets(sheetIter).Activate
            
            ' Set key cell to H7 at the beginning of a new sheet
            Set curKeyCell = Range("G3")
            curKeyCell.Activate
            
            ' All references will be relatively the same from each key cell in all cycles
            For cycleIter = 1 To 30
                
                'Keep key cell as A1 if first cycle,
                ' else go ten columns to the right of last
                ' key cell to get next key cell.
                If cycleIter <> 1 Then
                    Set curKeyCell = curKeyCell.Offset(0, 10)
                    curKeyCell.Activate
                End If
                
                curDistance = curKeyCell.Value
                
                ' Compare distances; if larger, replace value
                If curDistance > largestDistance Then
                    largestDistance = curDistance
                End If
                             
            Next cycleIter
        
        Next sheetIter
        
        ' Display largest value.
        MsgBox "Largest value is " & largestDistance
        
        ' Restore application defaults.
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True

End Sub





