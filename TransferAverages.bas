Attribute VB_Name = "TransferAverages"
' Module:  TransferAverages
'
' Author: Nicolas Perez
'
' Purpose:  This sub copies the 300-window averages of the interpolated signal by column to the specified sheets
'           in this workbook.  There are 90 values per distance for each speed, one column of values copied per cycle.
'
'

Public Sub TransferAverages()

    ' Reference cell for programmatic movement.
    Dim srcKeyCell As Range
    
    ' Iterator for distance intervals.
    Dim distanceIntervalIter As Integer
    
    ' Iterator for cycle sections.
    Dim cycleIter As Integer
    
    ' Iterator for worksheets.
    Dim sheetIter As Integer
    
    ' Beginning of range to copy.
    Dim beginRange As Range
    
    ' End of range to copy.
    Dim endRange As Range
    
    ' Total range to copy as defined by beginRange and endRange.
    Dim totalRange As Range
    
    ' Helps determine which column in the destination sheet to transfer the values.
    Dim destColumnOffset As Integer
    destColumnOffset = 0
    
    ' These numbers are changed manually depending on which set of
    '   sheets belong to the desired speed.
    For sheetIter = 2 To 2 '<----------- sheets belonging to one speed.
        
        ' Make sure source sheet is activated.
        Sheets(sheetIter).Activate
        
        ' This key cell becomes the reference cell in each cycle section.
        Set srcKeyCell = Range("H9")
        
        ' Iterates through all cycle sections in a worksheet,
        '   can be shortened for testing purposes.
        For cycleIter = 1 To 30
            
            'Keep key cell as initialized if first cycle,
            ' else go ten columns to the right of last
            ' key cell to get next key cell.
            If cycleIter <> 1 Then
                Set srcKeyCell = srcKeyCell.Offset(0, 10)
                srcKeyCell.Activate
            End If
            
            ' Get range of values to copy
            Set beginRange = srcKeyCell
            beginRange.Activate
            
            ' This emulates pressing the ctrl key and the down arrow key
            Set endRange = ActiveCell.End(xlDown)
            
            ' Copy selected range
            Set totalRange = Range(beginRange, endRange)
            totalRange.Select
            Selection.Copy
            
            ' Switch to destination sheet, paste values.
            Sheets(4).Activate '<-------- sheet that is being written to.
            
            ' Select which column to place values in destination sheet.
            Cells(3, (3 + destColumnOffset)).Select
            
            ' Paste values only
            Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
            
            ' Increment destination column for next cycle section.
            destColumnOffset = destColumnOffset + 1
            
            ' Re-activate source sheet
            Sheets(sheetIter).Activate
 
        Next cycleIter
 
    Next sheetIter

End Sub


