Attribute VB_Name = "PaintingGrayModule"
Sub PaintingGray()
Attribute PaintingGray.VB_ProcData.VB_Invoke_Func = "g\n14"
'
' PaintingGray Macro
'
' Shortcut Key: Ctrl+g
'
    Dim CurrentRow As Integer
    Dim CurrentColumn As Integer
    
    Dim DestinationCell As Integer
    
    CurrentRow = ActiveCell.Row
    CurrentColumn = ActiveCell.Column
    DestinationCell = ActiveCell.Column + 2
    Range(Cells(CurrentRow, CurrentColumn), Cells(CurrentRow, DestinationCell)).Interior.Color = 10921638
End Sub
