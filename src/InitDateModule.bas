Attribute VB_Name = "InitDateModule"
Sub InitDate()
Attribute InitDate.VB_ProcData.VB_Invoke_Func = "d\n14"
'
' InitDate Macro
'
' Shortcut Key: Ctrl+d
'
    Styling
End Sub

Sub Styling()
    Dim CurrentDate As String
    CurrentDate = Format(Date, "yyyy-mm-dd")
    
    CurrentRow = ActiveCell.Row
    CurrentColumn = ActiveCell.Column
    DestinationCell = CurrentColumn + 2
    Range(Cells(CurrentRow, CurrentColumn), Cells(CurrentRow, DestinationCell)).Interior.Color = 10284031
    Cells(CurrentRow, CurrentColumn).Value = CurrentDate
    Cells(CurrentRow, CurrentColumn).Font.Color = RGB(156, 101, 0)
    Cells(CurrentRow, CurrentColumn).HorizontalAlignment = xlCenter
End Sub
