Attribute VB_Name = "EasyMultipleGrayModule"
Sub EasyMultipleGray()
Attribute EasyMultipleGray.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' EasyMultipleGray Macro
'
' Shortcut Key: Ctrl+e
'
    For Each r In ActiveWindow.RangeSelection.Rows
        Range(Cells(r.Row, r.Column), Cells(r.Row, r.Column + 2)).Interior.Color = 10921638
    Next r
End Sub
