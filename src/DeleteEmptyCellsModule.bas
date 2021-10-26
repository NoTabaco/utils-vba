Attribute VB_Name = "DeleteEmptyCellsModule"
Sub DeleteEmptyCells()
Attribute DeleteEmptyCells.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' DeleteEmptyCells Macro
'
' Shortcut Key: Ctrl+t
'
    Dim Index As Integer
    Dim Count As Integer
    Dim PrevPoint As Integer

    Index = 1
    
    Do While (1)
        If (Cells(Index, 1).Interior.Color = 10921638 And IsEmpty(Cells(Index, 1))) Then
            Cells(Index, 1).EntireRow.Delete
            PrevPoint = Index
            Index = PrevPoint
            Count = 0
        Else:
            Index = Index + 1
            Count = Count + 1
        End If
        
        If (Count > 100) Then
            Exit Do
        End If
       
    Loop
End Sub



