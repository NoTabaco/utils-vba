Attribute VB_Name = "DateFilteringModule"
Private DeleteCount As Integer
Private StartNameCellRow As Integer

Public Sub DateFiltering()
    Dim JapanDBLastRow As Integer
    Dim InsertDateDBLastRow As Integer
    Dim InitIndex As Integer
    Dim NewestCount As Integer
    
    Const japanDB As String = "JapanDB"
    
    Dim CurrentDate As String
    CurrentDate = Format(Date, "yyyy-mm-dd")
    NewestCount = 0
    DeleteCount = 0
    StartNameCellRow = Cells.Find(What:="Name").Row
    
    JapanDBLastRow = GetDBLastRow(japanDB)
    If (Cells(JapanDBLastRow, 1).Value = CurrentDate) Then
        MsgBox "Don't use the date after the company data"
        Exit Sub
    End If
    ' A1 OK, A2 ~ OK, New Date OK
    If (Cells(1, 1).Value = CurrentDate Or Range("A1").End(xlDown).Value = CurrentDate) Then
        CurrentDateRow = FindDateRow(JapanDBLastRow, CurrentDate, NewestCount)
        InsertDateDBLastRow = DateDBLastRow(CurrentDateRow, NewestCount)
        InitIndex = FindInitIndex(JapanDBLastRow)
        FilteringData InitIndex, JapanDBLastRow, InsertDateDBLastRow
        If DeleteCount = 0 Then
            MsgBox "Data doesn't exist in " & japanDB
            Exit Sub
        End If
        Designing Val(CurrentDateRow), InsertDateDBLastRow
    Else
        MsgBox "Can't find Date"
        Exit Sub
    End If
End Sub

Function FindDateRow(JapanDBLastRow As Integer, CurrentDate As String, NewestCount As Integer) As Integer
    Dim NewestDateIndex As Integer
    For i = 1 To JapanDBLastRow
        If (Cells(i, 1).Value = CurrentDate) Then
            NewestCount = NewestCount + 1
            NewestDateIndex = i
        ElseIf (Cells(i, 1).Value = "Name") Then
            Exit For
        End If
    Next i
    ' First
    If (NewestCount = 1 And Cells(1, 1).Value = CurrentDate) Then
        FindDateRow = 1
    ' Multiple
    ElseIf (NewestCount > 1) Then
        FindDateRow = NewestDateIndex
    ' Space
    Else
        FindDateRow = Range("A1").End(xlDown).Row
    End If
End Function

Function GetDBLastRow(db As String) As Integer
    Sheets(db).Select
    GetDBLastRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
End Function

Function DateDBLastRow(ByRef CurrentDateRow As Variant, NewestCount As Integer) As Integer
    ' init or none data
    If (Range("A" & CurrentDateRow).End(xlDown).Row = StartNameCellRow And Range("A" & CurrentDateRow).End(xlDown).Value = "Name" And Not CurrentDateRow = 1) Then
        DateDBLastRow = Range("A" & CurrentDateRow - 1).End(xlDown).Row + 1
    ElseIf (CurrentDateRow = 1) Then
        If (IsEmpty(Cells(CurrentDateRow + 1, 1).Value)) Then
            DateDBLastRow = CurrentDateRow + 1
        Else: DateDBLastRow = Range("A" & CurrentDateRow).End(xlDown).Row + 1
        End If
    ' Multiple
    ElseIf (NewestCount > 1) Then
        'Multiple init
        If (IsEmpty(Cells(CurrentDateRow + 1, 1).Value)) Then
            DateDBLastRow = CurrentDateRow + 1
        Else: DateDBLastRow = Range("A" & CurrentDateRow).End(xlDown).Row + 1
        End If
    ' existing data And Only One Date
    Else
        DateDBLastRow = Range("A" & CurrentDateRow).End(xlDown).Row + 1
    End If
End Function

Function FindInitIndex(RealDBLastRow As Integer) As Integer
    For i = StartNameCellRow To RealDBLastRow
        If (Cells(i, 1).Value = "Name") Then
            FindInitIndex = i
            Exit For
        End If
    Next i
End Function

Sub FilteringData(InitIndex As Integer, JapanDBLastRow As Integer, InsertDateDBLastRow As Integer)
    For i = InitIndex To JapanDBLastRow
        If (i > InsertDateDBLastRow And Cells(i, 3).Interior.Color = 10921638 And IsEmpty(Cells(i, 3).Value) = False) Then
                Cells(i, 1).EntireRow.Copy Destination:=Cells(InsertDateDBLastRow, 1)
                Cells(i, 1).EntireRow.Delete
                DeleteCount = DeleteCount + 1
                InsertDateDBLastRow = InsertDateDBLastRow + 1
                If (IsEmpty(Cells(InsertDateDBLastRow, 3).Value) = False) Then
                    Rows(InsertDateDBLastRow).Insert shift:=xlShiftDown
                    i = InitIndex + 1
                Else: i = InitIndex
                End If
        End If
     Next i
End Sub

Sub Designing(CurrentDateRow As Integer, InsertDateDBLastRow As Integer)
    For i = CurrentDateRow To InsertDateDBLastRow - 1
        Range(Cells(i, 1), Cells(i, 3)).Interior.Color = 10284031
        Range(Cells(i, 1), Cells(i, 3)).HorizontalAlignment = xlCenter
    Next i
    Cells(999, 1).EntireRow.Copy Destination:=Cells(InsertDateDBLastRow, 1)
End Sub
