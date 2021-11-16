Attribute VB_Name = "MovingModule"
Sub Moving(isRaiseError As Boolean)
    Dim JapanDBLastRow As Integer
    Dim JapanDBFirstRow As Integer

    Const japanDB As String = "JapanDB"
    
    Dim CurrentDate As String
    CurrentDate = Format(Date, "yyyy-mm-dd")

    If (Range("A1").End(xlDown).Value = CurrentDate) Then
        MsgBox "Please use it only if there is no date data"
        Exit Sub
    End If
    
    If (Range("A1").End(xlDown).Row = 10 And Cells(10, 1).Value = "Name") Then
        If (isRaiseError) Then
          MsgBox "No need to control"
          Exit Sub
        Else
          Exit Sub
        End If
    End If
    
    JapanDBLastRow = GetDBLastRow(japanDB)
    If (Cells(JapanDBLastRow, 1).Value = CurrentDate) Then
        MsgBox "Please use it only if there is no date data"
        Exit Sub
    End If
    JapanDBFirstRow = Range("A1").End(xlDown).Row
    If (Cells(1, 1).Value = "Name") Then
        JapanDBFirstRow = 1
    End If
    Move JapanDBFirstRow, JapanDBLastRow, japanDB
End Sub

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

Sub Move(JapanDBFirstRow As Integer, JapanDBLastRow As Integer, japanDB As String)
    Dim InsertIndex As Integer
    InsertIndex = 10
    ' Top - Down
    If ((JapanDBFirstRow < InsertIndex) And Not (IsEmpty(Cells(InsertIndex, 1).Value))) Then
        Range(Cells(JapanDBFirstRow, 1), Cells(JapanDBLastRow, 3)).Copy Cells(InsertIndex, 1)
        For i = JapanDBFirstRow + 1 To JapanDBLastRow
        If (Cells(i, 1).Value = "Name") Then
            Cells(JapanDBFirstRow, 1).EntireRow.Clear
            Exit Sub
        End If
        Cells(i, 1).EntireRow.Clear
        Next i
        Exit Sub
    End If
    ' Bottom - Up
    For i = JapanDBFirstRow To JapanDBLastRow
        Cells(i, 1).EntireRow.Copy Destination:=Sheets(japanDB).Cells(InsertIndex, 1)
        Cells(i, 1).EntireRow.Clear
        InsertIndex = InsertIndex + 1
    Next i
End Sub
