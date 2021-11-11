Attribute VB_Name = "MovingModule"
Sub Moving()
    Dim JapanDBLastRow As Integer
    Dim JapanDBFirstRow As Integer

    
    Const japanDB As String = "JapanDB"
    
    Dim CurrentDate As String
    CurrentDate = Format(Date, "yyyy-mm-dd")
    
    If (Range("A1").End(xlDown).Value = CurrentDate) Then
        MsgBox "Please use it only if there is no date data"
        Exit Sub
    End If
    
    If (Range("A1").End(xlDown).Row = 10) Then
        MsgBox "No need to control"
        Exit Sub
    End If
    
    JapanDBLastRow = GetDBLastRow(japanDB)
    JapanDBFirstRow = Range("A1").End(xlDown).Row
    Move JapanDBFirstRow, JapanDBLastRow
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

Sub Move(JapanDBFirstRow As Integer, JapanDBLastRow As Integer)
    Dim InsertIndex As Integer
    InsertIndex = 10

    For i = JapanDBFirstRow To JapanDBLastRow
        Cells(JapanDBFirstRow, 1).EntireRow.Copy Destination:=Cells(InsertIndex, 1)
        Cells(JapanDBFirstRow, 1).EntireRow.Clear
        JapanDBFirstRow = JapanDBFirstRow + 1
        InsertIndex = InsertIndex + 1
    Next i
End Sub
