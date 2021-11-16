Attribute VB_Name = "MovingToDateDBModule"
Sub MovingToDateDB()
    Const japanDB As String = "JapanDB"
    
    Dim CurrentDate As String
    CurrentDate = Format(Date, "yyyy-mm-dd")
    Dim CurrentDateDB As String
    CurrentDateDB = CurrentDate
    
    Sheets(CurrentDateDB).Select
    Dim InsertDateRow As Integer
    InsertDateRow = GetDBLastRow(CurrentDateDB) + 1
    
    Sheets(japanDB).Select
    Dim CopyStartRow As Integer
    Dim CopyLastRow As Integer
    CopyStartRow = FindStartRow(CurrentDate)
    CopyLastRow = FindLastRow
    MovingFilteringData CopyStartRow, CopyLastRow, CurrentDate, InsertDateRow, CurrentDateDB
    Call Moving
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

' Date
Function FindStartRow(CurrentDate As String)
    If (Cells(1, 1).Value = CurrentDate) Then
        FindStartRow = 1
    ElseIf (Range("A1").End(xlDown).Value = CurrentDate) Then
        FindStartRow = Range("A1").End(xlDown).Row
    End If
End Function

' Name - 1
Function FindLastRow()
   FindLastRow = Cells.Find(What:="Name").Row - 1
End Function

Sub MovingFilteringData(CopyStartRow As Integer, CopyLastRow As Integer, CurrentDate As String, InsertDateRow As Integer, CurrentDateDB As String)
    For i = CopyStartRow To CopyLastRow
        If (Cells(i, 1).Value = CurrentDate) Then
            Cells(i, 1).EntireRow.Clear
        End If
        If (Not (IsEmpty(Cells(i, 1).Value))) Then
            Cells(i, 1).EntireRow.Copy Destination:=Sheets(CurrentDateDB).Cells(InsertDateRow, 1)
            Cells(i, 1).EntireRow.Clear
            InsertDateRow = InsertDateRow + 1
        End If
    Next i
End Sub
