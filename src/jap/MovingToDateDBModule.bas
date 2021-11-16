Attribute VB_Name = "MovingToDateDBModule"
Sub MovingToDateDB()
    Const japanDB As String = "JapanDB"
    
    Dim CurrentDate As String
    Dim CurrentDateDB As String
    
    Dim InsertDateRow As Integer
    Dim CopyStartRow As Integer
    Dim CopyLastRow As Integer
    CurrentDate = Format(Date, "yyyy-mm-dd")
    CurrentDateDB = CurrentDate

    Sheets(CurrentDateDB).Select
    InsertDateRow = GetDBLastRow(CurrentDateDB) + 1
    If (InsertDateRow = 1) Then
        MsgBox "Please use it only if there is date data (ctrl + d)"
        Exit Sub
    End If
    
    Sheets(japanDB).Select
    CopyStartRow = FindStartRow(CurrentDate)
    CopyLastRow = FindLastRow
    ' After the company data
    If (Cells(GetDBLastRow(japanDB), 1).Value = CurrentDate) Then
        MsgBox "Don't use the date after the company data"
        Exit Sub
    End If
    ' None CurrentDate Data
    If (CopyStartRow = 0) Then
        MsgBox "Please use it only if there is date data (ctrl + d)"
        Exit Sub
    End If
    MovingFilteringData CopyStartRow, CopyLastRow, CurrentDate, InsertDateRow, CurrentDateDB
    Call Moving(False)
End Sub

Function GetDBLastRow(db As String) As Integer
    Sheets(db).Select
    On Error GoTo NotFound
    GetDBLastRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
NotFound:
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
          '  Range(Cells(i, 1), Cells(i, 3)).Interior.Color = 10921638
            Cells(i, 1).EntireRow.Copy Destination:=Sheets(CurrentDateDB).Cells(InsertDateRow, 1)
            Cells(i, 1).EntireRow.Clear
            InsertDateRow = InsertDateRow + 1
        End If
    Next i
End Sub
