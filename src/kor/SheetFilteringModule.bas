Attribute VB_Name = "SheetFilteringModule"
Private DeleteCount As Integer
Public Sub SheetFiltering(dbName As String)
    Dim dbLastRow As Integer

    dbLastRow = GetDBLastRow(dbName)
    DeleteCount = 0
    DeleteItems dbName, dbLastRow
    If DeleteCount = 0 Then
        MsgBox "Data doesn't exist in " & dbName
        Exit Sub
    End If
    Trim dbName, dbLastRow
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

Sub DeleteItems(dbName As String, dbLastRow As Integer)
    For Each Ticker In Sheets(dbName).Range("C4:C" & dbLastRow)
            If (Ticker.Interior.Color = 10921638) Then
                Ticker.EntireRow.Clear
                DeleteCount = DeleteCount + 1
            End If
    Next Ticker
End Sub

Sub Trim(db As String, TestDBLastRow As Integer)
    Dim Index As Integer
    Dim Count As Integer

    Sheets(db).Select
    Index = 4
    Count = 0
    Do While Index <= TestDBLastRow
        If (IsEmpty(Cells(Index, 1))) Then
            Cells(Index, 1).EntireRow.Delete
            Index = 4
            Count = Count + 1
                If (Count > 30) Then
                    Exit Do
                End If
        Else: Index = Index + 1
        End If
    Loop
End Sub


