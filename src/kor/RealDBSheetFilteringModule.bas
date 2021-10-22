Attribute VB_Name = "RealDBSheetFilteringModule"
Public Sub RealDBSheetFiltering()
    Dim dbLastRow As Integer
    Const realDB As String = "RealDB"

    dbLastRow = GetDBLastRow(realDB)
    DeleteItems realDB, dbLastRow
    
    Trim realDB, dbLastRow
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
            End If
    Next Ticker
End Sub

Sub Trim(db As String, TestDBLastRow As Integer)
    Dim index As Integer
    Dim count As Integer

    Sheets(db).Select
    index = 4
    count = 0
    Do While index <= TestDBLastRow
        If (IsEmpty(Cells(index, 1))) Then
            Cells(index, 1).EntireRow.Delete
            index = 4
            count = count + 1
                If (count > 30) Then
                    Exit Do
                End If
        Else: index = index + 1
        End If
    Loop
End Sub


