Attribute VB_Name = "CreateDuplicateTickerModule"
Sub CreateDuplicateTicker()
    Dim RLDBLastRow As Integer
    Dim TestDBLastRow As Integer
    
    Dim SelectTestDBRow As String
    Dim FindRLDBRow As String
    
    Const rlDB As String = "RL_DB"
    Const testDB As String = "TestDB"
    
    RLDBLastRow = GetDBLastRow(rlDB)
    TestDBLastRow = GetDBLastRow(testDB)
    
    CreateTicker RLDBLastRow, TestDBLastRow, rlDB, testDB
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

Sub CreateTicker(RLDBLastRow As Integer, TestDBLastRow As Integer, rlDB As String, testDB As String)
    For i = 4 To TestDBLastRow
        Sheets(testDB).Select
        SelectTestDBRow = Cells(i, 1)
            For j = 4 To RLDBLastRow
                Sheets(rlDB).Select
                FindRLDBRow = Cells(j, 1)
                If (SelectTestDBRow = FindRLDBRow) Then
                    FindRLDBRow = Cells(j, 2)
                    Sheets(testDB).Select
                    Cells(i, 3).NumberFormat = "@"
                    Cells(i, 3).HorizontalAlignment = xlCenter
                    Cells(i, 3) = FindRLDBRow
                    Exit For
                End If
            Next j
        Next i
End Sub
