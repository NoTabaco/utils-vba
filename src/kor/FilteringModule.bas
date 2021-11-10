Attribute VB_Name = "FilteringModule"
Private DeleteCount As Integer

Public Sub Filtering()
    Dim RealDBLastRow As Integer
    Dim FilterArrayLength As Integer
    
    Dim TestDBLastRow As Integer
    Const realDB As String = "RealDB"
    Const testDB As String = "TestDB"
    
    Set filterArray = CreateObject("System.Collections.ArrayList")
    
    RealDBLastRow = GetDBLastRow(realDB)
    filterArray = AddItems(RealDBLastRow).ToArray
    FilterArrayLength = UBound(filterArray) - LBound(filterArray) + 1
    If FilterArrayLength = 0 Then
        MsgBox "Data doesn't exist in " & realDB
        Exit Sub
    End If
  
    TestDBLastRow = GetDBLastRow(testDB)
    DeleteCount = 0
    DeleteItems filterArray, FilterArrayLength, TestDBLastRow
    If DeleteCount = 0 Then
        MsgBox "Data doesn't exist in " & testDB
        Exit Sub
    End If
    Trim testDB, TestDBLastRow
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

Function AddItems(RealDBLastRow As Integer) As Object
    Set Items = CreateObject("System.Collections.ArrayList")
    For Each Ticker In Sheets("RealDB").Range("C4:C" & RealDBLastRow)
        If (Ticker.Interior.Color = 10921638) Then
            Items.Add Ticker
        End If
    Next Ticker
    Set AddItems = Items
End Function

Sub DeleteItems(ByRef filterArray As Variant, FilterArrayLength As Integer, TestDBLastRow As Integer)
    For Each Ticker In Sheets("TestDB").Range("C4:C" & TestDBLastRow)
        For i = 0 To FilterArrayLength - 1
            If (Ticker = filterArray(i)) Then
                Ticker.EntireRow.Clear
                DeleteCount = DeleteCount + 1
            End If
        Next i
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
                If (Count > 10) Then
                    Exit Do
                End If
        Else: Index = Index + 1
        End If
    Loop
End Sub
