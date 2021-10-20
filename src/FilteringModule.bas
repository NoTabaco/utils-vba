Attribute VB_Name = "FilteringModule"
Public Sub Filtering()
    Dim RealDBLastRow As Integer
    Dim FilterArrayLength As Integer
    
    Dim TestDBLastRow As Integer
    Dim index As Integer
    Dim count As Integer
    
    Set filterArray = CreateObject("System.Collections.ArrayList")
    
    RealDBLastRow = GetDBLastRow("RealDB")
    filterArray = AddItems(RealDBLastRow).ToArray
    FilterArrayLength = UBound(filterArray) - LBound(filterArray) + 1
  
    TestDBLastRow = GetDBLastRow("TestDB")
    DeleteItems filterArray, FilterArrayLength, TestDBLastRow
    
    index = 4
    count = 0
    Trim "TestDB", index, count, TestDBLastRow
  
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
            End If
        Next i
    Next Ticker
End Sub

Sub Trim(db As String, index As Integer, count As Integer, TestDBLastRow As Integer)
    Sheets(db).Select
    Do While index <= TestDBLastRow
        If (IsEmpty(Cells(index, 1))) Then
            Cells(index, 1).EntireRow.Delete
            index = 4
            count = count + 1
                If (count > 10) Then
                    Exit Do
                End If
        Else: index = index + 1
        End If
    Loop
End Sub
