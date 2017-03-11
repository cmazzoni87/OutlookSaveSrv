Option Explicit
Function filterBy(source As Range, target As Variant, Optional skipHeaders As Boolean = True, Optional secondTarget As Variant, _
Optional greaterlessThan As String, Optional shtTarget As String, _
Optional gretlessEquals As Boolean, Optional filterFor As Boolean) As Long
''Filter by or filter for target, will re dimennsion passed range accordingly Returns new count of cells
Dim iterator As Range, counter As Integer, i As Long, Ws As Worksheet, itr As Long, trigger As Boolean

    If shtTarget = "" Then Set Ws = ThisWorkbook.ActiveSheet
    If shtTarget <> "" Then Set Ws = ThisWorkbook.Worksheets(shtTarget)
    If greaterlessThan <> "<" And greaterlessThan <> "<" And greaterlessThan <> "=" Then GoTo ErrHand
    counter = source.Column
    With Ws
        Set iterator = .Cells(.Rows.Count, counter).End(xlUp)
        i = iterator.Row
        If skipHeaders = True Then itr = 2
        If skipHeaders <> True Then itr = 1
        For itr = itr To i
            If greaterlessThan = "" Then
                If .Cells(itr, counter).Value <> target Then
                    .Cells(itr, counter).Value = ""
                    trigger = True
                End If
            End If
            If greaterlessThan = ">" Then
                If .Cells(itr, counter).Value < target Then
                    .Cells(itr, counter).Value = ""
                    trigger = True
                End If
            End If
            If greaterlessThan = "<" Then
                If .Cells(itr, counter).Value > target Then
                    .Cells(itr, counter).Value = ""
                    trigger = True
                End If
            End If
        Next
        If trigger = True Then Range(.Cells(source.Row, source.Column), .Cells(iterator.Row, source.Column)) _
        .SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    End With
    filterBy = iterator.Row

ErrHand:
    MsgBox "Not a valid Character please re try"
    Exit Function

End Function
