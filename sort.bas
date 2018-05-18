Sub sort()
    Dim i As Integer
    Dim sht As Worksheet
    Dim lastRow As Long
    
    Set sht = Worksheets("First")
    lastRow = sht.Range("A1").CurrentRegion.Rows.Count
    If lastRow < 5 Then lastRow = 5
    
    With sht.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("D3:D" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=Range("E3:E" & lastRow), Order:=xlAscending
        .SetRange Range("A5:E" & lastRow)
        .Header = xlNo
        .Apply
    End With
    
    'it makes sense to number the rows here
    For i = 1 To lastRow - 5 '
        sht.Cells(i + 5, 1).Value = i
    Next i

End Sub
