Sub clearSheets()
    Dim sht As Worksheet
    Dim lastRow As Long
    
    'Instead of "First" and "Second", you have to name whatever sheets you want to clear
    Set sht = Worksheets("First")
    'We look at how many rows there is
    lastRow = sht.Range("A1").CurrentRegion.Rows.Count
    'If there's rows at all that we want to clear (rows below 5 in this example), we delete their content
    If lastRow > 5 Then sht.Range("A5:E" & lastRow).EntireRow.Delete

    Set sht = Worksheets("Second")
    lastRow = sht.Range("A1").CurrentRegion.Rows.Count
    If lastRow > 5 Then sht.Range("A5:E" & lastRow).EntireRow.Delete

End Sub
