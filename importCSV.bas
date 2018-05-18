Sub openCSV(filepath As String)
    Dim line As String
    Dim arrayOfElements
    Dim linenumber As Integer
    Dim elementnumber As Integer
    Dim element As Variant
    
    linenumber = 0
    elementnumber = 0
    
    Open filepath For Input As #1
        Do While Not EOF(1) 'loop until end of file
            linenumber = linenumber + 1
            Line Input #1, line
            arrayOfElements = Split(line, ";")
            
            elementnumber = 0
            For Each element In arrayOfElements
                elementnumber = elementnumber + 1
                If IsNumeric(element) Then
                    Sheets("CSV").Cells(linenumber, elementnumber).Value = CDbl(element)
                    Else
                    Sheets("CSV").Cells(linenumber, elementnumber).Value = element
                End If
            Next
        Loop
    Close #1 'Close file
        
End Sub
