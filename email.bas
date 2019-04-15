Sub send_mail()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim Message As String
    Dim d As String
    Dim DayName As String

	d = now
	'Instead of "now", you can get the right date from the worksheet:
    'd = Range("A1").Value
	
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    DayName = Application.Text(d, "[$-409]dddd")

	
	'Here is a brute force alternative for getting the DayName in the subject line
    'DayName = Choose(Weekday(d), "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
	
    Message = "Dear WHOEVER" & Chr(10) & Chr(10)
    Message = Message & "Attached you'll find the results from yesterday" & Chr(10) & Chr(10)
    Message = Message & "Kind regards," & Chr(10) & Chr(10)
    Message = Message & "YOUR NAME HERE"
    
    On Error Resume Next
    With OutMail
        .to = "EMAIL ADDRESS HERE"
        .CC = ""
        .BCC = ""
        .Subject = "Daily report (" & DayName & ") " & Year(d) & "-" & Format(Month(d), "00") & "-" & Format(Day(d), "00")
        .Body = Message
        .Attachments.Add ActiveWorkbook.FullName
        'You can add other files also like this
        '.Attachments.Add ("C:\test.txt")
        .Display   'or use .Send, if you're very confident
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
