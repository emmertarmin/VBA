Sub backup()
    'The name of the backup file will have the name backup_originalFileName_timestamp
    
    Dim path As String
    Dim name As String
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    path = Application.ActiveWorkbook.path
    name = Application.ActiveWorkbook.name
    path = path & "\backup_" & Left(name, InStr(name, ".") - 1) & Format(Now(), "yyyyMMddhhmmss") & Right(name, Len(name) - InStr(name, ".") + 1)
    Call fso.CopyFile(Application.ActiveWorkbook.FullName, path)
    
End Sub
