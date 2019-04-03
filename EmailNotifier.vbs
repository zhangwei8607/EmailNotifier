dim filesys, filetxt, cellLine, array, t1, isLongAssay
Set filesys = CreateObject("Scripting.FileSystemObject")
Set xlApp = CreateObject("Excel.Application")
Set MyEmail = CreateObject("CDO.Message")
Set emailConfig = MyEmail.Configuration
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "relay.smtpserver.net"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")    = 2  
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")      = false 
emailConfig.Fields.Update

current = filesys.GetParentFolderName(WScript.ScriptFullName)
set m1 = xlApp.Workbooks.Open(current & "\Mapping.xlsx")
set m1s = m1.Sheets("mapping")
MyEmail.From = Trim(m1s.Range("A2").Value)
folderName = Trim(m1s.Range("B2").Value)
notifyDays = CInt(Trim(m1s.Range("C2").Value))

index = 4
fileName = Trim(m1s.Range("A4").Value)
sheetName = Trim(m1s.Range("B4").Value)
contentCell = Trim(m1s.Range("C4").Value)
dateCell = Trim(m1s.Range("D4").Value)
emailCell = Trim(m1s.Range("E4").Value)
checkCell = Trim(m1s.Range("F4").Value)
postPoneCell = Trim(m1s.Range("G4").Value)
startCell = Trim(m1s.Range("H4").Value)

Do While Not fileName = ""

    index = index + 1
    set f1 = xlApp.Workbooks.Open(folderName & fileName)
    set f1s = f1.Sheets(sheetName)
    start = CInt(startCell)
    content = Trim(f1s.Range(contentCell & startCell).Value)
    timeLine = Trim(f1s.Range(dateCell & startCell).Value)
    email = Trim(f1s.Range(emailCell & startCell).Value)
    check = Trim(f1s.Range(checkCell & startCell).Value)
    postPone = Trim(f1s.Range(postPoneCell & startCell).Value)
    workGroup = Trim(f1s.Range("B1").Value)

    Do While email <> "" 
        start = start + 1
        If timeLine <> "" Then
            If (check = "Y" Or check = "") And CDate(timeLine) - Date <= notifyDays And CDate(timeLine) - Date >= 0 Then
                MyEmail.Subject = "Email Reminder"
                MyEmail.To = email
                MyEmail.HTMLBody= "<h4>This is the content for email reminder</h4>"
                MyEmail.Send
            End If
        End if
        If  postPone <> "" Then
            If CDate(postPone) - Date <= notifyDays And CDate(postPone) - Date >= 0 Then
                MyEmail.Subject = "Email Reminder" 
                MyEmail.To = email
                MyEmail.HTMLBody= "<h4>This is the content for email reminder</h4>"
                MyEmail.Send
            End If
        End If
        content = Trim(f1s.Range(contentCell & Cstr(start)).Value)
        timeLine = Trim(f1s.Range(dateCell & Cstr(start)).Value)
        email = Trim(f1s.Range(emailCell & Cstr(start)).Value)
        check = Trim(f1s.Range(checkCell & Cstr(start)).Value)
        postPone = Trim(f1s.Range(postPoneCell & Cstr(start)).Value)
    Loop

    fileName = Trim(m1s.Range("A" & Cstr(index)).Value)
    sheetName = Trim(m1s.Range("B" & Cstr(index)).Value)
    contentCell = Trim(m1s.Range("C" & Cstr(index)).Value)
    dateCell = Trim(m1s.Range("D" & Cstr(index)).Value) 
    emailCell = Trim(m1s.Range("E" & Cstr(index)).Value)
    checkCell = Trim(m1s.Range("F" & Cstr(index)).Value)
    postPoneCell = Trim(m1s.Range("G" & Cstr(index)).Value)
    startCell = Trim(m1s.Range("H" & Cstr(index)).Value)

    f1.Close TRUE
Loop

Set MyEmail = nothing
m1.Close TRUE
MsgBox "Check finished and sent email notifications."