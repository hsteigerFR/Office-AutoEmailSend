Sub Rectangle1_Click()

'Variable definitions'
Dim surname As String
Dim name As String
Dim email As String
Dim email2 As String
Dim department As String
Dim gpa As String
Dim decision As String
Dim file As String

Dim outlookapp As Object
Dim title As String
Dim message As String
Dim x As Integer

'Index definition'
x = Sheet1.Cells(1, 12)

'As long as the table row is not empty, get row info, send email and increment row index'
Do While Sheet1.Cells(x, 3) <> ""
   
    'Create a Outlook App instance to later send an email'
    Set outlookapp = CreateObject("Outlook.Application")
    Set outlookmailitem = outlookapp.createitem(0)
    
    'Set variables according to the current row in the table'
    surname = Sheet1.Cells(x, 1)
    name = Sheet1.Cells(x, 2)
    email = Sheet1.Cells(x, 3)
    email2 = Sheet1.Cells(x, 4)
    department = Sheet1.Cells(x, 5)
    gpa = Sheet1.Cells(x, 6)
    decision = Sheet1.Cells(x, 7)
    file = ThisWorkbook.Path & "\" & Sheet1.Cells(x, 8)
    
    'Create customized message'
    title = "Semester 8 - " & department & " department's class council decision"
    message = Sheet2.Cells(1, 1) & name & " " & surname & "," & "<br/>" & "<br/>"
    message = message & Sheet2.Cells(3, 1) & department & Sheet2.Cells(3, 3) & "<br/>"
    message = message & Sheet2.Cells(4, 1) & "<b>" & decision & ".</b>" & "<br/>"
    message = message & Sheet2.Cells(5, 1) & "<b>" & gpa & " / 4." & "</b>" & "<br/>"
    message = message & Sheet2.Cells(6, 1) & "<br/>" & "<br/>"
    message = message & Sheet2.Cells(8, 1) & "<br/>" & Sheet2.Cells(9, 1)
    
    'Create mail object thanks to cuurent row info'
    outlookmailitem.To = email
    outlookmailitem.cc = email2
    outlookmailitem.bcc = ""
    outlookmailitem.Subject = title
    With outlookmailitem
        .Attachments.Add file, 1, 0
    End With
    outlookmailitem.HTMLBody = message
       
    'Send email and get to next row'
    'outlookmailitem.display'
    outlookmailitem.send
    Sheet1.Cells(x, 9) = "OK"
    email = ""
    x = x + 1
        
    'Waits a little bit before processing the next row : it prevents any Outlook crash'
    Application.Wait (Now + TimeValue("0:00:02"))

Loop

'Set outlookapp = Nothing'
'Set outlookmailitem = Nothing'

End Sub
