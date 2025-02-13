Option Explicit

Private Sub btnLoadClients_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range

    ' Set worksheet
    Set ws = Sheets("Data")
    
    ' Find last row in Client Code column
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set rng = ws.Range("A2:B" & lastRow) ' Load both Client Code & Attachment Path

    ' Clear previous items
    lstClients.Clear

    ' Load client codes & attachments into ListBox
    For Each cell In rng.Columns(1).Cells
        If cell.Value <> "" Then
            lstClients.AddItem cell.Value
            lstClients.List(lstClients.ListCount - 1, 1) = cell.Offset(0, 1).Value ' Add attachment path
        End If
    Next cell

    ' Show message if no client codes found
    If lstClients.ListCount = 0 Then
        MsgBox "No client codes found!", vbExclamation
    End If
End Sub

Private Sub btnGenerateEmails_Click()
    Dim objOutlook As Object
    Dim objMail As Object
    Dim i As Integer
    Dim emailSubject As String
    Dim emailBody As String
    Dim attachPath As String
    
    ' Ensure subject and body are filled
    If Trim(txtSubject.Value) = "" Or Trim(txtBody.Value) = "" Then
        MsgBox "Please enter the email subject and body.", vbExclamation
        Exit Sub
    End If

    ' Ensure there are client codes
    If lstClients.ListCount = 0 Then
        MsgBox "No client codes loaded!", vbExclamation
        Exit Sub
    End If

    ' Initialize Outlook
    Set objOutlook = CreateObject("Outlook.Application")

    ' Loop through client codes and generate emails
    For i = 0 To lstClients.ListCount - 1
        Set objMail = objOutlook.CreateItem(0)
        
        emailSubject = txtSubject.Value & " - " & lstClients.List(i)
        emailBody = txtBody.Value
        attachPath = lstClients.List(i, 1) ' Get attachment path

        objMail.To = "recipient@example.com"  ' Change to dynamic if needed
        objMail.Subject = emailSubject
        objMail.Body = emailBody

        ' Attach file if exists
        If attachPath <> "" And Dir(attachPath) <> "" Then
            objMail.Attachments.Add attachPath
        End If

        objMail.Close 0 ' Save as draft

        Set objMail = Nothing
    Next i

    ' Cleanup
    Set objOutlook = Nothing
    MsgBox "Emails saved as drafts!", vbInformation
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
