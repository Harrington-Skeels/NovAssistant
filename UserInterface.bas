Attribute VB_Name = "UserInterface"
Function InitializeOutlook() As Boolean
    ' Initialize connection to Outlook
    On Error Resume Next
    
    Dim objOutlook As Object
    
    ' Try to get existing Outlook application
    Set objOutlook = GetObject(, "Outlook.Application")
    
    ' Create new instance if needed
    If objOutlook Is Nothing Then
        Set objOutlook = CreateObject("Outlook.Application")
    End If
    
    ' Check if successful
    If objOutlook Is Nothing Then
        InitializeOutlook = False
    Else
        InitializeOutlook = True
        Set objOutlook = Nothing
    End If
End Function

Sub SyncWithOutlook()
    ' Sync customer emails, calendar, and tasks with Outlook
    If Not InitializeOutlook() Then
        MsgBox "Could not connect to Outlook. Please ensure Outlook is running.", vbExclamation
        Exit Sub
    End If
    
    ' Sync emails
    SyncCustomerEmails
    
    ' Sync calendar items
    SyncOutlookCalendar
    
    ' Refresh dashboard
    RefreshDashboard
    
    ' Show confirmation
    MsgBox "Sync with Outlook completed successfully.", vbInformation
End Sub

Sub SyncCustomerEmails()
    ' Sync emails related to customers
    On Error Resume Next
    
    Dim objOutlook As Object
    Dim objNamespace As Object
    Dim objInbox As Object
    Dim objItems As Object
    Dim objMail As Object
    Dim customerSheet As Worksheet
    Dim customerEmails As Collection
    Dim i As Long
    Dim syncPeriod As Date
    Dim filterString As String
    Dim emailsProcessed As Integer
    
    ' Get customer sheet
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    
    ' Create collection of customer emails
    Set customerEmails = New Collection
    
    ' Build list of customer emails (column C in CustomerTracker)
    For i = 2 To customerSheet.UsedRange.Rows.count
        If Not IsEmpty(customerSheet.Cells(i, 3).Value) Then
            On Error Resume Next
            customerEmails.Add customerSheet.Cells(i, 3).Value, CStr(customerSheet.Cells(i, 3).Value)
            On Error GoTo 0
        End If
    Next i
    
    ' Set sync period to last 24 hours
    syncPeriod = Now - 1
    
    ' Get Outlook objects
    Set objOutlook = GetObject(, "Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objInbox = objNamespace.GetDefaultFolder(6) ' 6 = olFolderInbox
    
    ' Filter for emails in the last 24 hours
    filterString = "[ReceivedTime] >= '" & Format(syncPeriod, "mm/dd/yyyy hh:mm AMPM") & "'"
    Set objItems = objInbox.Items.Restrict(filterString)
    
    ' Init counter
    emailsProcessed = 0
    
    ' Process emails
    For Each objMail In objItems
        ' Check if email is from a customer
        On Error Resume Next
        Dim senderEmail As String
        senderEmail = objMail.SenderEmailAddress
        
        Dim customerName As String
        customerName = ""
        
        ' Find matching customer by email
        For i = 2 To customerSheet.UsedRange.Rows.count
            If LCase(customerSheet.Cells(i, 3).Value) = LCase(senderEmail) Then
                customerName = customerSheet.Cells(i, 2).Value
                Exit For
            End If
        Next i
        
        ' Log email if from customer
        If customerName <> "" Then
            AddContactHistoryRecord customerName, "Email Received", objMail.subject, objMail.ReceivedTime
            emailsProcessed = emailsProcessed + 1
            
            ' Update last contact date
            Dim customerRow As Range
            Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
            If Not customerRow Is Nothing Then
                customerRow.Offset(0, 4).Value = objMail.ReceivedTime
            End If
        End If
    Next objMail
    
    ' Clean up
    Set objItems = Nothing
    Set objInbox = Nothing
    Set objNamespace = Nothing
    Set objOutlook = Nothing
    
    ' Show status
    If emailsProcessed > 0 Then
        Application.StatusBar = "Processed " & emailsProcessed & " customer emails"
        Application.OnTime Now + TimeValue("00:00:10"), "ResetStatusBar"
    End If
End Sub

Sub SyncOutlookCalendar()
    ' Sync calendar appointments for today
    On Error Resume Next
    
    Dim objOutlook As Object
    Dim objNamespace As Object
    Dim objCalendar As Object
    Dim objItems As Object
    Dim objAppt As Object
    Dim callPlannerSheet As Worksheet
    Dim customerSheet As Worksheet
    Dim apptDate As Date
    Dim filterString As String
    Dim apptsProcessed As Integer
    
    ' Get sheets
    Set callPlannerSheet = ThisWorkbook.Sheets("CallPlanner")
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    
    ' Today's date
    apptDate = Date
    
    ' Get Outlook objects
    Set objOutlook = GetObject(, "Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objCalendar = objNamespace.GetDefaultFolder(9) ' 9 = olFolderCalendar
    
    ' Filter for today's appointments
    filterString = "[Start] >= '" & Format(apptDate, "mm/dd/yyyy") & " 12:00 AM' AND [End] <= '" & Format(apptDate, "mm/dd/yyyy") & " 11:59 PM'"
    Set objItems = objCalendar.Items.Restrict(filterString)
    
    ' Init counter
    apptsProcessed = 0
    
    ' Process appointments
    For Each objAppt In objItems
        ' Check if appointment is a customer follow-up
        If InStr(1, objAppt.subject, "Follow-up", vbTextCompare) > 0 Then
            ' Extract customer name from subject
            Dim customerName As String
            Dim dashPos As Integer
            
            dashPos = InStr(objAppt.subject, "-")
            If dashPos > 0 Then
                customerName = Trim(Mid(objAppt.subject, dashPos + 1))
            End If
            
            ' Find customer in tracker
            If customerName <> "" Then
                Dim customerRow As Range
                Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
                
                ' Add to call planner if customer exists
                If Not customerRow Is Nothing Then
                    ' Check if already in call planner
                    Dim existingCall As Range
                    Set existingCall = callPlannerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
                    
                    If existingCall Is Nothing Then
                        ' Add to call planner
                        Dim nextRow As Long
                        nextRow = callPlannerSheet.Cells(callPlannerSheet.Rows.count, "A").End(xlUp).row + 1
                        
                        callPlannerSheet.Cells(nextRow, 1).Value = Format(objAppt.Start, "h:mm AM/PM")
                        callPlannerSheet.Cells(nextRow, 2).Value = customerName
                        callPlannerSheet.Cells(nextRow, 3).Value = customerRow.Offset(0, 2).Value ' Phone
                        
                        ' Extract purpose from subject
                        If InStr(objAppt.subject, " - ") > 0 Then
                            callPlannerSheet.Cells(nextRow, 4).Value = left(objAppt.subject, InStr(objAppt.subject, " - ") - 1)
                        Else
                            callPlannerSheet.Cells(nextRow, 4).Value = "Follow-up"
                        End If
                        
                        callPlannerSheet.Cells(nextRow, 5).Value = customerRow.Offset(0, 3).Value ' Stage
                        callPlannerSheet.Cells(nextRow, 6).Value = customerRow.Offset(0, 12).Value ' Status
                        callPlannerSheet.Cells(nextRow, 7).Value = "Pending"
                        
                        apptsProcessed = apptsProcessed + 1
                    End If
                End If
            End If
        End If
    Next objAppt
    
    ' Clean up
    Set objItems = Nothing
    Set objCalendar = Nothing
    Set objNamespace = Nothing
    Set objOutlook = Nothing
    
    ' Show status
    If apptsProcessed > 0 Then
        Application.StatusBar = "Added " & apptsProcessed & " appointments to call planner"
        Application.OnTime Now + TimeValue("00:00:10"), "ResetStatusBar"
    End If
End Sub

Sub SendEmailToCustomer()
    ' Send an email to the selected customer
    Dim customerSheet As Worksheet
    Dim selectedCell As Range
    Dim customerRow As Range
    Dim customerName As String
    Dim customerEmail As String
    
    ' Get selection
    Set selectedCell = selection
    
    ' Check if we're in the customer tracker
    If selectedCell.Parent.Name <> "CustomerTracker" Then
        MsgBox "Please select a customer in the CustomerTracker sheet.", vbExclamation
        Exit Sub
    End If
    
    ' Get customer sheet
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    
    ' Find the customer row
    Set customerRow = customerSheet.Rows(selectedCell.row)
    
    ' Get customer details
    customerName = customerSheet.Cells(selectedCell.row, 2).Value ' Column B
    customerEmail = customerSheet.Cells(selectedCell.row, 3).Value ' Column C
    
    ' Check if we have valid customer
    If customerName = "" Or customerEmail = "" Then
        MsgBox "Please select a valid customer with an email address.", vbExclamation
        Exit Sub
    End If
    
    ' Show template selector
    Dim templateID As String
    templateID = ShowEmailTemplateSelector()
    
    If templateID <> "" Then
        ' Send email with template
        SendTemplateEmail customerName, customerEmail, templateID
    End If
End Sub

Function ShowEmailTemplateSelector() As String
    ' Show a simple template selector dialog and return the selected template ID
    Dim templatesSheet As Worksheet
    Dim templateList As String
    Dim i As Long
    Dim templateID As String
    
    ' Get templates sheet
    Set templatesSheet = ThisWorkbook.Sheets("Templates")
    
    ' Build list of templates
    templateList = ""
    For i = 2 To templatesSheet.UsedRange.Rows.count
        If templatesSheet.Cells(i, 1).Value = "EmailTemplate" Then
            If templateList <> "" Then templateList = templateList & vbCrLf
            templateList = templateList & i & ": " & templatesSheet.Cells(i, 2).Value
        End If
    Next i
    
    ' Show template selection dialog
    Dim selection As String
    selection = InputBox("Select an email template (enter number):" & vbCrLf & templateList, "Select Template")
    
    ' Check if canceled
    If selection = "" Then
        ShowEmailTemplateSelector = ""
        Exit Function
    End If
    
    ' Convert to template ID
    If IsNumeric(selection) Then
        templateID = "EM" & selection
    Else
        templateID = selection
    End If
    
    ShowEmailTemplateSelector = templateID
End Function

Sub SendTemplateEmail(customerName As String, customerEmail As String, templateID As String)
    ' Send an email using a template
    On Error Resume Next
    
    Dim objOutlook As Object
    Dim objMail As Object
    Dim templatesSheet As Worksheet
    Dim templateRow As Long
    Dim emailSubject As String
    Dim emailBody As String
    
    ' Get Outlook
    If Not InitializeOutlook() Then
        MsgBox "Could not connect to Outlook. Please ensure Outlook is running.", vbExclamation
        Exit Sub
    End If
    
    Set objOutlook = GetObject(, "Outlook.Application")
    
    ' Get templates sheet
    Set templatesSheet = ThisWorkbook.Sheets("Templates")
    
    ' Find template
    If left(templateID, 2) = "EM" Then
        ' Template ID is a row number
        templateRow = Val(Mid(templateID, 3))
    Else
        ' Look up template by ID
        Dim templateCell As Range
        Set templateCell = templatesSheet.Range("A:A").Find(templateID, LookIn:=xlValues)
        
        If templateCell Is Nothing Then
            MsgBox "Template not found: " & templateID, vbExclamation
            Exit Sub
        End If
        
        templateRow = templateCell.row
    End If
    
    ' Get template data
    emailSubject = templatesSheet.Cells(templateRow, 3).Value ' Column C
    emailBody = templatesSheet.Cells(templateRow, 4).Value ' Column D
    
    ' Replace merge fields
    emailSubject = Replace(emailSubject, "[Customer Name]", customerName)
    emailBody = Replace(emailBody, "[Customer Name]", customerName)
    
    ' Get customer info for other merge fields
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If Not customerRow Is Nothing Then
        ' Replace additional merge fields if found
        emailSubject = Replace(emailSubject, "[Stage]", customerRow.Offset(0, 3).Value)
        emailBody = Replace(emailBody, "[Stage]", customerRow.Offset(0, 3).Value)
        
        ' Vehicle info
        If Not IsEmpty(customerRow.Offset(0, 7).Value) Then
            emailSubject = Replace(emailSubject, "[Vehicle]", customerRow.Offset(0, 7).Value)
            emailBody = Replace(emailBody, "[Vehicle]", customerRow.Offset(0, 7).Value)
        End If
    End If
    
    ' Create email
    Set objMail = objOutlook.CreateItem(0) ' 0 = olMailItem
    
    With objMail
        .to = customerEmail
        .subject = emailSubject
        .HTMLBody = emailBody
        .Display ' Show email for review before sending
    End With
    
    ' Log email in contact history
    AddContactHistoryRecord customerName, "Email Sent", emailSubject, Now()
    
    ' Clean up
    Set objMail = Nothing
    Set objOutlook = Nothing
End Sub

Sub ResetStatusBar()
    ' Reset the status bar
    Application.StatusBar = False
End Sub
