Attribute VB_Name = "DataProcessor"
Sub StartNewCall()
    ' Launch the call assistant with customer selection
    Dim customerName As String
    
    ' Create customer selection form
    customerName = ShowCustomerSelector()
    
    If customerName <> "" Then
        ' Start call with selected customer
        StartCallWithCustomer customerName
    End If
End Sub

Function ShowCustomerSelector() As String
    ' Simple customer selector using InputBox for now
    ' In a production environment, this would be a proper form
    Dim customerName As String
    
    customerName = InputBox("Enter customer name or select from list:", "Start Call")
    
    ' If customer name is empty, user canceled
    If customerName = "" Then
        ShowCustomerSelector = ""
        Exit Function
    End If
    
    ' Check if customer exists
    Dim customerSheet As Worksheet
    Dim customerFound As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerFound = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If customerFound Is Nothing Then
        ' Customer not found, ask if they want to create new
        If MsgBox("Customer not found. Create new?", vbYesNo) = vbYes Then
            ' Create new customer
            CreateNewCustomer customerName
            ShowCustomerSelector = customerName
        Else
            ' Try again
            ShowCustomerSelector = ShowCustomerSelector()
        End If
    Else
        ' Customer found
        ShowCustomerSelector = customerName
    End If
End Function

Sub StartCallWithCustomer(customerName As String)
    ' Start a call with specified customer
    Dim callSheet As Worksheet
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    Dim callSheetExists As Boolean
    
    ' Check if CallAssistant sheet exists
    callSheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "CallAssistant" Then
            callSheetExists = True
            Exit For
        End If
    Next ws
    
    ' Create CallAssistant sheet if it doesn't exist
    If Not callSheetExists Then
        Set callSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        callSheet.Name = "CallAssistant"
        
        ' Create call assistant layout
        CreateCallAssistantLayout callSheet
    Else
        Set callSheet = ThisWorkbook.Sheets("CallAssistant")
    End If
    
    ' Find customer in tracker
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    ' If customer not found, exit
    If customerRow Is Nothing Then
        MsgBox "Customer not found: " & customerName, vbExclamation
        Exit Sub
    End If
    
    ' Populate call assistant with customer info
    callSheet.Range("C3").Value = customerName
    callSheet.Range("C4").Value = customerRow.Offset(0, 2).Value ' Phone from column D
    callSheet.Range("C5").Value = customerRow.Offset(0, 3).Value ' Stage from column E
    
    ' Start call timer
    callSheet.Range("G3").Value = Now()
    callSheet.Range("G4").Value = "00:00:00"
    Application.OnTime Now + TimeValue("00:00:01"), "UpdateCallTimer"
    
    ' Load appropriate script based on customer stage
    LoadCallScript customerRow.Offset(0, 3).Value
    
    ' Log call start in contact history
    AddContactHistoryRecord customerName, "Outbound Call", "Call started", Now()
    
    ' Activate call assistant sheet
    callSheet.Activate
End Sub

Sub CreateCallAssistantLayout(callSheet As Worksheet)
    ' Create layout for call assistant sheet
    
    ' Clear sheet
    callSheet.Cells.Clear
    
    ' Set column widths
    callSheet.Columns("A").ColumnWidth = 2
    callSheet.Columns("B").ColumnWidth = 12
    callSheet.Columns("C:D").ColumnWidth = 20
    callSheet.Columns("E").ColumnWidth = 2
    callSheet.Columns("F").ColumnWidth = 12
    callSheet.Columns("G").ColumnWidth = 20
    
    ' Add header
    callSheet.Range("B1:G1").Merge
    callSheet.Range("B1").Value = "CALL ASSISTANT"
    callSheet.Range("B1").Font.Bold = True
    callSheet.Range("B1").Font.Size = 16
    callSheet.Range("B1").HorizontalAlignment = xlCenter
    callSheet.Range("B1").Interior.Color = RGB(0, 66, 37) ' SG Fleet green
    callSheet.Range("B1").Font.Color = RGB(255, 255, 255) ' White
    
    ' Add customer info section
    callSheet.Range("B3").Value = "Customer:"
    callSheet.Range("B4").Value = "Phone:"
    callSheet.Range("B5").Value = "Stage:"
    callSheet.Range("B3:B5").Font.Bold = True
    
    ' Add call details section
    callSheet.Range("F3").Value = "Start Time:"
    callSheet.Range("F4").Value = "Duration:"
    callSheet.Range("F3:F4").Font.Bold = True
    
    ' Add script section
    callSheet.Range("B7:G7").Merge
    callSheet.Range("B7").Value = "SCRIPT"
    callSheet.Range("B7").Font.Bold = True
    callSheet.Range("B7").Interior.Color = RGB(240, 240, 240)
    
    ' Add script content area
    callSheet.Range("B8:G25").Merge
    callSheet.Range("B8").WrapText = True
    
    ' Add notes section
    callSheet.Range("B27:G27").Merge
    callSheet.Range("B27").Value = "NOTES"
    callSheet.Range("B27").Font.Bold = True
    callSheet.Range("B27").Interior.Color = RGB(240, 240, 240)
    
    ' Add notes text area
    callSheet.Range("B28:G32").Merge
    callSheet.Range("B28").WrapText = True
    
    ' Add buttons
    Dim btnEndCall As Button
    Set btnEndCall = callSheet.Buttons.Add(callSheet.Range("B34").left, callSheet.Range("B34").top, 100, 25)
    With btnEndCall
        .Name = "EndCallBtn"
        .caption = "End Call"
        .OnAction = "EndCurrentCall"
    End With
    
    Dim btnFollowUp As Button
    Set btnFollowUp = callSheet.Buttons.Add(callSheet.Range("D34").left, callSheet.Range("D34").top, 100, 25)
    With btnFollowUp
        .Name = "FollowUpBtn"
        .caption = "Schedule Follow-Up"
        .OnAction = "ScheduleCallFollowUp"
    End With
    
    Dim btnCreateQuote As Button
    Set btnCreateQuote = callSheet.Buttons.Add(callSheet.Range("F34").left, callSheet.Range("F34").top, 100, 25)
    With btnCreateQuote
        .Name = "CreateQuoteBtn"
        .caption = "Create Quote"
        .OnAction = "CreateQuoteForCustomer"
    End With
End Sub

Sub LoadCallScript(customerStage As String)
    ' Load appropriate script based on customer stage
    Dim callSheet As Worksheet
    Dim scriptContent As String
    
    Set callSheet = ThisWorkbook.Sheets("CallAssistant")
    
    ' Determine which script to use based on stage
    Select Case customerStage
        Case "Initial Call"
            scriptContent = GetInitialCallScript()
        Case "Quote Sent"
            scriptContent = GetQuoteFollowUpScript()
        Case "Finance Application"
            scriptContent = GetFinanceApplicationScript()
        Case "Vehicle Procurement"
            scriptContent = GetVehicleProcurementScript()
        Case "Settlement"
            scriptContent = GetSettlementScript()
        Case Else
            scriptContent = GetGeneralScript()
    End Select
    
    ' Update script content
    callSheet.Range("B8").Value = scriptContent
End Sub

Function GetInitialCallScript() As String
    ' Return the initial call script
    Dim Script As String
    
    Script = "INITIAL CALL SCRIPT" & vbCrLf & vbCrLf
    Script = Script & "Introduction:" & vbCrLf
    Script = Script & "Hi [Customer Name], It's Harrington calling from SG Fleet. How are you?" & vbCrLf
    Script = Script & "Just a reminder that our calls are recorded for training" & vbCrLf
    Script = Script & "I understand you are calling about a novated lease? Is that right?" & vbCrLf
    
    Script = Script & "Qualifying Questions:" & vbCrLf
    Script = Script & "Great! Well I'll Start by asking you a few questions? Just so I can make sure the lease best reflects what you need. Then we will go through how it works and I'll explain all the benefit available to you." & vbCrLf & vbCrLf & _
    Script = Script & "1. Have you ever had a novated lease before, or is this your first time?" & vbCrLf
    Script = Script & "2. Do you have a specific car in mind? (new/used electric or petrol)" & vbCrLf
    Script = Script & "3. Have you test driven the car?" & vbCrLf
    Script = Script & "5. When would you like to be in the new vehicle? Or do you already have an ETA for Delivery" & vbCrLf
    Script = Script & "6. How are you paying for your current car? Did you pay cash, car finance or throught the home loan?" & vbCrLf & vbCrLf
    Script = Script & "I will go through how a novated lease works and the benefits you get and then I will get some details from you so I can write up an indicative quote for you."

    Script = Script & "Education:" & vbCrLf
    Script = Script & "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf
    Script = Script & "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars."
    Script = Script & "The novated lease will include; the car, services, maintenance, tyres, rego, insurance and fuel. We will set a budget that reflects how many kilometres you drive each year." & vbCrLf & vbCrLf & _
    Script = Script & "Getting your budget correct is important with a Novated Lease, if we under estimate your running costs the lease will look cheap and attractive, however you won't have money to cover all the running costs. If we over-budget, it will look too expensive and you won't want to take the lease."
    
    GetInitialCallScript = Script
End Function

Function GetQuoteFollowUpScript() As String
    ' Return the quote follow-up script
    Dim Script As String
    
    Script = "QUOTE FOLLOW-UP SCRIPT" & vbCrLf & vbCrLf
    Script = Script & "Introduction:" & vbCrLf
    Script = Script & "Hi [Customer Name], it's [Your Name] from SG Fleet. I'm calling to follow up on the novated lease quote I sent through for the [Vehicle Make/Model]." & vbCrLf
    Script = Script & "Have you had a chance to review the quote?" & vbCrLf & vbCrLf
    
    Script = Script & "If Yes:" & vbCrLf
    Script = Script & "Great! What did you think of the quote? Do you have any questions I can answer for you?" & vbCrLf & vbCrLf
    
    Script = Script & "If No:" & vbCrLf
    Script = Script & "No problem. When would be a good time for us to discuss the quote? I'd be happy to walk you through it and answer any questions you might have." & vbCrLf & vbCrLf
    
    Script = Script & "Close:" & vbCrLf
    Script = Script & "Would you like to proceed with the application process now? I can guide you through the next steps. Most of my customer get finance Pre-Approved. The approval will last 6 months and it just prevents any unwelcomed delays should you decide to go ahead."
    
    GetQuoteFollowUpScript = Script
End Function

Function GetFinanceApplicationScript() As String
    ' Return finance application script
    Dim Script As String
    
    Script = "FINANCE APPLICATION SCRIPT" & vbCrLf & vbCrLf
    Script = Script & "Introduction:" & vbCrLf
    Script = Script & "Hi [Customer Name], it's [Your Name] from SG Fleet. I'm calling regarding your finance application for the [Vehicle Make/Model]." & vbCrLf & vbCrLf
    
    Script = Script & "Process Overview:" & vbCrLf
    Script = Script & "I'm going to help guide you through the finance application process. We'll need to collect some additional information to process your application." & vbCrLf & vbCrLf
    
    Script = Script & "Information Needed:" & vbCrLf
    Script = Script & "1. Employment details and duration" & vbCrLf
    Script = Script & "2. Income verification" & vbCrLf
    Script = Script & "3. Identification documents" & vbCrLf
    Script = Script & "4. Banking details" & vbCrLf & vbCrLf
    
    Script = Script & "Next Steps:" & vbCrLf
    Script = Script & "Once we have all your information, we'll submit your application to our finance team for approval. This typically takes 1-2 business days."
    
    GetFinanceApplicationScript = Script
End Function

Function GetVehicleProcurementScript() As String
    ' Return vehicle procurement script
    Dim Script As String
    
    Script = "VEHICLE PROCUREMENT SCRIPT" & vbCrLf & vbCrLf
    Script = Script & "Introduction:" & vbCrLf
    Script = Script & "Hi [Customer Name], it's [Your Name] from SG Fleet. I'm calling about your finance application. You have been approved! Congratulations!! SO, the next steps is procuring your [Vehicle Make/Model]." & vbCrLf & vbCrLf
    
    Script = Script & "Vehicle Status:" & vbCrLf
    Script = Script & "I'd like to update you on the status of your vehicle order and confirm some details with you." & vbCrLf & vbCrLf
    
    Script = Script & "Confirmation:" & vbCrLf
    Script = Script & "1. Confirm vehicle details (make, model, color, options)" & vbCrLf
    Script = Script & "2. Confirm delivery preferences" & vbCrLf
    Script = Script & "3. Discuss timeline expectations" & vbCrLf & vbCrLf
    
    Script = Script & "Next Steps:" & vbCrLf
    Script = Script & "I'll keep you updated on the progress of your vehicle order. Once we have a confirmed delivery date, I'll contact you to arrange the final documentation and delivery details."
    
    GetVehicleProcurementScript = Script
End Function

Function GetSettlementScript() As String
    ' Return settlement script
    Dim Script As String
    
    Script = "SETTLEMENT SCRIPT" & vbCrLf & vbCrLf
    Script = Script & "Introduction:" & vbCrLf
    Script = Script & "Hi [Customer Name], it's [Your Name] from SG Fleet. I'm calling to discuss the settlement process for your novated lease on the [Vehicle Make/Model]." & vbCrLf & vbCrLf
    
    Script = Script & "Settlement Process:" & vbCrLf
    Script = Script & "We're now ready to proceed with the settlement of your lease. I'd like to confirm a few final details and walk you through what happens next." & vbCrLf & vbCrLf
    
    Script = Script & "Final Steps:" & vbCrLf
    Script = Script & "1. Confirm final settlement figures" & vbCrLf
    Script = Script & "2. Arrange document signing" & vbCrLf
    Script = Script & "3. Schedule vehicle delivery/handover" & vbCrLf
    Script = Script & "4. Set up first payment date" & vbCrLf & vbCrLf
    
    Script = Script & "Support:" & vbCrLf
    Script = Script & "Once everything is settled, I'll be your main point of contact for any questions or support you need throughout your lease. You'll also receive details about your online account where you can manage your lease."
    
    GetSettlementScript = Script
End Function

Function GetGeneralScript() As String
    ' Return general script for other scenarios
    Dim Script As String
    
    Script = "GENERAL FOLLOW-UP SCRIPT" & vbCrLf & vbCrLf
    Script = Script & "Introduction:" & vbCrLf
    Script = Script & "Hi [Customer Name], it's [Your Name] from SG Fleet. I'm calling to follow up on your novated lease enquiry." & vbCrLf & vbCrLf
    
    Script = Script & "Status Check:" & vbCrLf
    Script = Script & "I wanted to check in and see if you have any questions about novated leasing that I can help with?" & vbCrLf & vbCrLf
    
    Script = Script & "Value Proposition:" & vbCrLf
    Script = Script & "Just to remind you of the benefits of a novated lease:" & vbCrLf
    Script = Script & "1. Potential tax savings" & vbCrLf
    Script = Script & "2. GST savings on the purchase price and running costs" & vbCrLf
    Script = Script & "3. Budgeting simplicity with one regular payment" & vbCrLf
    Script = Script & "4. No large upfront payment needed" & vbCrLf & vbCrLf
    
    Script = Script & "Next Steps:" & vbCrLf
    Script = Script & "Would you like me to prepare a quote for you to see the potential benefits for your specific situation?"
    
    GetGeneralScript = Script
End Function

Sub UpdateCallTimer()
    ' Update the call duration timer
    Dim callSheet As Worksheet
    Dim startTime As Date
    Dim currentDuration As String
    
    ' Only proceed if call assistant sheet exists
    On Error Resume Next
    Set callSheet = ThisWorkbook.Sheets("CallAssistant")
    If callSheet Is Nothing Then Exit Sub
    
    ' Get start time
    If IsEmpty(callSheet.Range("G3").Value) Then Exit Sub
    startTime = callSheet.Range("G3").Value
    
    ' Calculate duration
    currentDuration = Format(Now - startTime, "hh:mm:ss")
    
    ' Update duration display
    callSheet.Range("G4").Value = currentDuration
    
    ' Schedule next update
    Application.OnTime Now + TimeValue("00:00:01"), "UpdateCallTimer"
End Sub

Sub EndCurrentCall()
    ' End the current call
    Dim callSheet As Worksheet
    Dim customerName As String
    Dim startTime As Date
    Dim duration As String
    Dim notes As String
    Dim outcome As String
    
    ' Get call sheet
    Set callSheet = ThisWorkbook.Sheets("CallAssistant")
    
    ' Get call details
    customerName = callSheet.Range("C3").Value
    startTime = callSheet.Range("G3").Value
    duration = callSheet.Range("G4").Value
    notes = callSheet.Range("B28").Value
    
    ' Prompt for outcome
    outcome = InputBox("Enter call outcome (Completed, No Answer, Call Back, etc.):", "End Call")
    If outcome = "" Then outcome = "Completed" ' Default
    
    ' Log call in contact history
    AddContactHistoryRecord customerName, "Outbound Call", outcome & " (" & duration & ") - " & notes, Now()
    
    ' Update call planner if this was a scheduled call
    UpdateCallPlanner customerName, outcome
    
    ' Update customer record
    UpdateCustomerRecord customerName, notes
    
    ' Stop timer
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "UpdateCallTimer", , False
    
    ' Return to dashboard
    ThisWorkbook.Sheets("Dashboard").Activate
    
    ' Show confirmation
    MsgBox "Call with " & customerName & " ended. Duration: " & duration, vbInformation
End Sub

Sub ScheduleCallFollowUp()
    ' Schedule a follow-up call
    Dim callSheet As Worksheet
    Dim customerName As String
    Dim followupDate As Date
    Dim followupType As String
    
    ' Get call sheet
    Set callSheet = ThisWorkbook.Sheets("CallAssistant")
    
    ' Get customer name
    customerName = callSheet.Range("C3").Value
    
    ' Get follow-up details
    followupDate = DateValue(InputBox("Enter follow-up date (MM/DD/YYYY):", "Schedule Follow-up", Format(Date + 3, "MM/DD/YYYY")))
    followupType = InputBox("Enter follow-up type:", "Schedule Follow-up", "Follow-up call")
    
    ' Update customer record with follow-up info
    UpdateCustomerFollowUp customerName, followupDate, followupType
    
    ' Create Outlook reminder if available
    If InitializeOutlook() Then
        CreateOutlookReminder customerName, followupDate, followupType
    End If
    
    ' Show confirmation
    MsgBox "Follow-up scheduled for " & customerName & " on " & Format(followupDate, "dd-mmm-yyyy"), vbInformation
End Sub

Sub UpdateCustomerFollowUp(customerName As String, followupDate As Date, followupType As String)
    ' Update customer record with follow-up information
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    ' Get customer sheet
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    
    ' Find customer
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    ' Update customer record
    If Not customerRow Is Nothing Then
        customerRow.Offset(0, 5).Value = followupType ' Next Action column
        customerRow.Offset(0, 6).Value = followupDate ' Next Action Date column
    End If
End Sub

Sub CreateOutlookReminder(customerName As String, reminderDate As Date, reminderSubject As String, Optional notes As String = "")
    ' Creates a reminder in Outlook for a customer follow-up
    On Error Resume Next
    
    Dim objOutlook As Object
    Dim objTask As Object
    
    ' Try to get Outlook instance
    Set objOutlook = GetObject(, "Outlook.Application")
    If objOutlook Is Nothing Then
        Set objOutlook = CreateObject("Outlook.Application")
    End If
    
    If objOutlook Is Nothing Then
        MsgBox "Could not connect to Outlook. Please check if Outlook is installed and running.", vbExclamation
        Exit Sub
    End If
    
    ' Create a task as a reminder
    Set objTask = objOutlook.CreateItem(3) ' 3 = olTaskItem
    
    With objTask
        .subject = reminderSubject & " - " & customerName
        .body = "Follow-up with " & customerName & vbCrLf & vbCrLf & notes
        .dueDate = reminderDate
        .ReminderSet = True
        .ReminderTime = reminderDate
        .Save
    End With
    
    Set objTask = Nothing
    Set objOutlook = Nothing
    
    MsgBox "Reminder created in Outlook for " & Format(reminderDate, "dd-mmm-yyyy"), vbInformation
End Sub
