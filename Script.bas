Attribute VB_Name = "Script"
' ====================================================================
' DynamicScript Module
' ====================================================================
' This module creates an interactive script system that guides users
' through the novated leasing sales process with branching conversation paths

Option Explicit

' Script navigation states
Private currentScriptView As String
Private scriptHistory As Collection
Private customerResponses As Object ' Dictionary
Private scriptStartTime As Date
Private activeCustomerName As String
Private activeCustomerPhone As String
Private scriptNotes As String

' Script view constants
Private Const VIEW_INITIAL = "InitialView"
Private Const VIEW_NOT_MUCH = "NotMuchKnowledgeView"
Private Const VIEW_A_LITTLE = "ALittleKnowledgeView"
Private Const VIEW_WELL_EDUCATED = "WellEducatedView"
Private Const VIEW_QUALIFYING = "QualifyingView"
Private Const VIEW_EDUCATING = "EducatingView"
Private Const VIEW_BENEFITS = "BenefitsView"
Private Const VIEW_LEASE_END = "LeaseEndView"
Private Const VIEW_TRADE_IN = "TradeInView"
Private Const VIEW_OBJECTIONS = "ObjectionsView"
Private Const VIEW_GATHER_DETAILS = "GatherDetailsView"
Private Const VIEW_CLOSING = "ClosingView"

' Application colors for consistent styling
Private Const COLOR_PRIMARY = 4227072     ' Dark green (RGB 0, 66, 37)
Private Const COLOR_SECONDARY = 39423     ' Orange (RGB 255, 153, 0)
Private Const COLOR_TEXT_LIGHT = 16777215 ' White
Private Const COLOR_BACKGROUND = 15921906 ' Light gray (RGB 242, 242, 242)
Private Const COLOR_BUTTON = 5287936     ' Medium green (RGB 80, 160, 80)
Private Const COLOR_BUTTON_HIGHLIGHT = 49344 ' Orange highlight (RGB 192, 192, 0)

' Set up script system
Public Sub InitializeScriptSystem()
On Error Resume Next
    ' Create script sheet if it doesn't exist
    Dim scriptSheet As Worksheet
    
    On Error Resume Next
    Set scriptSheet = ThisWorkbook.Sheets("DynamicScript")
    On Error GoTo 0
    
    If scriptSheet Is Nothing Then
        ' Create the script sheet
        Set scriptSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        scriptSheet.Name = "DynamicScript"
        
        ' Setup initial layout
        SetupScriptLayout scriptSheet
    End If
    
    ' Initialize script state
    Set scriptHistory = New Collection
    Set customerResponses = CreateObject("Scripting.Dictionary")
    
    ' Set up the initial view
    currentScriptView = VIEW_INITIAL
    
    ' Update the view
    UpdateScriptView
End Sub

' Set up the script layout
Private Sub SetupScriptLayout(scriptSheet As Worksheet)
On Error Resume Next
    ' Clear any existing content
    scriptSheet.Cells.Clear
    
    ' Set worksheet background color
    scriptSheet.Tab.Color = COLOR_PRIMARY
    
    ' Set column widths for better layout
    scriptSheet.Columns("A:A").ColumnWidth = 2
    scriptSheet.Columns("B:I").ColumnWidth = 12
    scriptSheet.Columns("J:J").ColumnWidth = 2
    
    ' Add header
    With scriptSheet.Range("B1:I1")
        .Merge
        .Value = "NOVATED LEASE DYNAMIC SCRIPT"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
    End With
    
    ' Add navigation area
    With scriptSheet.Range("B2:I2")
        .Merge
        .Value = "Current Path: Initial Assessment"
        .Font.Bold = True
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_BACKGROUND
    End With
    
    ' Add back button
    AddBackButton scriptSheet
    
    ' Add customer info area - header row
    With scriptSheet.Range("B3:I3")
        .Interior.Color = COLOR_SECONDARY
        .Font.Color = COLOR_TEXT_LIGHT
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    With scriptSheet.Range("B3:D3")
        .Merge
        .Value = "Customer:"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    With scriptSheet.Range("E3:I3")
        .Merge
        .Value = ""
        .Interior.Color = COLOR_BACKGROUND
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
    End With
    
    ' Phone row
    With scriptSheet.Range("B4:I4")
        .Interior.Color = COLOR_SECONDARY
        .Font.Color = COLOR_TEXT_LIGHT
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    With scriptSheet.Range("B4:D4")
        .Merge
        .Value = "Phone:"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    With scriptSheet.Range("E4:I4")
        .Merge
        .Value = ""
        .Interior.Color = COLOR_BACKGROUND
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
    End With
    
    ' Duration row
    With scriptSheet.Range("B5:I5")
        .Interior.Color = COLOR_SECONDARY
        .Font.Color = COLOR_TEXT_LIGHT
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    With scriptSheet.Range("B5:D5")
        .Merge
        .Value = "Call Duration:"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    With scriptSheet.Range("E5:I5")
        .Merge
        .Value = "00:00:00"
        .Interior.Color = COLOR_BACKGROUND
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
    End With
    
    ' Script content area - header
    With scriptSheet.Range("B7:I7")
        .Merge
        .Value = "SCRIPT:"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
    End With
    
    ' Word track area
    With scriptSheet.Range("B8:I20")
        .Merge
        .Value = "Press 'Start New Call' to begin"
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = COLOR_PRIMARY
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With
    
    ' Response options area - header
    With scriptSheet.Range("B21:I21")
        .Merge
        .Value = "CUSTOMER RESPONSE:"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
    End With
    
    ' Response options area - content
    With scriptSheet.Range("B22:I32")
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = COLOR_PRIMARY
    End With
    
    ' Notes area - header
    With scriptSheet.Range("B33:I33")
        .Merge
        .Value = "CALL NOTES:"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
    End With
    
    ' Notes area - content
    With scriptSheet.Range("B34:I40")
        .Merge
        .Value = ""
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = COLOR_PRIMARY
        .Font.Name = "Calibri"
        .Font.Size = 11
    End With
    
    ' Add a separator line
    With scriptSheet.Range("B41:I41")
        .Merge
        .Interior.Color = COLOR_BACKGROUND
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = COLOR_PRIMARY
    End With
    
    ' Action buttons
    AddActionButtons scriptSheet
    
    ' Select a cell to avoid accidental editing
    scriptSheet.Range("A1").Select
End Sub

' Add back button to the script sheet
Private Sub AddBackButton(scriptSheet As Worksheet)
On Error Resume Next
    Dim backBtn As Button
    
    ' Remove existing button if any
    On Error Resume Next
    scriptSheet.Buttons("BackButton").Delete
    On Error GoTo 0
    
    ' Add new button
    Set backBtn = scriptSheet.Buttons.Add(scriptSheet.Range("B2").left, _
                                 scriptSheet.Range("B2").top, _
                                 25, 18)
    With backBtn
        .caption = "?"
        .Name = "BackButton"
        .OnAction = "ScriptNavigateBack"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
End Sub

' Add action buttons to the script sheet
Private Sub AddActionButtons(scriptSheet As Worksheet)
On Error Resume Next
    Dim startBtn As Button, endBtn As Button, followUpBtn As Button, saveNoteBtn As Button
    
    ' Remove existing buttons if any
    On Error Resume Next
    scriptSheet.Buttons("StartCallButton").Delete
    scriptSheet.Buttons("EndCallButton").Delete
    scriptSheet.Buttons("FollowUpButton").Delete
    scriptSheet.Buttons("SaveNoteButton").Delete
    On Error GoTo 0
    
    ' Calculate button positions
    Dim btnTop As Double
    Dim btnHeight As Double
    Dim btnWidth As Double
    Dim btnSpacing As Double
    
    btnTop = scriptSheet.Range("B42").top
    btnHeight = 30
    btnWidth = 120
    btnSpacing = 10
    
    ' Start call button
    Set startBtn = scriptSheet.Buttons.Add(scriptSheet.Range("B42").left, _
                                 btnTop, btnWidth, btnHeight)
    With startBtn
        .caption = "Start New Call"
        .Name = "StartCallButton"
        .OnAction = "StartNewScript"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    ' End call button
    Set endBtn = scriptSheet.Buttons.Add(scriptSheet.Range("D42").left, _
                                 btnTop, btnWidth, btnHeight)
    With endBtn
        .caption = "End Call"
        .Name = "EndCallButton"
        .OnAction = "EndCurrentScript"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    ' Schedule follow-up button
    Set followUpBtn = scriptSheet.Buttons.Add(scriptSheet.Range("F42").left, _
                                 btnTop, btnWidth, btnHeight)
    With followUpBtn
        .caption = "Schedule Follow-up"
        .Name = "FollowUpButton"
        .OnAction = "ScheduleScriptFollowUp"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    ' Save note button
    Set saveNoteBtn = scriptSheet.Buttons.Add(scriptSheet.Range("H42").left, _
                                 btnTop, btnWidth, btnHeight)
    With saveNoteBtn
        .caption = "Save Notes"
        .Name = "SaveNoteButton"
        .OnAction = "SaveScriptNotes"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
End Sub

' Start a new script session with a customer
Public Sub StartNewScript()
On Error Resume Next
    Dim customerName As String
    Dim customerPhone As String
    Dim callSheet As Worksheet
    Dim selectedCell As Range
    Dim customerRow As Range
    
    ' Check if we're in the CallPlanner sheet
    Set callSheet = ActiveSheet
    Set selectedCell = selection
    
    If callSheet.Name = "CallPlanner" Then
        ' Get customer from the selected row
        customerName = callSheet.Cells(selectedCell.row, 2).Value
        customerPhone = callSheet.Cells(selectedCell.row, 3).Value
        
        ' Mark call as in progress
        callSheet.Cells(selectedCell.row, 7).Value = "In Progress"
    Else
        ' Prompt for customer details
        customerName = InputBox("Enter customer name:", "Start New Call")
        If customerName = "" Then Exit Sub
        
        customerPhone = InputBox("Enter customer phone number:", "Start New Call")
    End If
    
    ' Initialize script state
    Set scriptHistory = New Collection
    Set customerResponses = CreateObject("Scripting.Dictionary")
    currentScriptView = VIEW_INITIAL
    scriptStartTime = Now
    activeCustomerName = customerName
    activeCustomerPhone = customerPhone
    scriptNotes = ""
    
    ' Navigate to the script sheet
    ThisWorkbook.Sheets("DynamicScript").Activate
    
    ' Update customer info
    With ThisWorkbook.Sheets("DynamicScript")
        .Range("E3").Value = customerName
        .Range("E4").Value = customerPhone
    End With
    
    ' Start the timer
    Application.OnTime Now + TimeValue("00:00:01"), "UpdateScriptTimer"
    
    ' Update the view
    UpdateScriptView
    
    ' Log the call start
    LogCustomerContact customerName, "Outbound Call", "Call started", Now()
    
    ' Update customer's last contact in tracker
    UpdateCustomerLastContact customerName
End Sub

' Update the customer's last contact date
Private Sub UpdateCustomerLastContact(customerName As String)
On Error Resume Next
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If Not customerRow Is Nothing Then
        customerRow.Offset(0, 4).Value = Now() ' Last Contact column
    End If
End Sub

' Update the script timer display
Public Sub UpdateScriptTimer()
On Error Resume Next
    Dim scriptSheet As Worksheet
    Dim elapsedTime As Date
    
    ' Only continue if we have an active call
    If activeCustomerName = "" Then Exit Sub
    
    Set scriptSheet = ThisWorkbook.Sheets("DynamicScript")
    
    ' Calculate elapsed time
    elapsedTime = Now - scriptStartTime
    
    ' Update the display
    scriptSheet.Range("E5").Value = Format(elapsedTime, "hh:mm:ss")
    
    ' Schedule next update
    Application.OnTime Now + TimeValue("00:00:01"), "UpdateScriptTimer"
End Sub

' End the current script session
Public Sub EndCurrentScript()
On Error Resume Next
    Dim outcome As String
    Dim callDuration As Date
    Dim callSheet As Worksheet
    Dim callRow As Range
    
    ' Check if we have an active call
    If activeCustomerName = "" Then
        MsgBox "No active call to end.", vbExclamation
        Exit Sub
    End If
    
    ' Calculate call duration
    callDuration = Now - scriptStartTime
    
    ' Get call outcome
    outcome = InputBox("Enter call outcome (e.g., Quote Sent, Callback, Not Interested):", "End Call")
    If outcome = "" Then outcome = "Completed"
    
    ' Get notes from the script sheet
    scriptNotes = ThisWorkbook.Sheets("DynamicScript").Range("B34").Value
    
    ' Log call end
    LogCustomerContact activeCustomerName, "Outbound Call", _
        "Call completed (" & Format(callDuration, "hh:mm:ss") & ") - " & outcome & vbCrLf & _
        "Notes: " & scriptNotes, Now()
    
    ' Update call planner if this was a scheduled call
    Set callSheet = ThisWorkbook.Sheets("CallPlanner")
    Set callRow = callSheet.Range("B:B").Find(activeCustomerName, LookIn:=xlValues)
    
    If Not callRow Is Nothing Then
        callRow.Offset(0, 5).Value = outcome ' Outcome column
    End If
    
    ' Stop the timer
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "UpdateScriptTimer", , False
    On Error GoTo 0
    
    ' Reset values
    activeCustomerName = ""
    activeCustomerPhone = ""
    scriptStartTime = 0
    
    ' Clear the script display
    With ThisWorkbook.Sheets("DynamicScript")
        .Range("E3").Value = ""
        .Range("E4").Value = ""
        .Range("E5").Value = "00:00:00"
        .Range("B8").Value = "Press 'Start New Call' to begin"
        .Range("B34").Value = ""
    End With
    
    ' Reset the view
    currentScriptView = VIEW_INITIAL
    
    ' Show completion message
    MsgBox "Call completed and logged successfully.", vbInformation
    
    ' Return to call planner
    ThisWorkbook.Sheets("CallPlanner").Activate
End Sub

' Save notes from the script
Public Sub SaveScriptNotes()
On Error Resume Next
    Dim notes As String
    
    ' Check if we have an active call
    If activeCustomerName = "" Then
        MsgBox "No active call to save notes for.", vbExclamation
        Exit Sub
    End If
    
    ' Get notes from the script sheet
    notes = ThisWorkbook.Sheets("DynamicScript").Range("B34").Value
    
    ' Update customer record with notes
    UpdateCustomerNotes activeCustomerName, notes
    
    ' Confirm save
    MsgBox "Notes saved to customer record.", vbInformation
End Sub

' Update customer notes in the tracker
Private Sub UpdateCustomerNotes(customerName As String, notes As String)
On Error Resume Next
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If Not customerRow Is Nothing Then
        ' Append to existing notes or create new
        If Not IsEmpty(customerRow.Offset(0, 10).Value) Then ' Notes column
            customerRow.Offset(0, 10).Value = customerRow.Offset(0, 10).Value & vbCrLf & _
                Format(Now, "yyyy-mm-dd hh:mm") & " - " & notes
        Else
            customerRow.Offset(0, 10).Value = Format(Now, "yyyy-mm-dd hh:mm") & " - " & notes
        End If
    End If
End Sub

' Schedule a follow-up from the script
Public Sub ScheduleScriptFollowUp()
On Error Resume Next
    Dim followupDate As Date
    Dim followupType As String
    
    ' Check if we have an active call
    If activeCustomerName = "" Then
        MsgBox "No active call to schedule follow-up for.", vbExclamation
        Exit Sub
    End If
    
    ' Prompt for follow-up details
    followupDate = DateValue(InputBox("Enter follow-up date (MM/DD/YYYY):", "Schedule Follow-up", Format(Date + 3, "MM/DD/YYYY")))
    
    ' Validate date
    If followupDate < Date Then
        MsgBox "Follow-up date must be in the future.", vbExclamation
        Exit Sub
    End If
    
    followupType = InputBox("Enter follow-up type:", "Schedule Follow-up", "Follow-up call")
    
    ' Schedule the follow-up
    ScheduleCustomerFollowUp activeCustomerName, followupType, followupDate
    
    ' Confirm scheduling
    MsgBox "Follow-up scheduled for " & Format(followupDate, "mm/dd/yyyy") & ".", vbInformation
End Sub

' Schedule customer follow-up in tracker
Private Sub ScheduleCustomerFollowUp(customerName As String, actionType As String, actionDate As Date)
On Error Resume Next
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If Not customerRow Is Nothing Then
        customerRow.Offset(0, 6).Value = actionType ' Next Action column
        customerRow.Offset(0, 7).Value = actionDate ' Next Action Date column
    End If
End Sub

' Navigate back in the script history
Public Sub ScriptNavigateBack()
On Error Resume Next
    ' Check if we have history to go back to
    If scriptHistory.count = 0 Then
        MsgBox "You're at the beginning of the script.", vbInformation
        Exit Sub
    End If
    
    ' Get the previous view
    currentScriptView = scriptHistory(scriptHistory.count)
    
    ' Remove the current view from history
    If scriptHistory.count > 0 Then
        scriptHistory.Remove scriptHistory.count
    End If
    
    ' Update the view
    UpdateScriptView
End Sub

' Handle customer response option selection
Public Sub HandleScriptOptionSelection(optionValue As String)
On Error Resume Next
    ' Save the current view to history
    scriptHistory.Add currentScriptView
    
    ' Store the customer's response
    customerResponses(currentScriptView) = optionValue
    
    ' Determine next view based on current view and option selected
    Select Case currentScriptView
        Case VIEW_INITIAL
            ' Initial knowledge assessment
            Select Case optionValue
                Case "First Timer"
                    currentScriptView = VIEW_NOT_MUCH
                Case "Knows a little"
                    currentScriptView = VIEW_A_LITTLE
                Case "Mr. Know-it-All"
                    currentScriptView = VIEW_WELL_EDUCATED
                Case Else
                    currentScriptView = VIEW_QUALIFYING
            End Select
            
        Case VIEW_NOT_MUCH, VIEW_A_LITTLE, VIEW_WELL_EDUCATED
            ' After knowledge assessment, go to qualifying questions
            currentScriptView = VIEW_QUALIFYING
            
        Case VIEW_QUALIFYING
            ' After qualifying, go to education
            currentScriptView = VIEW_EDUCATING
            
        Case VIEW_EDUCATING
            ' After education, discuss benefits
            currentScriptView = VIEW_BENEFITS
            
        Case VIEW_BENEFITS
            ' After benefits, discuss lease end options
            currentScriptView = VIEW_LEASE_END
            
        Case VIEW_LEASE_END
            ' After lease end, discuss trade-in options
            currentScriptView = VIEW_TRADE_IN
            
        Case VIEW_TRADE_IN
            ' After trade-in, gather customer details
            currentScriptView = VIEW_GATHER_DETAILS
            
        Case VIEW_OBJECTIONS
            ' After handling objections, return to previous section or gather details
            If optionValue = "Objection Resolved" Then
                currentScriptView = VIEW_GATHER_DETAILS
            Else
                currentScriptView = VIEW_CLOSING
            End If
            
        Case VIEW_GATHER_DETAILS
            ' After gathering details, move to closing
            currentScriptView = VIEW_CLOSING
            
        Case VIEW_CLOSING
            ' After closing, end call or handle objections
            If optionValue = "Has Objections" Then
                currentScriptView = VIEW_OBJECTIONS
            Else
                ' End of script, return to initial
                currentScriptView = VIEW_INITIAL
                MsgBox "Script completed. Ready for next call.", vbInformation
            End If
            
        Case Else
            ' Default to qualifying if uncertain
            currentScriptView = VIEW_QUALIFYING
    End Select
    
    ' Update the view
    UpdateScriptView
End Sub

' Update the script view based on the current state
Private Sub UpdateScriptView()
On Error Resume Next
    Dim scriptSheet As Worksheet
    Dim scriptText As String
    Dim pathText As String
    
    Set scriptSheet = ThisWorkbook.Sheets("DynamicScript")
    
    ' Clear previous response options
    ClearResponseButtons scriptSheet
    
    ' Set the script text and response options based on the current view
    Select Case currentScriptView
        Case VIEW_INITIAL
            ' Initial script - introduction
            scriptText = "Hi " & activeCustomerName & ", my name is " & Application.userName & ", before I start, just a reminder that our calls are recorded for training purposes." & vbCrLf & vbCrLf & _
                         "I understand you are calling about a novated lease?"

            pathText = "Initial Assessment"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "First Timer"
            AddResponseOption scriptSheet, 2, "Knows a little"
            AddResponseOption scriptSheet, 3, "Mr. Know-it-All"
            AddResponseOption scriptSheet, 4, "Skip to Qualifying"
            
        Case VIEW_NOT_MUCH
            ' Not much knowledge script - exactly as in Script.V3
            scriptText = "That's perfectly fine - many people are new to novated leasing. Let me explain the basics:" & vbCrLf & vbCrLf & _
                         "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf & _
                         "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars."

            pathText = "Initial Assessment > Not Much Knowledge"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Qualifying"
            
        Case VIEW_A_LITTLE
            ' A little knowledge script
            scriptText = "Great, so you have some familiarity with novated leasing. Let me confirm and expand on what you might already know:" & vbCrLf & vbCrLf & _
                         "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf & _
                         "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars."
            
            pathText = "Initial Assessment > Some Knowledge"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Qualifying"
            
        Case VIEW_WELL_EDUCATED
            ' Well educated script
            scriptText = "Excellent! Since you're already familiar with novated leasing, let's focus on the specific aspects that would be most beneficial for your situation." & vbCrLf & vbCrLf & _
                         "I'd like to ask you some specific questions about your needs to tailor our discussion to what would work best for you."
            
            pathText = "Initial Assessment > Well Educated"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Qualifying"
            
        Case VIEW_QUALIFYING
            ' Qualifying questions script - exactly as in Script.V3
            scriptText = "So, I can make sure the lease reflects what you need, can I ask you a few questions?" & vbCrLf & vbCrLf & _
                         "1. Have you ever had a novated lease before?" & vbCrLf & _
                         "2. Do you have a car in mind? (new/used electric or petrol)" & vbCrLf & _
                         "3. Have you test driven the car?" & vbCrLf & _
                         "4. Do you have a budget in mind?" & vbCrLf & _
                         "5. When would you like to be in the new vehicle?" & vbCrLf & _
                         "6. How are you paying for your current car? Did you pay cash, car finance, home loan?" & vbCrLf & vbCrLf & _
                         "I will go through how a novated lease works and the benefits you get and then I will get some details from you so I can write up an indicative quote for you."
            
            pathText = "Qualifying Questions"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Education"
            AddResponseOption scriptSheet, 2, "Customer Has Objections"
            
        Case VIEW_EDUCATING
            ' Educating the customer script - exactly as in Script.V3
            scriptText = "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf & _
                         "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars." & vbCrLf & vbCrLf & _
                         "The novated lease will include; the car, services, maintenance, tyres, rego, insurance and fuel. We will set a budget that reflects how many kilometres you drive each year." & vbCrLf & vbCrLf & _
                         "Getting your budget correct is important with a Novated Lease, if we under estimate your running costs the lease will look cheap and attractive, however you won't have money to cover all the running costs. If we over-budget, it will look too expensive and you won't want to take the lease."
            
            pathText = "Education > Basics"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Benefits"
            AddResponseOption scriptSheet, 2, "Customer Has Questions"
            
        Case VIEW_BENEFITS
            ' Benefits script - exactly as in Script.V3
            scriptText = "The vehicle is also fully maintained, meaning you'll get fuel cards for your fuel, the dealer will invoice us for your servicing, same with the tyre shop. Rego renewals can be uploaded, and we'll pay them directly." & vbCrLf & vbCrLf & _
                         "So other than tolls and fines you should not need to outlay any further money on your vehicle outside of those deductions." & vbCrLf & vbCrLf & _
                         "The easy way to think of this, is that the deductions will go to an account held at sgfleet. As you spend money on the vehicle, it will deduct from that account and you can manage and monitor this through our online app." & vbCrLf & vbCrLf & _
                         "Throughout the lease you're paying the vehicle down to a residual or balloon amount set by the ATO."
            
            pathText = "Education > Benefits"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Lease End Options"
            AddResponseOption scriptSheet, 2, "Customer Has Questions"
            
        Case VIEW_LEASE_END
            ' Lease end options script - exactly as in Script.V3
            scriptText = "At the end of the lease, you will have a few options." & vbCrLf & vbCrLf & _
                         "1. If you love the car, you can pay it out and the car is yours;" & vbCrLf & _
                         "2. If you love the car and the lease, you can refinance and extend" & vbCrLf & _
                         "3. Another option is to sell the car -- any money you make above the residual is yours tax free" & vbCrLf & _
                         "4. Or the most popular option, is to trade the car in and upgrade." & vbCrLf & vbCrLf & _
                         "We can also source the vehicle for you, biggest benefit is that we are entitled to fleet discounts which we pass on to you." & vbCrLf & vbCrLf & _
                         "All our new quotes will include an Eco Protection Pack and Minor Damage Repair membership."
            
            pathText = "Education > Lease End Options"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Trade-In"
            AddResponseOption scriptSheet, 2, "Customer Has Questions"
            
        Case VIEW_TRADE_IN
            ' Trade-in information script - exactly as in Script.V3
            scriptText = "Do you have a car you are looking to trade in or sell?" & vbCrLf & vbCrLf & _
                         "We have a trade advantage program where we will have several wholesalers value your existing car. If you're happy with the price, you can trade it in and we will even come and collect the car for you." & vbCrLf & vbCrLf & _
                         "Do you have any questions at this point?"
            
            pathText = "Trade-In Information"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Gather Details"
            AddResponseOption scriptSheet, 2, "Customer Has Questions"
            
        Case VIEW_OBJECTIONS
            ' Handling objections script
            scriptText = "I understand your concerns. Many people have similar questions when first considering a novated lease." & vbCrLf & vbCrLf & _
                         "Common objections and our responses:" & vbCrLf & vbCrLf & _
                         "1. ""It seems expensive"" - While there is an upfront cost, the tax benefits and GST savings often make it more cost-effective than traditional car ownership over time." & vbCrLf & vbCrLf & _
                         "2. ""I'm worried about changing jobs"" - If you change employers, you have options: transfer the lease to your new employer, convert to a consumer loan, or pay it out." & vbCrLf & vbCrLf & _
                         "3. ""I don't understand the tax benefits"" - We can prepare a personalized calculation showing exactly how much you could save based on your salary and the vehicle you're interested in."
            
            pathText = "Handling Objections"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Objection Resolved"
            AddResponseOption scriptSheet, 2, "Not Interested"
            
        Case VIEW_GATHER_DETAILS
            ' Gather details script - exactly as in Script.V3
            scriptText = "So I can generate a quote, can I ask you some questions:" & vbCrLf & vbCrLf & _
                         "1. What is your gross annual income?" & vbCrLf & vbCrLf & _
                         "2. For insurance purposes, can I confirm your postcode and suburb? And in the last 3 years have you had 2 or more at fault claims? In the last 3 years have you been charged with a DUI or negligent driving?" & vbCrLf & vbCrLf & _
                         "3. Can I confirm the car details:" & vbCrLf & _
                         "   Vehicle Brand:" & vbCrLf & _
                         "   Model:" & vbCrLf & _
                         "   Variant:" & vbCrLf & _
                         "   Series:" & vbCrLf & _
                         "   Auto/Manual:" & vbCrLf & _
                         "   Preference on colour:" & vbCrLf & _
                         "   Any additional accessories (e.g. Floor Mats, Tow Bar):" & vbCrLf & vbCrLf & _
                         "4. How many years would you like the lease to be?" & vbCrLf & vbCrLf & _
                         "5. How many kilometres a year are you currently travelling?"
            
            pathText = "Gathering Details"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Closing"
            AddResponseOption scriptSheet, 2, "Customer Has Objections"
            
        Case VIEW_CLOSING
            ' Closing script - exactly as in Script.V3
            scriptText = "Thank you very much, I will send you an indicative quote to give you an idea of how the lease will look. Along with the quote you will also receive a link to complete the finance pre-approval and the email regarding our trade in program." & vbCrLf & vbCrLf & _
                         "Most customers choose to get their finance pre-approved while we find a car for you as the approval lasts for 6 months." & vbCrLf & vbCrLf & _
                         "Do you have any questions? Or would you like me to go over any point again?" & vbCrLf & vbCrLf & _
                         "Okay great, so just to reiterate, we're looking at [CAR DETAILS], you generally keep your cars for [YEARS], and you're doing about [KM] PA." & vbCrLf & vbCrLf & _
                         "Once I receive the pricing (expected 24 hours), I will call you and we can move on to the next step." & vbCrLf & vbCrLf & _
                         "If you have any questions that springs to mind, please don't hesitate to call/email." & vbCrLf & vbCrLf & _
                         "Thank you for your time today " & activeCustomerName & ", talk to you on [NEXT CONTACT DAY]."
            
            pathText = "Closing"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Complete Call"
            AddResponseOption scriptSheet, 2, "Has Objections"
            
        Case Else
            ' Default view
            scriptText = "Press 'Start New Call' to begin a new script session."
            pathText = "Initial Assessment"
    End Select
    
    ' Update the script sheet
    scriptSheet.Range("B2").Value = "Current Path: " & pathText
    scriptSheet.Range("B8").Value = scriptText
    
    ' If no active customer, clear customer info fields
    If activeCustomerName = "" Then
        scriptSheet.Range("E3").Value = ""
        scriptSheet.Range("E4").Value = ""
        scriptSheet.Range("E5").Value = "00:00:00"
    End If
End Sub

' Clear response buttons from the script sheet
Private Sub ClearResponseButtons(scriptSheet As Worksheet)
On Error Resume Next
    Dim i As Integer
    
    ' Remove existing response buttons
    For i = 1 To 6
        On Error Resume Next
        scriptSheet.Buttons("ResponseOption" & i).Delete
        On Error GoTo 0
    Next i
End Sub

' Add a response option button to the script sheet
Private Sub AddResponseOption(scriptSheet As Worksheet, optionNumber As Integer, optionText As String)
On Error Resume Next
    Dim responseBtn As Button
    Dim topRow As Integer
    
    ' Calculate position based on option number
    topRow = 22 + ((optionNumber - 1) \ 3) * 4 ' 3 buttons per row
    Dim columnIndex As Integer
    columnIndex = ((optionNumber - 1) Mod 3) * 3 + 2 ' B, E, or H columns
    
    ' Add the button
    Set responseBtn = scriptSheet.Buttons.Add(scriptSheet.Cells(topRow, columnIndex).left, _
                                 scriptSheet.Cells(topRow, columnIndex).top, _
                                 120, 25)
    With responseBtn
        .caption = optionText
        .Name = "ResponseOption" & optionNumber
        .OnAction = "ScriptOptionClicked"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
End Sub

' Handle script option button click
Public Sub ScriptOptionClicked()
On Error Resume Next
    Dim buttonName As String
    Dim optionText As String
    
    ' Get button name
    buttonName = Application.Caller
    
    ' Get button caption (option text)
    optionText = ThisWorkbook.Sheets("DynamicScript").Buttons(buttonName).caption
    
    ' Handle the option selection
    HandleScriptOptionSelection optionText
End Sub

' Log customer contact to history
Private Sub LogCustomerContact(customerName As String, contactType As String, details As String, contactDate As Date)
On Error Resume Next
    Dim historySheet As Worksheet
    Dim nextRow As Long
    
    ' Ensure we have a ContactHistory sheet
    On Error Resume Next
    Set historySheet = ThisWorkbook.Sheets("ContactHistory")
    On Error GoTo 0
    
    If historySheet Is Nothing Then
        ' Create the history sheet
        Set historySheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        historySheet.Name = "ContactHistory"
        
        ' Style the sheet
        historySheet.Tab.Color = COLOR_PRIMARY
        
        ' Add headers
        With historySheet.Range("A1:E1")
            .Value = Array("Customer", "Contact Type", "Details", "Date/Time", "User")
            .Font.Bold = True
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Interior.Color = COLOR_PRIMARY
            .Font.Color = COLOR_TEXT_LIGHT
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .HorizontalAlignment = xlCenter
        End With
        
        ' Autofit columns
        historySheet.Columns("A:E").AutoFit
        
        ' Set column widths
        historySheet.Columns("C:C").ColumnWidth = 50 ' Details column wider
    End If
    
    ' Find next empty row
    nextRow = historySheet.Cells(historySheet.Rows.count, "A").End(xlUp).row + 1
    
    ' Add contact record
    historySheet.Cells(nextRow, 1).Value = customerName
    historySheet.Cells(nextRow, 2).Value = contactType
    historySheet.Cells(nextRow, 3).Value = details
    historySheet.Cells(nextRow, 4).Value = contactDate
    historySheet.Cells(nextRow, 5).Value = Application.userName
    
    ' Format the new row
    With historySheet.Range(historySheet.Cells(nextRow, 1), historySheet.Cells(nextRow, 5))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Alternate row colors for readability
    If nextRow Mod 2 = 0 Then
        historySheet.Range(historySheet.Cells(nextRow, 1), historySheet.Cells(nextRow, 5)).Interior.Color = RGB(240, 240, 240)
    End If
End Sub


