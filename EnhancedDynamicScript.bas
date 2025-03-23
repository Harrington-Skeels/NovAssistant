Attribute VB_Name = "EnhancedDynamicScript"
' ====================================================================
' EnhancedDynamicScript Module - Version 2.0
' ====================================================================
' This module creates an advanced interactive script system that guides users
' through the novated leasing sales process with intelligent conversation paths,
' customer profiling, and automated follow-up actions

Option Explicit

' Script navigation states
Private currentScriptView As String
Private scriptHistory As Collection
Private customerResponses As Object ' Dictionary
Private scriptStartTime As Date
Private activeCustomerName As String
Private activeCustomerPhone As String
Private scriptNotes As String
Private activeCustomerEmail As String
Private activeCustomerStage As String
Private activeCustomerID As String
Private scriptSuggestions As String

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
Private Const VIEW_QUOTE_FOLLOW_UP = "QuoteFollowUpView" ' New follow-up view
Private Const VIEW_APPLICATION = "ApplicationView" ' New application view
Private Const VIEW_SETTLEMENT = "SettlementView" ' New settlement view

' Colors for consistent styling (matches dashboard)
Private Const COLOR_PRIMARY_DARK = 3368601  ' Darker green (RGB 1, 51, 32)
Private Const COLOR_PRIMARY = 4227072       ' Dark green (RGB 0, 66, 37)
Private Const COLOR_PRIMARY_LIGHT = 5287936 ' Light green (RGB 80, 160, 80)
Private Const COLOR_SECONDARY = 39423       ' Orange (RGB 255, 153, 0)
Private Const COLOR_ACCENT = 49344          ' Gold (RGB 192, 192, 0)
Private Const COLOR_TEXT_LIGHT = 16777215   ' White
Private Const COLOR_BACKGROUND_DARK = 15132390 ' Dark gray (RGB 230, 230, 230)
Private Const COLOR_BACKGROUND = 15921906   ' Light gray (RGB 242, 242, 242)
Private Const COLOR_SUCCESS = 5287936       ' Green (RGB 80, 160, 80)
Private Const COLOR_WARNING = 49344         ' Orange (RGB 192, 192, 0)
Private Const COLOR_DANGER = 5459839        ' Red (RGB 83, 83, 255)
Private Const COLOR_INFO = 16764108         ' Light blue (RGB 204, 236, 255)

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
    
    ' Activate script sheet
    scriptSheet.Activate
End Sub

' Set up the script layout with improved visual design
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
        .Font.Size = 18
        .Font.Bold = True
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY_DARK
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
    
    ' Stage row - new
    With scriptSheet.Range("B6:I6")
        .Interior.Color = COLOR_SECONDARY
        .Font.Color = COLOR_TEXT_LIGHT
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    With scriptSheet.Range("B6:D6")
        .Merge
        .Value = "Current Stage:"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
    With scriptSheet.Range("E6:I6")
        .Merge
        .Value = ""
        .Interior.Color = COLOR_BACKGROUND
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
    End With
    
    ' Script content area - header
    With scriptSheet.Range("B8:I8")
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
    With scriptSheet.Range("B9:I19")
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
    With scriptSheet.Range("B20:I20")
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
    With scriptSheet.Range("B21:I29")
        .Interior.Color = RGB(240, 240, 240)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = COLOR_PRIMARY
    End With
    
    ' Suggestions section - new
    With scriptSheet.Range("B30:I30")
        .Merge
        .Value = "SUGGESTIONS:"
        .Font.Bold = True
        .Font.Size = 12
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
    End With
    
    ' Suggestions area - content
    With scriptSheet.Range("B31:I33")
        .Merge
        .Value = ""
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Interior.Color = RGB(255, 255, 240)  ' Very light yellow
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = COLOR_PRIMARY
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Italic = True
    End With
    
    ' Notes area - header
    With scriptSheet.Range("B34:I34")
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
    With scriptSheet.Range("B35:I41")
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
    With scriptSheet.Range("B42:I42")
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
        .OnAction = "EnhancedDynamicScript.ScriptNavigateBack"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
End Sub

' Add action buttons to the script sheet
Private Sub AddActionButtons(scriptSheet As Worksheet)
On Error Resume Next
    Dim startBtn As Button, endBtn As Button, followUpBtn As Button, saveNoteBtn As Button
    Dim updateStageBtn As Button, emailBtn As Button
    
    ' Remove existing buttons if any
    On Error Resume Next
    scriptSheet.Buttons("StartCallButton").Delete
    scriptSheet.Buttons("EndCallButton").Delete
    scriptSheet.Buttons("FollowUpButton").Delete
    scriptSheet.Buttons("SaveNoteButton").Delete
    scriptSheet.Buttons("UpdateStageButton").Delete
    scriptSheet.Buttons("SendEmailButton").Delete
    On Error GoTo 0
    
    ' Calculate button positions
    Dim btnTop As Double
    Dim btnHeight As Double
    Dim btnWidth As Double
    Dim btnSpacing As Double
    
    btnTop = scriptSheet.Range("B43").top
    btnHeight = 28
    btnWidth = 78
    btnSpacing = 5
    
    ' Start call button
    Set startBtn = scriptSheet.Buttons.Add(scriptSheet.Range("B43").left, _
                                 btnTop, btnWidth, btnHeight)
    With startBtn
        .caption = "Start New Call"
        .Name = "StartCallButton"
        .OnAction = "EnhancedDynamicScript.StartNewScript"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    ' End call button
    Set endBtn = scriptSheet.Buttons.Add(scriptSheet.Range("B43").left + btnWidth + btnSpacing, _
                                 btnTop, btnWidth, btnHeight)
    With endBtn
        .caption = "End Call"
        .Name = "EndCallButton"
        .OnAction = "EnhancedDynamicScript.EndCurrentScript"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    ' Save note button
    Set saveNoteBtn = scriptSheet.Buttons.Add(scriptSheet.Range("B43").left + (btnWidth + btnSpacing) * 2, _
                                 btnTop, btnWidth, btnHeight)
    With saveNoteBtn
        .caption = "Save Notes"
        .Name = "SaveNoteButton"
        .OnAction = "EnhancedDynamicScript.SaveScriptNotes"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    ' Update stage button - new
    Set updateStageBtn = scriptSheet.Buttons.Add(scriptSheet.Range("B43").left + (btnWidth + btnSpacing) * 3, _
                                 btnTop, btnWidth, btnHeight)
    With updateStageBtn
        .caption = "Update Stage"
        .Name = "UpdateStageButton"
        .OnAction = "EnhancedDynamicScript.UpdateCustomerStageFromScript"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    ' Schedule follow-up button
    Set followUpBtn = scriptSheet.Buttons.Add(scriptSheet.Range("B43").left + (btnWidth + btnSpacing) * 4, _
                                 btnTop, btnWidth, btnHeight)
    With followUpBtn
        .caption = "Schedule Follow-up"
        .Name = "FollowUpButton"
        .OnAction = "EnhancedDynamicScript.ScheduleScriptFollowUp"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
    
    ' Send email button - new
    Set emailBtn = scriptSheet.Buttons.Add(scriptSheet.Range("B43").left + (btnWidth + btnSpacing) * 5, _
                                 btnTop, btnWidth, btnHeight)
    With emailBtn
        .caption = "Send Email"
        .Name = "SendEmailButton"
        .OnAction = "EnhancedDynamicScript.SendEmailFromScript"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
    End With
End Sub

' Start a new script session with a customer
Public Sub StartNewScript()
On Error Resume Next
    Dim customerName As String
    Dim customerPhone As String
    Dim customerEmail As String
    Dim customerStage As String
    Dim customerID As String
    Dim callSheet As Worksheet
    Dim customerSheet As Worksheet
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
        
        ' Try to get additional customer info from tracker
        Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
        Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
        
        If Not customerRow Is Nothing Then
            customerEmail = customerRow.Offset(0, 1).Value ' Email column
            customerStage = customerRow.Offset(0, 3).Value ' Stage column
            customerID = customerRow.Offset(0, -1).Value  ' ID column
        End If
    Else
        ' Prompt for customer details
        customerName = InputBox("Enter customer name:", "Start New Call")
        If customerName = "" Then Exit Sub
        
        customerPhone = InputBox("Enter customer phone number:", "Start New Call")
        
        ' Try to get additional customer info from tracker
        Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
        Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
        
        If Not customerRow Is Nothing Then
            customerEmail = customerRow.Offset(0, 1).Value ' Email column
            customerStage = customerRow.Offset(0, 3).Value ' Stage column
            customerID = customerRow.Offset(0, -1).Value  ' ID column
        End If
    End If
    
    ' Initialize script state
    Set scriptHistory = New Collection
    Set customerResponses = CreateObject("Scripting.Dictionary")
    
    ' Determine appropriate starting view based on customer stage
    If customerStage = "Quote Sent" Then
        currentScriptView = VIEW_QUOTE_FOLLOW_UP
    ElseIf customerStage = "Finance Application" Then
        currentScriptView = VIEW_APPLICATION
    ElseIf customerStage = "Settlement" Then
        currentScriptView = VIEW_SETTLEMENT
    Else
        currentScriptView = VIEW_INITIAL
    End If
    
    ' Set active customer info
    scriptStartTime = Now
    activeCustomerName = customerName
    activeCustomerPhone = customerPhone
    activeCustomerEmail = customerEmail
    activeCustomerStage = customerStage
    activeCustomerID = customerID
    scriptNotes = ""
    
    ' Clear suggestions
    scriptSuggestions = ""
    
    ' Navigate to the script sheet
    ThisWorkbook.Sheets("DynamicScript").Activate
    
    ' Update customer info
    With ThisWorkbook.Sheets("DynamicScript")
        .Range("E3").Value = customerName
        .Range("E4").Value = customerPhone
        .Range("E6").Value = IIf(customerStage = "", "New Lead", customerStage)
    End With
    
    ' Start the timer
    Application.OnTime Now + TimeValue("00:00:01"), "EnhancedDynamicScript.UpdateScriptTimer"
    
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
    Application.OnTime Now + TimeValue("00:00:01"), "EnhancedDynamicScript.UpdateScriptTimer"
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
                Case "Knowledgeable"
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
            
        Case VIEW_QUOTE_FOLLOW_UP
            ' After quote follow-up, handle customer decision
            Select Case optionValue
                Case "Ready to Proceed"
                    currentScriptView = VIEW_APPLICATION
                Case "Needs More Time"
                    currentScriptView = VIEW_OBJECTIONS
                Case "Not Interested"
                    currentScriptView = VIEW_CLOSING
                Case Else
                    currentScriptView = VIEW_GATHER_DETAILS
            End Select
            
        Case VIEW_APPLICATION
            ' After application, move to next steps
            Select Case optionValue
                Case "Application Complete"
                    currentScriptView = VIEW_SETTLEMENT
                Case "Application Issues"
                    currentScriptView = VIEW_OBJECTIONS
                Case Else
                    currentScriptView = VIEW_CLOSING
            End Select
            
        Case VIEW_SETTLEMENT
            ' After settlement discussion, go to closing
            currentScriptView = VIEW_CLOSING
            
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

' End the current script session
Public Sub EndCurrentScript()
On Error Resume Next
    Dim outcome As String
    Dim callDuration As Date
    Dim callSheet As Worksheet
    Dim callRow As Range
    Dim stageChanged As Boolean
    Dim newStage As String
    
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
    
    ' Ask if customer stage needs updating
    stageChanged = (MsgBox("Do you want to update the customer's stage?", vbQuestion + vbYesNo) = vbYes)
    
    If stageChanged Then
        ' Show list of available stages
        Dim stages As String
        stages = "Initial Call" & vbCrLf & _
                "Quote Sent" & vbCrLf & _
                "Finance Application" & vbCrLf & _
                "Settlement"
        
        newStage = InputBox("Select new customer stage:" & vbCrLf & vbCrLf & stages, "Update Stage", activeCustomerStage)
        
        If newStage <> "" And newStage <> activeCustomerStage Then
            ' Update customer stage
            UpdateCustomerStage activeCustomerName, newStage, "Updated stage during call"
        End If
    End If
    
    ' Get notes from the script sheet
    scriptNotes = ThisWorkbook.Sheets("DynamicScript").Range("B35").Value
    
    ' Log call end
    LogCustomerContact activeCustomerName, "Outbound Call", _
        "Call completed (" & Format(callDuration, "hh:mm:ss") & ") - " & outcome & vbCrLf & _
        "Notes: " & scriptNotes, Now()
    
    ' Update call planner if this was a scheduled call
    Set callSheet = ThisWorkbook.Sheets("CallPlanner")
    Set callRow = callSheet.Range("B:B").Find(activeCustomerName, LookIn:=xlValues)
    
    If Not callRow Is Nothing Then
        callRow.Offset(0, 5).Value = outcome ' Outcome column
        callRow.Offset(0, 6).Value = "Completed" ' Status column
    End If
    
    ' Ask about follow-up if appropriate
    If outcome <> "Not Interested" Then
        If MsgBox("Would you like to schedule a follow-up?", vbQuestion + vbYesNo) = vbYes Then
            ScheduleScriptFollowUp
        End If
    End If
    
    ' Stop the timer
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "EnhancedDynamicScript.UpdateScriptTimer", , False
    On Error GoTo 0
    
    ' Reset values
    activeCustomerName = ""
    activeCustomerPhone = ""
    activeCustomerEmail = ""
    activeCustomerStage = ""
    activeCustomerID = ""
    scriptStartTime = 0
    
    ' Clear the script display
    With ThisWorkbook.Sheets("DynamicScript")
        .Range("E3").Value = ""
        .Range("E4").Value = ""
        .Range("E5").Value = "00:00:00"
        .Range("E6").Value = ""
        .Range("B9").Value = "Press 'Start New Call' to begin"
        .Range("B31").Value = "" ' Clear suggestions
        .Range("B35").Value = ""
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
    notes = ThisWorkbook.Sheets("DynamicScript").Range("B35").Value
    
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
    Dim followupTime As String
    Dim notes As String
    
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
    
    ' Prompt for time
    followupTime = InputBox("Enter follow-up time (HH:MM AM/PM):", "Schedule Follow-up", "10:00 AM")
    
    ' Suggest follow-up type based on current stage
    Dim suggestedType As String
    
    Select Case activeCustomerStage
        Case "Initial Call"
            suggestedType = "Quote Follow-up"
        Case "Quote Sent"
            suggestedType = "Decision Follow-up"
        Case "Finance Application"
            suggestedType = "Application Status Check"
        Case Else
            suggestedType = "Follow-up call"
    End Select
    
    followupType = InputBox("Enter follow-up type:", "Schedule Follow-up", suggestedType)
    
    ' Get any notes for the follow-up
    notes = InputBox("Enter any notes for this follow-up:", "Schedule Follow-up")
    
    ' Schedule the follow-up in Excel
    If ScheduleCustomerFollowUp(activeCustomerName, followupType, followupDate) Then
        ' Add to call planner
        AddToCallPlanner activeCustomerName, followupType, followupDate, followupTime
        
        ' Create Outlook appointment if integration is available
        If SystemIntegrationAvailable("Outlook") Then
            On Error Resume Next
            Dim fullDateTime As Date
            fullDateTime = DateAdd("h", Val(left(followupTime, 2)), followupDate)
            fullDateTime = DateAdd("n", Val(Mid(followupTime, 4, 2)), fullDateTime)
            
            ' Try to create an Outlook appointment
            If ScheduleFollowUpEnhanced(activeCustomerName, activeCustomerPhone, fullDateTime, 30, followupType, notes) Then
                ' Success
            End If
            On Error GoTo 0
        End If
        
        ' Confirm scheduling
        MsgBox "Follow-up scheduled for " & Format(followupDate, "mm/dd/yyyy") & " at " & followupTime & ".", vbInformation
    Else
        MsgBox "Failed to schedule follow-up. Please check customer details.", vbExclamation
    End If
End Sub

' Add follow-up to call planner
Private Sub AddToCallPlanner(customerName As String, followupType As String, followupDate As Date, followupTime As String)
On Error Resume Next
    Dim callSheet As Worksheet
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    Dim nextRow As Long
    
    ' Get sheets
    Set callSheet = ThisWorkbook.Sheets("CallPlanner")
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    
    ' Find customer in tracker to get additional info
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If customerRow Is Nothing Then Exit Sub
    
    ' Find next empty row in call planner
    nextRow = callSheet.Cells(callSheet.Rows.count, "A").End(xlUp).row + 1
    
    ' Add to call planner
    callSheet.Cells(nextRow, 1).Value = followupTime ' Time
    callSheet.Cells(nextRow, 2).Value = customerName
    callSheet.Cells(nextRow, 3).Value = customerRow.Offset(0, 2).Value ' Phone
    callSheet.Cells(nextRow, 4).Value = followupType
    callSheet.Cells(nextRow, 5).Value = customerRow.Offset(0, 3).Value ' Stage
    callSheet.Cells(nextRow, 6).Value = customerRow.Offset(0, 12).Value ' Status
    callSheet.Cells(nextRow, 7).Value = "Pending"
    callSheet.Cells(nextRow, 8).Value = followupDate ' Date of follow-up
End Sub

' Schedule customer follow-up in tracker
Private Function ScheduleCustomerFollowUp(customerName As String, actionType As String, actionDate As Date) As Boolean
On Error Resume Next
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If Not customerRow Is Nothing Then
        customerRow.Offset(0, 6).Value = actionType ' Next Action column
        customerRow.Offset(0, 7).Value = actionDate ' Next Action Date column
        
        ' Log in history
        LogCustomerContact customerName, "Follow-up Scheduled", actionType & " scheduled for " & Format(actionDate, "mm/dd/yyyy"), Now()
        
        ScheduleCustomerFollowUp = True
    Else
        ScheduleCustomerFollowUp = False
    End If
End Function

' Update the script view based on the current state
Private Sub UpdateScriptView()
On Error Resume Next
    Dim scriptSheet As Worksheet
    Dim scriptText As String
    Dim pathText As String
    
    Set scriptSheet = ThisWorkbook.Sheets("DynamicScript")
    
    ' Clear previous response options
    ClearResponseButtons scriptSheet
    
    ' Set the script text, response options, and suggestions based on the current view
    Select Case currentScriptView
        Case VIEW_INITIAL
            ' Initial script - introduction
            scriptText = "Hi " & activeCustomerName & ", my name is " & Application.userName & ". Before I start, just a reminder that our calls are recorded for training purposes." & vbCrLf & vbCrLf & _
                         "I understand you are calling about a novated lease?" & vbCrLf & vbCrLf & _
                         "How familiar are you with novated leasing?"

            pathText = "Initial Assessment"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "First Timer"
            AddResponseOption scriptSheet, 2, "Knows a little"
            AddResponseOption scriptSheet, 3, "Knowledgeable"
            AddResponseOption scriptSheet, 4, "Skip to Qualifying"
            
            ' Set suggestions
            scriptSuggestions = "• For new customers, explain the basic concept clearly" & vbCrLf & _
                               "• Listen for specific car interests or budget concerns" & vbCrLf & _
                               "• Note their current vehicle and payment method"
            
        Case VIEW_NOT_MUCH
            ' Not much knowledge script
            scriptText = "That's perfectly fine - many people are new to novated leasing. Let me explain the basics:" & vbCrLf & vbCrLf & _
                         "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf & _
                         "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars."

            pathText = "Initial Assessment > Not Much Knowledge"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Qualifying"
            
            ' Set suggestions
            scriptSuggestions = "• Use simple terms and avoid technical jargon" & vbCrLf & _
                               "• Check for understanding frequently" & vbCrLf & _
                               "• Emphasize the tax benefits and simplicity"
            
        Case VIEW_A_LITTLE
            ' A little knowledge script
            scriptText = "Great, so you have some familiarity with novated leasing. Let me confirm and expand on what you might already know:" & vbCrLf & vbCrLf & _
                         "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf & _
                         "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars."
            
            pathText = "Initial Assessment > Some Knowledge"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Qualifying"
            
            ' Set suggestions
            scriptSuggestions = "• Focus on filling knowledge gaps" & vbCrLf & _
                               "• Ask what aspects they're most interested in" & vbCrLf & _
                               "• Address any misconceptions they might have"
            
        Case VIEW_WELL_EDUCATED
            ' Well educated script
            scriptText = "Excellent! Since you're already familiar with novated leasing, let's focus on the specific aspects that would be most beneficial for your situation." & vbCrLf & vbCrLf & _
                         "I'd like to ask you some specific questions about your needs to tailor our discussion to what would work best for you."
            
            pathText = "Initial Assessment > Well Educated"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Qualifying"
            
            ' Set suggestions
            scriptSuggestions = "• Focus on more advanced benefits" & vbCrLf & _
                               "• Discuss specific tax implications" & vbCrLf & _
                               "• Explore customization options for their needs"
            
        Case VIEW_QUALIFYING
            ' Qualifying questions script
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
            
            ' Set suggestions
            scriptSuggestions = "• Note vehicle preferences and budget range" & vbCrLf & _
                               "• Listen for purchase timeline indicators" & vbCrLf & _
                               "• Identify potential objections early"
            
        Case VIEW_EDUCATING
            ' Educating the customer script
            scriptText = "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf & _
                         "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars." & vbCrLf & vbCrLf & _
                         "The novated lease will include; the car, services, maintenance, tyres, rego, insurance and fuel. We will set a budget that reflects how many kilometres you drive each year." & vbCrLf & vbCrLf & _
                         "Getting your budget correct is important with a Novated Lease, if we under estimate your running costs the lease will look cheap and attractive, however you won't have money to cover all the running costs. If we over-budget, it will look too expensive and you won't want to take the lease."
            
            pathText = "Education > Basics"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Benefits"
            AddResponseOption scriptSheet, 2, "Customer Has Questions"
            
            ' Set suggestions
            scriptSuggestions = "• Tailor explanation to their knowledge level" & vbCrLf & _
                               "• Use examples relevant to their situation" & vbCrLf & _
                               "• Check for understanding before proceeding"
            
        Case VIEW_BENEFITS
            ' Benefits script
            scriptText = "The vehicle is also fully maintained, meaning you'll get fuel cards for your fuel, the dealer will invoice us for your servicing, same with the tyre shop. Rego renewals can be uploaded, and we'll pay them directly." & vbCrLf & vbCrLf & _
                         "So other than tolls and fines you should not need to outlay any further money on your vehicle outside of those deductions." & vbCrLf & vbCrLf & _
                         "The easy way to think of this, is that the deductions will go to an account held at sgfleet. As you spend money on the vehicle, it will deduct from that account and you can manage and monitor this through our online app." & vbCrLf & vbCrLf & _
                         "Throughout the lease you're paying the vehicle down to a residual or balloon amount set by the ATO."
            
            pathText = "Education > Benefits"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Lease End Options"
            AddResponseOption scriptSheet, 2, "Customer Has Questions"
            
            ' Set suggestions
            scriptSuggestions = "• Emphasize convenience and budgeting benefits" & vbCrLf & _
                               "• Mention the mobile app for expense tracking" & vbCrLf & _
                               "• Stress the GST benefits on all running costs"
            
        Case VIEW_LEASE_END
            ' Lease end options script
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
            
            ' Set suggestions
            scriptSuggestions = "• Highlight flexibility at lease end" & vbCrLf & _
                               "• Focus on the most relevant option for this customer" & vbCrLf & _
                               "• Mention the fleet discounts for new vehicles"
            
        Case VIEW_TRADE_IN
            ' Trade-in information script
            scriptText = "Do you have a car you are looking to trade in or sell?" & vbCrLf & vbCrLf & _
                         "We have a trade advantage program where we will have several wholesalers value your existing car. If you're happy with the price, you can trade it in and we will even come and collect the car for you." & vbCrLf & vbCrLf & _
                         "Do you have any questions at this point?"
            
            pathText = "Trade-In Information"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Gather Details"
            AddResponseOption scriptSheet, 2, "Customer Has Questions"
            
            ' Set suggestions
            scriptSuggestions = "• Get details about their current vehicle" & vbCrLf & _
                               "• Explain the valuation process" & vbCrLf & _
                               "• Emphasize the convenience of the trade-in service"
            
        Case VIEW_QUOTE_FOLLOW_UP
            ' Quote follow-up script (new view)
            scriptText = "Hi " & activeCustomerName & ", I'm following up regarding the novated lease quote we sent you recently." & vbCrLf & vbCrLf & _
                         "Have you had a chance to review the quote? What did you think of it?" & vbCrLf & vbCrLf & _
                         "Do you have any questions about the quote or the next steps?"
            
            pathText = "Quote Follow-Up"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Ready to Proceed"
            AddResponseOption scriptSheet, 2, "Needs More Time"
            AddResponseOption scriptSheet, 3, "Has Questions"
            AddResponseOption scriptSheet, 4, "Not Interested"
            
            ' Set suggestions
            scriptSuggestions = "• Review quote details before the call" & vbCrLf & _
                               "• Ask open-ended questions about their thoughts" & vbCrLf & _
                               "• Be prepared to address common concerns" & vbCrLf & _
                               "• If interested, move to finance application steps"
            
        Case VIEW_APPLICATION
            ' Application script (new view)
            scriptText = "Great! To proceed with your novated lease, we'll need to complete a finance application." & vbCrLf & vbCrLf & _
                         "I can guide you through this process - it typically takes about 15-20 minutes to complete online." & vbCrLf & vbCrLf & _
                         "You'll need to have your driver's license, employment details, and income information ready." & vbCrLf & vbCrLf & _
                         "Would you like to complete this now, or would you prefer me to email you the link so you can complete it at your convenience?"
            
            pathText = "Finance Application"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Complete Now"
            AddResponseOption scriptSheet, 2, "Send Email Link"
            AddResponseOption scriptSheet, 3, "Application Issues"
            AddResponseOption scriptSheet, 4, "Application Complete"
            
            ' Set suggestions
            scriptSuggestions = "• Explain all required documentation clearly" & vbCrLf & _
                               "• Offer to stay on the line during the application" & vbCrLf & _
                               "• Explain the expected timeline for approval" & vbCrLf & _
                               "• Set clear next steps after application submission"
            
        Case VIEW_SETTLEMENT
            ' Settlement script (new view)
            scriptText = "Great news! Your finance application has been approved, and we're now ready to move to the settlement phase." & vbCrLf & vbCrLf & _
                         "Here's what happens next:" & vbCrLf & _
                         "1. We'll finalize the vehicle order with the dealership" & vbCrLf & _
                         "2. Once the vehicle is ready, we'll arrange delivery" & vbCrLf & _
                         "3. Your salary packaging will be set up with your employer" & vbCrLf & _
                         "4. You'll receive your fuel cards and welcome pack" & vbCrLf & vbCrLf & _
                         "Do you have any questions about this process or timeline?"
            
            pathText = "Settlement"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Continue to Closing"
            AddResponseOption scriptSheet, 2, "Has Questions"
            
            ' Set suggestions
            scriptSuggestions = "• Confirm vehicle selection and delivery preferences" & vbCrLf & _
                               "• Explain the employer documentation process" & vbCrLf & _
                               "• Set clear expectations about delivery timeline" & vbCrLf & _
                               "• Mention the introduction to the service team"
            
        Case VIEW_OBJECTIONS
            ' Handling objections script
            scriptText = "I understand your concerns. Many people have similar questions when considering a novated lease." & vbCrLf & vbCrLf & _
                         "Common questions and our responses:" & vbCrLf & vbCrLf & _
                         "1. ""It seems expensive"" - While there is an upfront cost, the tax benefits and GST savings often make it more cost-effective than traditional car ownership over time." & vbCrLf & vbCrLf & _
                         "2. ""I'm worried about changing jobs"" - If you change employers, you have options: transfer the lease to your new employer, convert to a consumer loan, or pay it out." & vbCrLf & vbCrLf & _
                         "3. ""I don't understand the tax benefits"" - We can prepare a personalized calculation showing exactly how much you could save based on your salary and the vehicle you're interested in."
            
            pathText = "Handling Objections"
            
            ' Add response options
            AddResponseOption scriptSheet, 1, "Objection Resolved"
            AddResponseOption scriptSheet, 2, "Not Interested"
            
            ' Set suggestions
            scriptSuggestions = "• Listen carefully to specific concerns" & vbCrLf & _
                               "• Address objections with specific benefits" & vbCrLf & _
                               "• Offer to provide personalized calculations" & vbCrLf & _
                               "• Be prepared with competitor comparisons"
            
        Case VIEW_GATHER_DETAILS
            ' Gather details script
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
            
            ' Set suggestions
            scriptSuggestions = "• Record all vehicle specifications accurately" & vbCrLf & _
                               "• Note preferred lease term and kilometers" & vbCrLf & _
                               "• Confirm income for finance pre-approval" & vbCrLf & _
                               "• Ask about any special requirements"
            
        Case VIEW_CLOSING
            ' Closing script
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
            
            ' Set suggestions
            scriptSuggestions = "• Summarize next steps clearly" & vbCrLf & _
                               "• Confirm when you'll send the quote" & vbCrLf & _
                               "• Schedule a specific follow-up time" & vbCrLf & _
                               "• Thank them for their time"
            
        Case Else
            ' Default view
            scriptText = "Press 'Start New Call' to begin a new script session."
            pathText = "Initial Assessment"
    End Select
    
    ' Update the script sheet
    scriptSheet.Range("B2").Value = "Current Path: " & pathText
    scriptSheet.Range("B9").Value = scriptText
    scriptSheet.Range("B31").Value = scriptSuggestions
    
    ' If no active customer, clear customer info fields
    If activeCustomerName = "" Then
        scriptSheet.Range("E3").Value = ""
        scriptSheet.Range("E4").Value = ""
        scriptSheet.Range("E5").Value = "00:00:00"
        scriptSheet.Range("E6").Value = ""
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
        .OnAction = "EnhancedDynamicScript.ScriptOptionClicked"
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

' Update customer stage from script form
Public Sub UpdateCustomerStageFromScript()
On Error Resume Next
    ' Check if we have an active call
    If activeCustomerName = "" Then
        MsgBox "No active call to update stage for.", vbExclamation
        Exit Sub
    End If
    
    ' Show list of available stages
    Dim stages As String
    stages = "Initial Call" & vbCrLf & _
            "Quote Sent" & vbCrLf & _
            "Finance Application" & vbCrLf & _
            "Settlement"
    
    Dim newStage As String
    newStage = InputBox("Select new customer stage:" & vbCrLf & vbCrLf & stages, "Update Stage", activeCustomerStage)
    
    If newStage <> "" And newStage <> activeCustomerStage Then
        ' Update customer stage
        If UpdateCustomerStage(activeCustomerName, newStage, "Updated stage during call") Then
            ' Update stage in script form
            activeCustomerStage = newStage
            ThisWorkbook.Sheets("DynamicScript").Range("E6").Value = newStage
            
            ' Show confirmation
            MsgBox "Customer stage updated to: " & newStage, vbInformation
            
            ' If stage changed to Quote Sent, ask about sending quote
            If newStage = "Quote Sent" And MsgBox("Would you like to send a quote email now?", vbQuestion + vbYesNo) = vbYes Then
                SendEmailFromScript
            End If
        Else
            MsgBox "Failed to update customer stage. Please check customer details.", vbExclamation
        End If
    End If
End Sub

' Send email from script form
Public Sub SendEmailFromScript()
On Error Resume Next
    ' Check if we have an active call
    If activeCustomerName = "" Then
        MsgBox "No active customer to send email to.", vbExclamation
        Exit Sub
    End If
    
    ' Check if we have customer email
    If activeCustomerEmail = "" Then
        activeCustomerEmail = InputBox("Please enter the customer's email address:", "Send Email")
        If activeCustomerEmail = "" Then Exit Sub
        
        ' Update customer record with email
        UpdateCustomerEmail activeCustomerName, activeCustomerEmail
    End If
    
    ' Select email template based on customer stage
    Dim templateName As String
    
    Select Case activeCustomerStage
        Case "Initial Call"
            templateName = "Initial Quote"
        Case "Quote Sent"
            templateName = "Quote Follow-up"
        Case "Finance Application"
            templateName = "Application Status"
        Case "Settlement"
            templateName = "Delivery Information"
        Case Else
            templateName = "General Follow-up"
    End Select
    
    ' Get email subject and body
    Dim emailSubject As String
    Dim emailBody As String
    
    emailSubject = GetEmailTemplate(templateName & " Subject")
    If emailSubject = "" Then
        emailSubject = InputBox("Enter email subject:", "Send Email", "Your Novated Lease Enquiry")
        If emailSubject = "" Then Exit Sub
    End If
    
    emailBody = GetEmailTemplate(templateName & " Body")
    If emailBody = "" Then
        emailBody = InputBox("Enter email message:", "Send Email", _
                           "Dear " & activeCustomerName & "," & vbCrLf & vbCrLf & _
                           "Thank you for your interest in a novated lease. " & _
                           "Please let me know if you have any questions." & vbCrLf & vbCrLf & _
                           "Best regards," & vbCrLf & _
                           Application.userName)
        If emailBody = "" Then Exit Sub
    Else
        ' Replace placeholders in template
        emailBody = Replace(emailBody, "[CustomerName]", activeCustomerName)
        emailBody = Replace(emailBody, "[UserName]", Application.userName)
    End If
    
    ' Send email
    If SystemIntegrationAvailable("Outlook") Then
        If SendEnhancedEmail(activeCustomerName, activeCustomerEmail, emailSubject, emailBody) Then
            MsgBox "Email sent to " & activeCustomerName & ".", vbInformation
        Else
            MsgBox "Failed to send email. Please check your Outlook connection.", vbExclamation
        End If
    Else
        MsgBox "Outlook is not available. Email could not be sent.", vbExclamation
    End If
End Sub

' Update customer email address
Private Sub UpdateCustomerEmail(customerName As String, email As String)
On Error Resume Next
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If Not customerRow Is Nothing Then
        customerRow.Offset(0, 1).Value = email ' Email column
    End If
End Sub

' Get email template by name
Private Function GetEmailTemplate(templateName As String) As String
On Error Resume Next
    Dim templatesSheet As Worksheet
    Dim templateRow As Range
    
    ' Check if Templates sheet exists
    On Error Resume Next
    Set templatesSheet = ThisWorkbook.Sheets("Templates")
    On Error GoTo 0
    
    If templatesSheet Is Nothing Then
        GetEmailTemplate = ""
        Exit Function
    End If
    
    ' Find template
    Set templateRow = templatesSheet.Range("B:B").Find(templateName, LookIn:=xlValues)
    
    If templateRow Is Nothing Then
        GetEmailTemplate = ""
    Else
        ' Return template content (column D)
        GetEmailTemplate = templateRow.Offset(0, 2).Value
    End If
End Function

' Check if system integration is available
Private Function SystemIntegrationAvailable(integrationType As String) As Boolean
On Error Resume Next
    Select Case integrationType
        Case "Outlook"
            ' Try to get Outlook
            Dim objOutlook As Object
            On Error Resume Next
            Set objOutlook = GetObject(, "Outlook.Application")
            If objOutlook Is Nothing Then
                Set objOutlook = CreateObject("Outlook.Application")
            End If
            SystemIntegrationAvailable = Not (objOutlook Is Nothing)
            Set objOutlook = Nothing
            
        Case "Dynamics"
            ' Check for Dynamics CRM connection function
            SystemIntegrationAvailable = True ' Placeholder - replace with actual check
            
        Case Else
            SystemIntegrationAvailable = False
    End Select
End Function

' Update customer stage
Public Function UpdateCustomerStage(customerName As String, newStage As String, Optional notes As String = "") As Boolean
On Error Resume Next
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If customerRow Is Nothing Then
        UpdateCustomerStage = False
        Exit Function
    End If
    
    ' Update Excel
    Dim oldStage As String
    oldStage = customerRow.Offset(0, 3).Value
    customerRow.Offset(0, 3).Value = newStage ' Stage column
    customerRow.Offset(0, 4).Value = Date ' Last Contact column
    
    ' Add stage change to notes
    If Not IsEmpty(customerRow.Offset(0, 9).Value) Then ' Notes column
        customerRow.Offset(0, 9).Value = customerRow.Offset(0, 9).Value & vbCrLf & _
            Format(Now, "yyyy-mm-dd hh:mm") & " - Stage changed from " & oldStage & " to " & newStage
        
        If notes <> "" Then
            customerRow.Offset(0, 9).Value = customerRow.Offset(0, 9).Value & vbCrLf & _
                Format(Now, "yyyy-mm-dd hh:mm") & " - " & notes
        End If
    Else
        customerRow.Offset(0, 9).Value = Format(Now, "yyyy-mm-dd hh:mm") & " - Stage changed from " & oldStage & " to " & newStage
        
        If notes <> "" Then
            customerRow.Offset(0, 9).Value = customerRow.Offset(0, 9).Value & vbCrLf & _
                Format(Now, "yyyy-mm-dd hh:mm") & " - " & notes
        End If
    End If
    
    ' Log in history
    LogCustomerContact customerName, "Stage Change", "Changed from " & oldStage & " to " & newStage, Now()
    
    UpdateCustomerStage = True
End Function

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

' Function to integrate with the enhanced follow-up system
' This is a connection point to the SystemIntegration module
Public Function ScheduleFollowUpEnhanced(customerName As String, customerPhone As String, appointmentDate As Date, durationMinutes As Integer, actionType As String, Optional notes As String = "") As Boolean
On Error Resume Next
    ' Check if the enhanced integration function exists
    Dim hasEnhancedIntegration As Boolean
    
    On Error Resume Next
    Application.Run "SystemIntegration.ScheduleFollowUpEnhanced", customerName, customerPhone, appointmentDate, durationMinutes, actionType, notes
    hasEnhancedIntegration = (Err.Number = 0)
    On Error GoTo 0
    
    If hasEnhancedIntegration Then
        ' Function exists, was already called via Application.Run
        ScheduleFollowUpEnhanced = True
    Else
        ' Fall back to basic scheduling
        Dim objOutlook As Object
        Dim objAppointment As Object
        
        ' Try to get running instance of Outlook
        On Error Resume Next
        Set objOutlook = GetObject(, "Outlook.Application")
        If objOutlook Is Nothing Then
            Set objOutlook = CreateObject("Outlook.Application")
        End If
        
        If objOutlook Is Nothing Then
            ScheduleFollowUpEnhanced = False
            Exit Function
        End If
        
        ' Create appointment
        Set objAppointment = objOutlook.CreateItem(1) ' 1 = olAppointmentItem
        
        With objAppointment
            .subject = "Follow-up with " & customerName
            .Location = "Phone: " & customerPhone
            .Start = appointmentDate
            .duration = durationMinutes
            .ReminderSet = True
            .ReminderMinutesBeforeStart = 15
            .body = "Follow-up call with " & customerName & vbCrLf & vbCrLf & _
                    "Phone: " & customerPhone & vbCrLf & vbCrLf & _
                    "Notes: " & notes
            .Save
        End With
        
        ScheduleFollowUpEnhanced = True
    End If
End Function

' Function to send an email via enhanced integration or basic Outlook
Public Function SendEnhancedEmail(customerName As String, email As String, subject As String, body As String, Optional crmLogging As Boolean = True) As Boolean
On Error Resume Next
    ' Check if the enhanced integration function exists
    Dim hasEnhancedIntegration As Boolean
    
    On Error Resume Next
    Application.Run "SystemIntegration.SendEnhancedEmail", customerName, email, subject, body, crmLogging
    hasEnhancedIntegration = (Err.Number = 0)
    On Error GoTo 0
    
    If hasEnhancedIntegration Then
        ' Function exists, was already called via Application.Run
        SendEnhancedEmail = True
    Else
        ' Fall back to basic email sending
        Dim objOutlook As Object
        Dim objMail As Object
        
        ' Try to get running instance of Outlook
        On Error Resume Next
        Set objOutlook = GetObject(, "Outlook.Application")
        If objOutlook Is Nothing Then
            Set objOutlook = CreateObject("Outlook.Application")
        End If
        
        If objOutlook Is Nothing Then
            SendEnhancedEmail = False
            Exit Function
        End If
        
        ' Create email
        Set objMail = objOutlook.CreateItem(0) ' 0 = olMailItem
        
        With objMail
            .to = email
            .subject = subject
            .HTMLBody = body
            
            ' Display email for review before sending
            .Display
        End With
        
        ' Log in contact history
        LogCustomerContact customerName, "Email Sent", subject, Now()
        
        ' Update last contact date
        UpdateCustomerLastContact customerName
        
        SendEnhancedEmail = True
    End If
End Function
you 're absolutely right about creating a self-contained reminder and task management system directly within Excel. This approach eliminates the need for Outlook integration while giving you more control over notifications and prioritization.

Let me outline a comprehensive solution for implementing this functionality:

# Excel-Based Reminder & Task Management System

## 1. Data Structure Setup

First , we 'll create a structured repository for all reminders and follow-up tasks:

```vba
Sub SetupReminderSystem()
    ' Create a dedicated worksheet for reminders if it doesn't exist
    Dim reminderSheet As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Reminders" Then
            sheetExists = True
            Set reminderSheet = ws
            Exit For
        End If
    Next ws
    
    If Not sheetExists Then
        ' Create new reminders sheet
        Set reminderSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        reminderSheet.Name = "Reminders"
        
        ' Create header row
        With reminderSheet
            .Range("A1").Value = "Customer"
            .Range("B1").Value = "Due Date"
            .Range("C1").Value = "Priority"
            .Range("D1").Value = "Task Type"
            .Range("E1").Value = "Description"
            .Range("F1").Value = "Status"
            .Range("G1").Value = "Created Date"
            .Range("H1").Value = "Created By"
            
            ' Format headers
            .Range("A1:H1").Font.Bold = True
            .Range("A1:H1").Interior.Color = RGB(0, 66, 37) ' SG Fleet green
            .Range("A1:H1").Font.Color = RGB(255, 255, 255) ' White
            
            ' Autofit columns
            .Columns("A:H").AutoFit
            
            ' Add data validation for Status column
            .Range("F2:F1000").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="Pending,In Progress,Completed,Cancelled"
                
            ' Add data validation for Priority column
            .Range("C2:C1000").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="High,Medium,Low"
        End With
    End If
End Sub
```

## 2. Core Reminder Functions

These functions will handle the creation, modification, and completion of reminders:

```vba
Sub CreateReminder(customerName As String, dueDate As Date, taskType As String, _
                  description As String, Optional priority As String = "Medium")
    ' Create a new reminder/task
    Dim reminderSheet As Worksheet
    Dim nextRow As Long
    
    ' Get or create the reminder sheet
    On Error Resume Next
    Set reminderSheet = ThisWorkbook.Sheets("Reminders")
    On Error GoTo 0
    
    If reminderSheet Is Nothing Then
        SetupReminderSystem
        Set reminderSheet = ThisWorkbook.Sheets("Reminders")
    End If
    
    ' Find next empty row
    nextRow = reminderSheet.Cells(reminderSheet.Rows.count, "A").End(xlUp).row + 1
    
    ' Add the reminder
    With reminderSheet
        .Cells(nextRow, 1).Value = customerName
        .Cells(nextRow, 2).Value = dueDate
        .Cells(nextRow, 3).Value = priority
        .Cells(nextRow, 4).Value = taskType
        .Cells(nextRow, 5).Value = description
        .Cells(nextRow, 6).Value = "Pending"
        .Cells(nextRow, 7).Value = Now()
        .Cells(nextRow, 8).Value = Application.userName
        
        ' Format date columns
        .Cells(nextRow, 2).NumberFormat = "dd-mmm-yyyy"
        .Cells(nextRow, 7).NumberFormat = "dd-mmm-yyyy hh:mm"
        
        ' Format priority cell with color
        Select Case priority
            Case "High"
                .Cells(nextRow, 3).Interior.Color = RGB(255, 200, 200) ' Light red
            Case "Medium"
                .Cells(nextRow, 3).Interior.Color = RGB(255, 255, 200) ' Light yellow
            Case "Low"
                .Cells(nextRow, 3).Interior.Color = RGB(200, 255, 200) ' Light green
        End Select
    End With
    
    MsgBox "Reminder created for " & customerName & " due on " & Format(dueDate, "dd-mmm-yyyy"), _
           vbInformation, "Reminder Created"
End Sub

Sub CompleteReminder(reminderRow As Long, Optional completionNotes As String = "")
    ' Mark a reminder as completed
    Dim reminderSheet As Worksheet
    Set reminderSheet = ThisWorkbook.Sheets("Reminders")
    
    With reminderSheet
        ' Update status
        .Cells(reminderRow, 6).Value = "Completed"
        
        ' Add completion notes if provided
        If completionNotes <> "" Then
            .Cells(reminderRow, 5).Value = .Cells(reminderRow, 5).Value & vbCrLf & _
                "Completed: " & Format(Now, "dd-mmm-yyyy hh:mm") & " - " & completionNotes
        End If
        
        ' Format completed row
        .Range(.Cells(reminderRow, 1), .Cells(reminderRow, 8)).Interior.Color = RGB(240, 240, 240)
    End With
End Sub
```

## 3. Reminder Dashboard/Task Manager

This provides a prioritized view of upcoming tasks:

```vba
Sub ShowReminderDashboard()
    ' Display a dashboard of upcoming reminders
    Dim reminderSheet As Worksheet
    Dim dashboardSheet As Worksheet
    Dim sheetExists As Boolean
    
    ' Get reminder sheet
    On Error Resume Next
    Set reminderSheet = ThisWorkbook.Sheets("Reminders")
    On Error GoTo 0
    
    If reminderSheet Is Nothing Then
        MsgBox "No reminders found. Please create reminders first.", vbExclamation
        Exit Sub
    End If
    
    ' Check if dashboard exists
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "ReminderDashboard" Then
            sheetExists = True
            Set dashboardSheet = ws
            Exit For
        End If
    Next ws
    
    If Not sheetExists Then
        ' Create new dashboard sheet
        Set dashboardSheet = ThisWorkbook.Worksheets.Add(After:=reminderSheet)
        dashboardSheet.Name = "ReminderDashboard"
    Else
        ' Clear existing content
        dashboardSheet.Cells.Clear
    End If
    
    ' Set up dashboard
    With dashboardSheet
        ' Create header
        .Range("A1:G1").Value = Array("Customer", "Due Date", "Priority", "Task Type", "Description", "Status", "Actions")
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Interior.Color = RGB(0, 66, 37) ' SG Fleet green
        .Range("A1:G1").Font.Color = RGB(255, 255, 255) ' White
        
        ' Get today's date for highlighting
        Dim today As Date
        today = Date
        
        ' Filter and sort reminders from main sheet
        Dim dataArray As Variant
        Dim filteredData() As Variant
        Dim resultCount As Long
        Dim i As Long, j As Long
        
        ' Get data from reminder sheet (exclude completed/cancelled)
        dataArray = reminderSheet.Range("A2:F" & reminderSheet.Cells(reminderSheet.Rows.count, "A").End(xlUp).row).Value
        
        ' Count non-completed items
        resultCount = 0
        For i = 1 To UBound(dataArray, 1)
            If dataArray(i, 6) <> "Completed" And dataArray(i, 6) <> "Cancelled" Then
                resultCount = resultCount + 1
            End If
        Next i
        
        ' If no active reminders, show message
        If resultCount = 0 Then
            .Range("A2").Value = "No active reminders found."
            .Range("A2:G2").Merge
            .Range("A2").HorizontalAlignment = xlCenter
            .Range("A2").Font.Italic = True
            GoTo FormatColumns
        End If
        
        ' Create array for filtered results
        ReDim filteredData(1 To resultCount, 1 To 6)
        
        ' Fill filtered array
        j = 1
        For i = 1 To UBound(dataArray, 1)
            If dataArray(i, 6) <> "Completed" And dataArray(i, 6) <> "Cancelled" Then
                filteredData(j, 1) = dataArray(i, 1) ' Customer
                filteredData(j, 2) = dataArray(i, 2) ' Due Date
                filteredData(j, 3) = dataArray(i, 3) ' Priority
                filteredData(j, 4) = dataArray(i, 4) ' Task Type
                filteredData(j, 5) = dataArray(i, 5) ' Description
                filteredData(j, 6) = dataArray(i, 6) ' Status
                j = j + 1
            End If
        Next i
        
        ' Sort array by Priority and Due Date
        ' This is a simple bubble sort - could be optimized for large datasets
        Dim tempRow As Variant
        ReDim tempRow(1 To 6)
        
        For i = 1 To UBound(filteredData, 1) - 1
            For j = i + 1 To UBound(filteredData, 1)
                ' First sort by priority (High > Medium > Low)
                Dim priority1 As Integer, priority2 As Integer
                
                Select Case filteredData(i, 3)
                    Case "High": priority1 = 1
                    Case "Medium": priority1 = 2
                    Case "Low": priority1 = 3
                    Case Else: priority1 = 4
                End Select
                
                Select Case filteredData(j, 3)
                    Case "High": priority2 = 1
                    Case "Medium": priority2 = 2
                    Case "Low": priority2 = 3
                    Case Else: priority2 = 4
                End Select
                
                If priority1 > priority2 Or _
                   (priority1 = priority2 And filteredData(i, 2) > filteredData(j, 2)) Then
                    ' Swap rows
                    For k = 1 To 6
                        tempRow(k) = filteredData(i, k)
                        filteredData(i, k) = filteredData(j, k)
                        filteredData(j, k) = tempRow(k)
                    Next k
                End If
            Next j
        Next i
        
        ' Write sorted data to dashboard
        .Range("A2").Resize(UBound(filteredData, 1), 6).Value = filteredData
        
        ' Format dates
        .Range("B2:B" & (1 + UBound(filteredData, 1))).NumberFormat = "dd-mmm-yyyy"
        
        ' Add action buttons
        For i = 2 To (1 + UBound(filteredData, 1))
            ' Create Complete button
            Dim completeBtn As Button
            Set completeBtn = .Buttons.Add(.Cells(i, 7).left, .Cells(i, 7).top, 70, 20)
            With completeBtn
                .caption = "Complete"
                .Name = "CompleteBtn_" & (i - 1)
                .OnAction = "HandleReminderAction"
            End With
        Next i
        
        ' Format based on due date and priority
        For i = 2 To (1 + UBound(filteredData, 1))
            ' Format priority
            Select Case .Cells(i, 3).Value
                Case "High"
                    .Cells(i, 3).Interior.Color = RGB(255, 200, 200) ' Light red
                Case "Medium"
                    .Cells(i, 3).Interior.Color = RGB(255, 255, 200) ' Light yellow
                Case "Low"
                    .Cells(i, 3).Interior.Color = RGB(200, 255, 200) ' Light green
            End Select
            
            ' Format overdue dates
            If .Cells(i, 2).Value < today Then
                .Cells(i, 2).Interior.Color = RGB(255, 150, 150) ' Overdue red
                .Cells(i, 2).Font.Bold = True
            ElseIf .Cells(i, 2).Value = today Then
                .Cells(i, 2).Interior.Color = RGB(255, 255, 100) ' Today yellow
                .Cells(i, 2).Font.Bold = True
            End If
        Next i

FormatColumns:
        ' Format columns
        .Columns("A:G").AutoFit
        .Columns("E:E").ColumnWidth = 50 ' Description column wider
        
        ' Add refresh button
        Dim refreshBtn As Button
        Set refreshBtn = .Buttons.Add(.Range("A" & (.Cells(.Rows.count, "A").End(xlUp).row + 2)).left, _
                                    .Range("A" & (.Cells(.Rows.count, "A").End(xlUp).row + 2)).top, _
                                    120, 25)
        With refreshBtn
            .caption = "Refresh Dashboard"
            .Name = "RefreshDashboardBtn"
            .OnAction = "ShowReminderDashboard"
        End With
        
        ' Add new reminder button
        Dim newReminderBtn As Button
        Set newReminderBtn = .Buttons.Add(.Range("C" & (.Cells(.Rows.count, "A").End(xlUp).row + 2)).left, _
                                        .Range("C" & (.Cells(.Rows.count, "A").End(xlUp).row + 2)).top, _
                                        120, 25)
        With newReminderBtn
            .caption = "New Reminder"
            .Name = "NewReminderBtn"
            .OnAction = "ShowNewReminderForm"
        End With
        
        ' Add timestamp
        .Range("F" & (.Cells(.Rows.count, "A").End(xlUp).row + 2)).Value = "Last updated: " & Format(Now, "dd-mmm-yyyy hh:mm:ss")
    End With
    
    ' Activate dashboard
    dashboardSheet.Activate
End Sub
```

## 4. Notification System

This function checks for due reminders when Excel opens and on a timer:

```vba
Sub CheckForDueReminders()
    ' Check for reminders due today or overdue
    Dim reminderSheet As Worksheet
    Dim today As Date
    Dim i As Long
    Dim lastRow As Long
    Dim dueCount As Long
    Dim dueList As String
    
    ' Get reminder sheet
    On Error Resume Next
    Set reminderSheet = ThisWorkbook.Sheets("Reminders")
    On Error GoTo 0
    
    If reminderSheet Is Nothing Then Exit Sub
    
    today = Date
    lastRow = reminderSheet.Cells(reminderSheet.Rows.count, "A").End(xlUp).row
    dueCount = 0
    dueList = ""
    
    ' Loop through all reminders
    For i = 2 To lastRow
        ' Check only Pending or In Progress items
        If reminderSheet.Cells(i, 6).Value = "Pending" Or reminderSheet.Cells(i, 6).Value = "In Progress" Then
            ' Check if due today or overdue
            If reminderSheet.Cells(i, 2).Value <= today Then
                dueCount = dueCount + 1
                
                ' Add to notification list
                dueList = dueList & dueCount & ". " & reminderSheet.Cells(i, 1).Value & _
                          " - " & reminderSheet.Cells(i, 4).Value & _
                          " (" & Format(reminderSheet.Cells(i, 2).Value, "dd-mmm-yyyy") & ")" & vbCrLf
            End If
        End If
    Next i
    
    ' Show notification if due items exist
    If dueCount > 0 Then
        MsgBox "You have " & dueCount & " reminder(s) due today or overdue:" & vbCrLf & vbCrLf & _
               dueList & vbCrLf & _
               "Open the Reminder Dashboard to view and manage these items.", _
               vbExclamation, "Due Reminders"
    End If
End Sub

' Call this from Workbook_Open() event
Sub SetupReminderChecks()
    ' Set up automatic reminder checking
    
    ' Check immediately
    CheckForDueReminders
    
    ' Schedule periodic checks (every 30 minutes)
    Application.OnTime Now + TimeValue("00:30:00"), "CheckForDueReminders"
End Sub
```

## 5. User Interface for Creating New Reminders

```vba
Sub ShowNewReminderForm()
    ' Display a form for creating a new reminder
    Dim customerName As String
    Dim dueDate As Date
    Dim priority As String
    Dim taskType As String
    Dim description As String
    
    ' Get customer list
    Dim customerSheet As Worksheet
    Dim customerList As String
    Dim i As Long
    Dim lastRow As Long
    
    On Error Resume Next
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    On Error GoTo 0
    
    If Not customerSheet Is Nothing Then
        lastRow = customerSheet.Cells(customerSheet.Rows.count, "B").End(xlUp).row
        For i = 2 To lastRow
            If customerList <> "" Then customerList = customerList & ","
            customerList = customerList & customerSheet.Cells(i, 2).Value
        Next i
    End If
    
    ' Get customer name (with validation if customer list exists)
    If customerList <> "" Then
        customerName = Application.InputBox("Select or enter customer name:", "New Reminder", Type:=8)
    Else
        customerName = Application.InputBox("Enter customer name:", "New Reminder", Type:=2)
    End If
    If customerName = "False" Then Exit Sub ' User cancelled
    
    ' Get due date
    Dim dueDateStr As String
    dueDateStr = Application.InputBox("Enter due date (MM/DD/YYYY):", "New Reminder", Format(Date + 7, "MM/DD/YYYY"), Type:=2)
    If dueDateStr = "False" Then Exit Sub ' User cancelled
    
    On Error Resume Next
    dueDate = DateValue(dueDateStr)
    If Err.Number <> 0 Then
        MsgBox "Invalid date format. Please try again using MM/DD/YYYY format.", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Get priority
    priority = Application.InputBox("Enter priority (High, Medium, Low):", "New Reminder", "Medium", Type:=2)
    If priority = "False" Then Exit Sub ' User cancelled
    
    ' Validate priority
    Select Case UCase(priority)
        Case "HIGH", "H": priority = "High"
        Case "MEDIUM", "MED", "M": priority = "Medium"
        Case "LOW", "L": priority = "Low"
        Case Else: priority = "Medium"
    End Select
    
    ' Get task type
    taskType = Application.InputBox("Enter task type (e.g., Follow-up Call, Quote Review, etc.):", "New Reminder", "Follow-up Call", Type:=2)
    If taskType = "False" Then Exit Sub ' User cancelled
    
    ' Get description
    description = Application.InputBox("Enter description or notes:", "New Reminder", "", Type:=2)
    If description = "False" Then Exit Sub ' User cancelled
    
    ' Create the reminder
    CreateReminder customerName, dueDate, taskType, description, priority
    
    ' Refresh dashboard if open
    On Error Resume Next
    If ActiveSheet.Name = "ReminderDashboard" Then
        ShowReminderDashboard
    End If
End Sub

' Handler for reminder action buttons
Sub HandleReminderAction()
    ' Process button clicks from reminder dashboard
    Dim buttonName As String
    Dim buttonType As String
    Dim reminderIndex As String
    Dim reminderRow As Long
    
    ' Get button name
    buttonName = Application.Caller
    
    ' Parse button name to get type and index
    If InStr(buttonName, "CompleteBtn_") > 0 Then
        buttonType = "Complete"
        reminderIndex = Mid(buttonName, Len("CompleteBtn_") + 1)
    Else
        Exit Sub ' Unknown button
    End If
    
    ' Get corresponding row in reminders sheet
    Dim dashboardSheet As Worksheet
    Dim reminderSheet As Worksheet
    
    Set dashboardSheet = ThisWorkbook.Sheets("ReminderDashboard")
    Set reminderSheet = ThisWorkbook.Sheets("Reminders")
    
    ' Get customer and task from dashboard
    Dim customerName As String
    Dim taskType As String
    
    ' Convert dashboard row index to integer
    On Error Resume Next
    reminderIndex = Val(reminderIndex)
    If Err.Number <> 0 Or reminderIndex = 0 Then Exit Sub
    On Error GoTo 0
    
    customerName = dashboardSheet.Cells(reminderIndex + 1, 1).Value
    taskType = dashboardSheet.Cells(reminderIndex + 1, 4).Value
    
    ' Find matching reminder in Reminders sheet
    Dim lastRow As Long
    lastRow = reminderSheet.Cells(reminderSheet.Rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        If reminderSheet.Cells(i, 1).Value = customerName And _
           reminderSheet.Cells(i, 4).Value = taskType And _
           (reminderSheet.Cells(i, 6).Value = "Pending" Or reminderSheet.Cells(i, 6).Value = "In Progress") Then
            reminderRow = i
            Exit For
        End If
    Next i
    
    If reminderRow = 0 Then
        MsgBox "Could not find the corresponding reminder in the database.", vbExclamation
        Exit Sub
    End If
    
    ' Handle the action based on button type
    If buttonType = "Complete" Then
        ' Ask for completion notes
        Dim completionNotes As String
        completionNotes = Application.InputBox("Enter any notes about this completion:", "Complete Reminder", Type:=2)
        If completionNotes = "False" Then Exit Sub ' User cancelled
        
        ' Complete the reminder
        CompleteReminder reminderRow, completionNotes
        
        ' Refresh dashboard
        ShowReminderDashboard
    End If
End Sub
```

## 6. Integration with Your Existing System

You can integrate this with your EnhancedDynamicScript module by adding a call to create reminders:

```vba
' Add this to your ScheduleScriptFollowUp function
Public Sub ScheduleScriptFollowUp()
    ' Existing code...
    
    ' After scheduling the follow-up in your current system
    ' Also create a reminder
    CreateReminder activeCustomerName, followupDate, followupType, _
                  "Follow-up scheduled during call. " & notes, "Medium"
    
    ' Existing code...
End Sub
```

## Benefits of This Approach

1. **Self-contained within Excel** - No external dependencies on Outlook or other systems
2. **Prioritization built-in** - Tasks are automatically ordered by priority and due date
3. **Visual dashboard** - Clear visualization of upcoming tasks
4. **Permission control** - No automatic actions without your approval
5. **Customizable** - Easy to modify for your specific workflow needs
6. **Integrated notifications** - Proactive alerts for due tasks

This system provides a comprehensive replacement for Outlook reminders while maintaining complete control over notifications and actions. It also creates a centralized view of upcoming tasks prioritized by importance and due date, making it easier to manage your customer follow-ups effectively.

Would you like me to provide more detail on any particular aspect of this reminder system?
