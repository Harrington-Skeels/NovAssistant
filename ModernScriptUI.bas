Attribute VB_Name = "ModernScriptUI"
' ====================================================================
' ModernScriptUI Module
' ====================================================================
' This module creates an enhanced user interface for the dynamic script system
' with a modern look and feel, improved navigation, and better visualization

Option Explicit

' UI colors for consistent styling
Private Const COLOR_PRIMARY = 4227072       ' Dark green (RGB 0, 66, 37)
Private Const COLOR_PRIMARY_DARK = 3368601  ' Darker green (RGB 1, 51, 32)
Private Const COLOR_PRIMARY_LIGHT = 5287936 ' Light green (RGB 80, 160, 80)
Private Const COLOR_SECONDARY = 39423       ' Orange (RGB 255, 153, 0)
Private Const COLOR_TEXT_LIGHT = 16777215   ' White
Private Const COLOR_BACKGROUND = 15921906   ' Light gray (RGB 242, 242, 242)
Private Const COLOR_BACKGROUND_DARK = 15132390 ' Dark gray (RGB 230, 230, 230)
Private Const COLOR_TEXT_DARK = 3158064     ' Dark gray (RGB 48, 48, 48)

' UI state variables
Private scriptSheetName As String
Private currentScriptSection As String
Private isScriptInitialized As Boolean
Private useCompactMode As Boolean

' Initialize the modern script UI
Public Sub CreateModernScriptUI()
On Error GoTo ErrorHandler
    ' Define the script sheet name
    scriptSheetName = "ScriptUI"
    
    ' Check if ScriptUI sheet already exists
    Dim scriptSheet As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = scriptSheetName Then
            sheetExists = True
            Set scriptSheet = ws
            Exit For
        End If
    Next ws
    
    ' Create the sheet if it doesn't exist
    If Not sheetExists Then
        Set scriptSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        scriptSheet.Name = scriptSheetName
    Else
        ' Clear existing content
        scriptSheet.Cells.Clear
        
        ' Remove existing shapes and form controls
        ClearExistingControls scriptSheet
    End If
    
    ' Set up the modern UI layout
    SetupModernUILayout scriptSheet
    
    ' Initialize state variables
    currentScriptSection = "welcome"
    isScriptInitialized = True
    useCompactMode = False
    
    ' Activate the script sheet
    scriptSheet.Activate
    scriptSheet.Range("A1").Select
    
    ' Update the UI with welcome content
    UpdateScriptContent "welcome"
    
    ' Setup right-click menu for quick navigation
    SetupRightClickMenu scriptSheet
    
    ' Success message
    MsgBox "Modern script UI created successfully! Use the navigation buttons to move through the script.", vbInformation, "Script UI"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating the modern script UI: " & Err.description, vbExclamation, "Error"
End Sub

' Clear existing controls from the worksheet
Private Sub ClearExistingControls(ws As Worksheet)
On Error Resume Next
    ' Remove buttons
    Dim btn As Button
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    
    ' Remove shapes
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' Remove other form controls
    Dim ctl As OLEObject
    For Each ctl In ws.OLEObjects
        ctl.Delete
    Next ctl
End Sub

' Set up the modern UI layout
Private Sub SetupModernUILayout(ws As Worksheet)
On Error Resume Next
    ' Set column widths
    ws.Columns("A").ColumnWidth = 2       ' Left margin
    ws.Columns("B:N").ColumnWidth = 10    ' Content columns
    ws.Columns("O").ColumnWidth = 2       ' Right margin
    
    ' Freeze panes to keep header visible when scrolling
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    
    ' Create header
    With ws.Range("B1:N2")
        .Merge
        .Value = "NOVATED LEASE CONVERSATION GUIDE"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
        .RowHeight = 40
    End With
    
    ' Create navigation bar
    CreateNavigationBar ws
    
    ' Create customer info section
    CreateCustomerInfoSection ws
    
    ' Create script content area
    CreateScriptContentArea ws
    
    ' Create customer response area
    CreateCustomerResponseArea ws
    
    ' Create notes area
    CreateNotesArea ws
    
    ' Create action buttons
    CreateActionButtons ws
    
    ' Add navigation buttons
    CreateNavigationButtons ws
    
    ' Add version info at bottom
    ws.Range("B52:N52").Merge
    ws.Range("B52").Value = "NovAssistant v2.0 - Modern Script UI - " & Format(Now, "yyyy-mm-dd")
    ws.Range("B52").Font.Italic = True
    ws.Range("B52").Font.Size = 8
    ws.Range("B52").HorizontalAlignment = xlCenter
End Sub

' Create the navigation bar
Private Sub CreateNavigationBar(ws As Worksheet)
On Error Resume Next
    ' Create the navigation bar background
    With ws.Range("B3:N3")
        .Merge
        .Interior.Color = COLOR_PRIMARY_DARK
        .RowHeight = 24
    End With
    
    ' Create navigation sections
    CreateNavButton ws, "navIntro", "Introduction", "B3", 1
    CreateNavButton ws, "navQualify", "Qualifying", "C3", 2
    CreateNavButton ws, "navEducate", "Education", "D3", 3
    CreateNavButton ws, "navBenefits", "Benefits", "E3", 4
    CreateNavButton ws, "navEndOptions", "End Options", "F3", 5
    CreateNavButton ws, "navTradeIn", "Trade-In", "G3", 6
    CreateNavButton ws, "navDetails", "Details", "H3", 7
    CreateNavButton ws, "navObject", "Objections", "I3", 8
    CreateNavButton ws, "navClosing", "Closing", "J3", 9
    CreateNavButton ws, "navFollow", "Follow-Up", "K3", 10
    CreateNavButton ws, "navAppl", "Application", "L3", 11
    CreateNavButton ws, "navSettl", "Settlement", "M3", 12
    CreateNavButton ws, "navSettings", "?", "N3", 13
End Sub

' Create a navigation button on the nav bar
Private Sub CreateNavButton(ws As Worksheet, btnName As String, caption As String, position As String, index As Integer)
On Error Resume Next
    Dim btn As Button
    
    ' Calculate position
    Dim left As Double, top As Double, width As Double, height As Double
    left = ws.Range(position).left
    top = ws.Range(position).top
    width = ws.Range(position).width - 1
    height = ws.Range(position).height - 1
    
    ' Create button
    Set btn = ws.Buttons.Add(left, top, width, height)
    With btn
        .Name = btnName
        .caption = caption
        .Font.Size = 9
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Color = COLOR_TEXT_LIGHT
        .OnAction = "ModernScriptUI.NavigateToSection"
    End With
End Sub

' Create the customer info section
Private Sub CreateCustomerInfoSection(ws As Worksheet)
On Error Resume Next
    ' Create customer info header
    With ws.Range("B4:N4")
        .Merge
        .Value = "CUSTOMER INFORMATION"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_BACKGROUND_DARK
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Create labels
    ws.Range("B5").Value = "Name:"
    ws.Range("B5").Font.Bold = True
    ws.Range("B5").HorizontalAlignment = xlRight
    
    ws.Range("F5").Value = "Phone:"
    ws.Range("F5").Font.Bold = True
    ws.Range("F5").HorizontalAlignment = xlRight
    
    ws.Range("J5").Value = "Stage:"
    ws.Range("J5").Font.Bold = True
    ws.Range("J5").HorizontalAlignment = xlRight
    
    ' Create value fields
    With ws.Range("C5:E5")
        .Merge
        .Value = ""
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Name = "CustomerName"
    End With
    
    With ws.Range("G5:I5")
        .Merge
        .Value = ""
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Name = "CustomerPhone"
    End With
    
    With ws.Range("K5:N5")
        .Merge
        .Value = ""
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Name = "CustomerStage"
    End With
    
    ' Add call timer
    ws.Range("B6").Value = "Duration:"
    ws.Range("B6").Font.Bold = True
    ws.Range("B6").HorizontalAlignment = xlRight
    
    With ws.Range("C6:E6")
        .Merge
        .Value = "00:00:00"
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Name = "CallDuration"
    End With
    
    ' Add customer summary area
    ws.Range("F6").Value = "Summary:"
    ws.Range("F6").Font.Bold = True
    ws.Range("F6").HorizontalAlignment = xlRight
    
    With ws.Range("G6:N6")
        .Merge
        .Value = ""
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Name = "CustomerSummary"
    End With
End Sub

' Create the script content area
Private Sub CreateScriptContentArea(ws As Worksheet)
On Error Resume Next
    ' Create script section header
    With ws.Range("B7:N7")
        .Merge
        .Value = "SCRIPT CONTENT"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Create script content area
    With ws.Range("B8:N25")
        .Merge
        .Value = "Select a section from the navigation bar above to view script content."
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Name = "ScriptContent"
    End With
    
    ' Add section title
    With ws.Range("B26:N26")
        .Merge
        .Value = "SECTION: Welcome"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = COLOR_SECONDARY
        .Font.Color = COLOR_TEXT_LIGHT
        .Font.Bold = True
        .Name = "SectionTitle"
    End With
End Sub

' Create the customer response area
Private Sub CreateCustomerResponseArea(ws As Worksheet)
On Error Resume Next
    ' Create response header
    With ws.Range("B27:N27")
        .Merge
        .Value = "CUSTOMER RESPONSE OPTIONS"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Create response option area
    With ws.Range("B28:N35")
        .Interior.Color = COLOR_BACKGROUND
    End With
    
    ' Add response options
    CreateResponseOption ws, 1, "First Timer", "HandleFirstTimer"
    CreateResponseOption ws, 2, "Knows a Little", "HandleSomeKnowledge"
    CreateResponseOption ws, 3, "Well Educated", "HandleWellEducated"
    CreateResponseOption ws, 4, "Skip to Qualifying", "JumpToQualifying"
End Sub

' Create a response option button
Private Sub CreateResponseOption(ws As Worksheet, index As Integer, caption As String, handlerName As String)
On Error Resume Next
    Dim btn As Button
    Dim row As Integer, col As Integer
    
    ' Calculate position (2 buttons per row, starting at row 28)
    row = 28 + ((index - 1) \ 2) * 3
    col = IIf((index Mod 2) = 1, "B", "H")
    
    ' Create button
    Set btn = ws.Buttons.Add(ws.Range(col & row).left, ws.Range(col & row).top, _
                            ws.Range(col & row & ":E" & row).width, _
                            ws.Range(col & row & ":" & col & (row + 2)).height)
    
    With btn
        .Name = "resp" & index
        .caption = caption
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI." & handlerName
    End With
End Sub

' Create the notes area
Private Sub CreateNotesArea(ws As Worksheet)
On Error Resume Next
    ' Create notes header
    With ws.Range("B36:N36")
        .Merge
        .Value = "CALL NOTES"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Create notes content area
    With ws.Range("B37:N44")
        .Merge
        .Value = ""
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Name = "CallNotes"
    End With
End Sub

' Create the action buttons
Private Sub CreateActionButtons(ws As Worksheet)
On Error Resume Next
    ' Create action button container
    With ws.Range("B45:N45")
        .Merge
        .Value = "ACTIONS"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .Interior.Color = COLOR_PRIMARY
        .Font.Color = COLOR_TEXT_LIGHT
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' Create action buttons
    With ws.Range("B46:N50")
        .Interior.Color = COLOR_BACKGROUND
    End With
    
    ' Add Start Call button
    Dim btnStart As Button
    Set btnStart = ws.Buttons.Add(ws.Range("B46").left, ws.Range("B46").top, _
                              ws.Range("B46:C48").width, ws.Range("B46:B48").height)
    With btnStart
        .Name = "btnStartCall"
        .caption = "Start New Call"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI.StartNewCall"
    End With
    
    ' Add End Call button
    Dim btnEnd As Button
    Set btnEnd = ws.Buttons.Add(ws.Range("D46").left, ws.Range("D46").top, _
                              ws.Range("D46:E48").width, ws.Range("D46:D48").height)
    With btnEnd
        .Name = "btnEndCall"
        .caption = "End Call"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI.EndCurrentCall"
    End With
    
    ' Add Save Notes button
    Dim btnSave As Button
    Set btnSave = ws.Buttons.Add(ws.Range("F46").left, ws.Range("F46").top, _
                              ws.Range("F46:G48").width, ws.Range("F46:F48").height)
    With btnSave
        .Name = "btnSaveNotes"
        .caption = "Save Notes"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI.SaveNotes"
    End With
    
    ' Add Schedule Follow-up button
    Dim btnFollowUp As Button
    Set btnFollowUp = ws.Buttons.Add(ws.Range("H46").left, ws.Range("H46").top, _
                              ws.Range("H46:I48").width, ws.Range("H46:H48").height)
    With btnFollowUp
        .Name = "btnFollowUp"
        .caption = "Schedule Follow-up"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI.ScheduleFollowUp"
    End With
    
    ' Add Update Stage button
    Dim btnStage As Button
    Set btnStage = ws.Buttons.Add(ws.Range("J46").left, ws.Range("J46").top, _
                              ws.Range("J46:K48").width, ws.Range("J46:J48").height)
    With btnStage
        .Name = "btnUpdateStage"
        .caption = "Update Stage"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI.UpdateStage"
    End With
    
    ' Add Send Email button
    Dim btnEmail As Button
    Set btnEmail = ws.Buttons.Add(ws.Range("L46").left, ws.Range("L46").top, _
                              ws.Range("L46:N48").width, ws.Range("L46:L48").height)
    With btnEmail
        .Name = "btnSendEmail"
        .caption = "Send Email"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI.SendEmail"
    End With
End Sub

' Create navigation buttons
Private Sub CreateNavigationButtons(ws As Worksheet)
On Error Resume Next
    ' Add Previous button
    Dim btnPrev As Button
    Set btnPrev = ws.Buttons.Add(ws.Range("B49").left, ws.Range("B49").top, _
                              ws.Range("B49:E50").width, ws.Range("B49:B50").height)
    With btnPrev
        .Name = "btnPrevious"
        .caption = "? Previous"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI.NavigateToPrevious"
    End With
    
    ' Add Toggle View button
    Dim btnToggle As Button
    Set btnToggle = ws.Buttons.Add(ws.Range("F49").left, ws.Range("F49").top, _
                              ws.Range("F49:J50").width, ws.Range("F49:F50").height)
    With btnToggle
        .Name = "btnToggleView"
        .caption = "Toggle Compact Mode"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI.ToggleCompactMode"
    End With
    
    ' Add Next button
    Dim btnNext As Button
    Set btnNext = ws.Buttons.Add(ws.Range("K49").left, ws.Range("K49").top, _
                              ws.Range("K49:N50").width, ws.Range("K49:K50").height)
    With btnNext
        .Name = "btnNext"
        .caption = "Next ?"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Name = "Calibri"
        .OnAction = "ModernScriptUI.NavigateToNext"
    End With
End Sub

' Setup right-click menu for quick navigation
Private Sub SetupRightClickMenu(ws As Worksheet)
On Error Resume Next
    ' This requires Application.OnKey which we can't fully implement in this code
    ' But we can prepare the CommandBar code here for future use
    
    ' This would create a custom right-click menu
    ' We'll leave this as a placeholder for now
End Sub

' Update the script content based on the selected section
Private Sub UpdateScriptContent(sectionName As String)
On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(scriptSheetName)
    
    ' Update the current section
    currentScriptSection = sectionName
    
    ' Get section title and content
    Dim sectionTitle As String
    Dim sectionContent As String
    Dim responseOptions As Variant
    
    ' Set content based on section
    Select Case sectionName
        Case "welcome"
            sectionTitle = "Welcome & Introduction"
            sectionContent = "Hi [Customer Name], my name is [Your Name]. Before I start, just a reminder that our calls are recorded for training purposes." & vbCrLf & vbCrLf & _
                            "I understand you are calling about a novated lease?" & vbCrLf & vbCrLf & _
                            "How familiar are you with novated leasing?"
            responseOptions = Array("First Timer", "Knows a Little", "Well Educated", "Skip to Qualifying")
            
        Case "first-timer"
            sectionTitle = "Education - New to Novated Leasing"
            sectionContent = "That's perfectly fine - many people are new to novated leasing. Let me explain the basics:" & vbCrLf & vbCrLf & _
                            "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf & _
                            "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars."
            responseOptions = Array("Continue to Qualifying", "Has Questions", "", "")
            
        Case "some-knowledge"
            sectionTitle = "Education - Some Knowledge"
            sectionContent = "Great, so you have some familiarity with novated leasing. Let me confirm and expand on what you might already know:" & vbCrLf & vbCrLf & _
                            "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf & _
                            "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars."
            responseOptions = Array("Continue to Qualifying", "Has Questions", "", "")
            
        Case "well-educated"
            sectionTitle = "Education - Well Educated"
            sectionContent = "Excellent! Since you're already familiar with novated leasing, let's focus on the specific aspects that would be most beneficial for your situation." & vbCrLf & vbCrLf & _
                            "I'd like to ask you some specific questions about your needs to tailor our discussion to what would work best for you."
            responseOptions = Array("Continue to Qualifying", "", "", "")
            
        Case "qualifying"
            sectionTitle = "Qualifying Questions"
            sectionContent = "So, I can make sure the lease reflects what you need, can I ask you a few questions?" & vbCrLf & vbCrLf & _
                            "1. Have you ever had a novated lease before?" & vbCrLf & _
                            "2. Do you have a car in mind? (new/used electric or petrol)" & vbCrLf & _
                            "3. Have you test driven the car?" & vbCrLf & _
                            "4. Do you have a budget in mind?" & vbCrLf & _
                            "5. When would you like to be in the new vehicle?" & vbCrLf & _
                            "6. How are you paying for your current car? Did you pay cash, car finance, home loan?" & vbCrLf & vbCrLf & _
                            "I will go through how a novated lease works and the benefits you get and then I will get some details from you so I can write up an indicative quote for you."
            responseOptions = Array("Continue to Education", "Customer Has Objections", "", "")
            
        Case "educating"
            sectionTitle = "Education - Lease Basics"
            sectionContent = "A novated lease is a combination of a pre and post-tax deduction that is tied to your payroll that wraps up all the vehicle running costs that you typically pay for." & vbCrLf & vbCrLf & _
                            "The benefit of a novated lease is that the vehicle is financed less the GST, your running costs are GST free and you get to pay for a portion of the transaction using your pre-tax dollars." & vbCrLf & vbCrLf & _
                            "The novated lease will include; the car, services, maintenance, tyres, rego, insurance and fuel. We will set a budget that reflects how many kilometres you drive each year." & vbCrLf & vbCrLf & _
                            "Getting your budget correct is important with a Novated Lease, if we under estimate your running costs the lease will look cheap and attractive, however you won't have money to cover all the running costs. If we over-budget, it will look too expensive and you won't want to take the lease."
            responseOptions = Array("Continue to Benefits", "Customer Has Questions", "", "")
            
        Case "benefits"
            sectionTitle = "Benefits of Novated Leasing"
            sectionContent = "The vehicle is also fully maintained, meaning you'll get fuel cards for your fuel, the dealer will invoice us for your servicing, same with the tyre shop. Rego renewals can be uploaded, and we'll pay them directly." & vbCrLf & vbCrLf & _
                            "So other than tolls and fines you should not need to outlay any further money on your vehicle outside of those deductions." & vbCrLf & vbCrLf & _
                            "The easy way to think of this, is that the deductions will go to an account held at sgfleet. As you spend money on the vehicle, it will deduct from that account and you can manage and monitor this through our online app." & vbCrLf & vbCrLf & _
                            "Throughout the lease you're paying the vehicle down to a residual or balloon amount set by the ATO."
            responseOptions = Array("Continue to Lease End Options", "Customer Has Questions", "", "")
            
        Case "lease-end"
            sectionTitle = "Lease End Options"
            sectionContent = "At the end of the lease, you will have a few options." & vbCrLf & vbCrLf & _
                            "1. If you love the car, you can pay it out and the car is yours;" & vbCrLf & _
                            "2. If you love the car and the lease, you can refinance and extend" & vbCrLf & _
                            "3. Another option is to sell the car -- any money you make above the residual is yours tax free" & vbCrLf & _
                            "4. Or the most popular option, is to trade the car in and upgrade." & vbCrLf & vbCrLf & _
                            "We can also source the vehicle for you, biggest benefit is that we are entitled to fleet discounts which we pass on to you." & vbCrLf & vbCrLf & _
                            "All our new quotes will include an Eco Protection Pack and Minor Damage Repair membership."
            responseOptions = Array("Continue to Trade-In", "Customer Has Questions", "", "")
            
        Case "trade-in"
            sectionTitle = "Trade-In Information"
            sectionContent = "Do you have a car you are looking to trade in or sell?" & vbCrLf & vbCrLf & _
                            "We have a trade advantage program where we will have several wholesalers value your existing car. If you're happy with the price, you can trade it in and we will even come and collect the car for you." & vbCrLf & vbCrLf & _
                            "Do you have any questions at this point?"
            responseOptions = Array("Continue to Gather Details", "Customer Has Questions", "", "")
            
        Case "objections"
            sectionTitle = "Handling Objections"
            section

