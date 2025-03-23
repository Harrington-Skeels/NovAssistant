Attribute VB_Name = "ScriptEngine"

' ====================================================================
' ScriptEngine Module - Version #1
' ====================================================================
' This module powers the adaptive conversation flow with
' improved branching logic and context-awareness

Option Explicit

' Script navigation states
Private currentScriptState As String
Private scriptHistory As Collection
Private customerResponses As Object ' Dictionary
Private customerAttributes As Object ' Dictionary

' State constants - expanded for more flexibility
Private Const STATE_INITIAL = "Initial"
Private Const STATE_KNOWLEDGE_ASSESS = "KnowledgeAssessment"
Private Const STATE_BEGINNER = "BeginnerEducation"
Private Const STATE_INTERMEDIATE = "IntermediateEducation"
Private Const STATE_ADVANCED = "AdvancedEducation"
Private Const STATE_QUALIFYING = "Qualifying"
Private Const STATE_VEHICLE_NEEDS = "VehicleNeeds"
Private Const STATE_FINANCIAL = "FinancialDiscussion"
Private Const STATE_BENEFITS = "BenefitsExplanation"
Private Const STATE_LEASE_OPTIONS = "LeaseEndOptions"
Private Const STATE_OBJECTIONS = "Objections"
Private Const STATE_QUOTE_INFO = "QuoteInformation"
Private Const STATE_CLOSING = "Closing"
Private Const STATE_FOLLOW_UP = "FollowUp"

' Initialize the script engine
Public Sub InitializeScriptEngine()
    ' Initialize collections
    Set scriptHistory = New Collection
    Set customerResponses = CreateObject("Scripting.Dictionary")
    Set customerAttributes = CreateObject("Scripting.Dictionary")
    
    ' Set initial state
    currentScriptState = STATE_INITIAL
    
    ' Initialize data capture
    InitializeDataCapture GetActiveCustomerName
    
    ' Set initial attributes
    customerAttributes("KnowledgeLevel") = "Unknown"
    customerAttributes("InterestLevel") = "Unknown"
    customerAttributes("Budget") = "Unknown"
    customerAttributes("Timeline") = "Unknown"
    
    ' Load script content
    UpdateScriptContent
End Sub

' Get active customer name
Private Function GetActiveCustomerName() As String
    Dim scriptSheet As Worksheet
    
    ' Get modern script sheet
    On Error Resume Next
    Set scriptSheet = ThisWorkbook.Sheets("ModernScript")
    On Error GoTo 0
    
    If scriptSheet Is Nothing Then
        GetActiveCustomerName = ""
        Exit Function
    End If
    
    ' Get customer name from UI
    GetActiveCustomerName = scriptSheet.Range("CustomerName").Value
End Function

' Update script content based on current state
Public Sub UpdateScriptContent()
    Dim scriptSheet As Worksheet
    Dim scriptContent As String
    Dim scriptPath As String
    
    ' Get modern script sheet
    On Error Resume Next
    Set scriptSheet = ThisWorkbook.Sheets("ModernScript")
    On Error GoTo 0
    
    If scriptSheet Is Nothing Then Exit Sub
    
    ' Generate script content and path based on state
    Select Case currentScriptState
        Case STATE_INITIAL
            scriptContent = GetInitialGreeting()
            scriptPath = "Initial Greeting"
            
        Case STATE_KNOWLEDGE_ASSESS
            scriptContent = GetKnowledgeAssessment()
            scriptPath = "Knowledge Assessment"
            
        Case STATE_BEGINNER
            scriptContent = GetBeginnerEducation()
            scriptPath = "Basic Education"
            
        Case STATE_INTERMEDIATE
            scriptContent = GetIntermediateEducation()
            scriptPath = "Education"
            
        Case STATE_ADVANCED
            scriptContent = GetAdvancedEducation()
            scriptPath = "Overview"
            
        Case STATE_QUALIFYING
            scriptContent = GetQualifyingQuestions()
            scriptPath = "Qualifying Questions"
            
        Case STATE_VEHICLE_NEEDS
            scriptContent = GetVehicleNeedsQuestions()
            scriptPath = "Vehicle Requirements"
            
        Case STATE_FINANCIAL
            scriptContent = GetFinancialQuestions()
            scriptPath = "Financial Information"
            
        Case STATE_BENEFITS
            scriptContent = GetBenefitsExplanation()
            scriptPath = "Benefits Explanation"
            
        Case STATE_LEASE_OPTIONS
            scriptContent = GetLeaseEndOptions()
            scriptPath = "Lease End Options"
            
        Case STATE_OBJECTIONS
            scriptContent = GetObjectionsResponses()
            scriptPath = "Handling Objections"
            
        Case STATE_QUOTE_INFO
            scriptContent = GetQuoteInfoCollection()
            scriptPath = "Quote Information"
            
        Case STATE_CLOSING
            scriptContent = GetClosingScript()
            scriptPath = "Closing"
            
        Case STATE_FOLLOW_UP
            scriptContent = GetFollowUpScript()
            scriptPath = "Follow-Up"
            
        Case Else
            scriptContent = "Unknown script state. Please restart the call."
            scriptPath = "Error"
    End Select
    
    ' Update UI with content
    scriptSheet.Range("ScriptContent").Value = scriptContent
    scriptSheet.Range("ScriptPath").Value = "Current Path: " & scriptPath
    
    ' Create response buttons
    CreateResponseButtons scriptSheet
    
    ' Update customer data display
    UpdateModernUI
End Sub

' Handle response selection
Public Sub HandleResponse(response As String)
    ' Add current state to history
    scriptHistory.Add currentScriptState
    
    ' Store response
    customerResponses(currentScriptState) = response
    
    ' Store in data capture
    StoreCustomerResponse currentScriptState, response, currentScriptState
    
    ' Update attributes based on response
    UpdateCustomerAttributes response
    
    ' Determine next state based on current state and response
    DetermineNextState response
    
    ' Update script content
    UpdateScriptContent
End Sub

' Update customer attributes based on response
Private Sub UpdateCustomerAttributes(response As String)
    Select Case currentScriptState
        Case STATE_KNOWLEDGE_ASSESS
            If response = "First Time" Then
                customerAttributes("KnowledgeLevel") = "Beginner"
            ElseIf response = "Some Knowledge" Then
                customerAttributes("KnowledgeLevel") = "Intermediate"
            ElseIf response = "Very Familiar" Then
                customerAttributes("KnowledgeLevel") = "Advanced"
            End If
            
        Case STATE_QUALIFYING
            If InStr(1, response, "looking to purchase", vbTextCompare) > 0 Then
                customerAttributes("Timeline") = "Near Term"
            ElseIf InStr(1, response, "just researching", vbTextCompare) > 0 Then
                customerAttributes("Timeline") = "Research Phase"
            End If
            
        Case STATE_VEHICLE_NEEDS
            If InStr(1, response, "specific model", vbTextCompare) > 0 Then
                customerAttributes("DecisionReadiness") = "High"
            End If
            
        Case STATE_FINANCIAL
            If InStr(1, response, "budget", vbTextCompare) > 0 Then
                customerAttributes("Budget") = ExtractBudget(response)
            End If
    End Select
End Sub

' Extract budget information from response
Private Function ExtractBudget(response As String) As String
    ' Simple extraction - look for dollar amounts
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "(\$[\d,]+)|(\d+k)"
    regex.Global = True
    regex.IgnoreCase = True
    
    Dim matches As Object
    Set matches = regex.Execute(response)
    
    If matches.count > 0 Then
        ExtractBudget = matches(0).Value
    Else
        ExtractBudget = "Unknown"
    End If
End Function

' Determine next state based on current state and response
Private Sub DetermineNextState(response As String)
    Select Case currentScriptState
        Case STATE_INITIAL
            currentScriptState = STATE_KNOWLEDGE_ASSESS
            
        Case STATE_KNOWLEDGE_ASSESS
            If customerAttributes("KnowledgeLevel") = "Beginner" Then
                currentScriptState = STATE_BEGINNER
            ElseIf customerAttributes("KnowledgeLevel") = "Intermediate" Then
                currentScriptState = STATE_INTERMEDIATE
            Else
                currentScriptState = STATE_ADVANCED
            End If
            
        Case STATE_BEGINNER, STATE_INTERMEDIATE, STATE_ADVANCED
            currentScriptState = STATE_QUALIFYING
            
        Case STATE_QUALIFYING
            If InStr(1, response, "vehicle questions", vbTextCompare) > 0 Then
                currentScriptState = STATE_VEHICLE_NEEDS
            Else
                currentScriptState = STATE_FINANCIAL
            End If
            
        Case STATE_VEHICLE_NEEDS
            currentScriptState = STATE_FINANCIAL
            
        Case STATE_FINANCIAL
            currentScriptState = STATE_BENEFITS
            
        Case STATE_BENEFITS
            currentScriptState = STATE_LEASE_OPTIONS
            
        Case STATE_LEASE_OPTIONS
            If InStr(1, response, "objection", vbTextCompare) > 0 Then
                currentScriptState = STATE_OBJECTIONS
            Else
                currentScriptState = STATE_QUOTE_INFO
            End If
            
        Case STATE_OBJECTIONS
            If InStr(1, response, "resolved", vbTextCompare) > 0 Then
                currentScriptState = STATE_QUOTE_INFO
            Else
                currentScriptState = STATE_CLOSING
            End If
            
        Case STATE_QUOTE_INFO
            currentScriptState = STATE_CLOSING
            
        Case STATE_CLOSING
            If InStr(1, response, "follow-up", vbTextCompare) > 0 Then
                currentScriptState = STATE_FOLLOW_UP
            End If
            
        Case Else
            ' Default to qualifying if uncertain
            currentScriptState = STATE_QUALIFYING
    End Select
End Sub

' Create response buttons
Private Sub CreateResponseButtons(ws As Worksheet)
    ' Clear existing buttons
    ClearResponseButtons ws
    
    ' Get response options based on current state
    Dim options As Variant
    options = GetResponseOptions()
    
    ' Create buttons for each option
    Dim i As Integer
    Dim btnTop As Long, btnLeft As Long
    Dim btnHeight As Long, btnWidth As Long
    Dim btnRow As Integer, btnCol As Integer
    Dim maxButtonsPerRow As Integer
    
    ' Set button dimensions
    btnHeight = 22
    btnWidth = 100
    maxButtonsPerRow = 3
    
    ' Create buttons
    For i = LBound(options) To UBound(options)
        ' Calculate position
        btnRow = i \ maxButtonsPerRow
        btnCol = i Mod maxButtonsPerRow
        
        btnTop = ws.Range("B22").top + (btnRow * (btnHeight + 5))
        btnLeft = ws.Range("B22").left + (btnCol * (btnWidth + 5))
        
        ' Create button
        Dim btn As Button
        Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
        With btn
            .caption = options(i)
            .Name = "ResponseBtn" & i
            .OnAction = "HandleResponseClick"
            .Font.Size = 9
            
            ' Assign Alt shortcut key if within first 6 options
            If i < 6 Then
                .Accelerator = CStr(i + 1)
            End If
        End With
    Next i
End Sub

' Clear existing response buttons
Private Sub ClearResponseButtons(ws As Worksheet)
    Dim btn As Button
    
    ' Remove all buttons in the response area
    For Each btn In ws.Buttons
        If left(btn.Name, 11) = "ResponseBtn" Then
            btn.Delete
        End If
    Next btn
End Sub

' Get response options based on current state
Private Function GetResponseOptions() As Variant
    Select Case currentScriptState
        Case STATE_INITIAL
            GetResponseOptions = Array("Continue")
            
        Case STATE_KNOWLEDGE_ASSESS
            GetResponseOptions = Array("First Time", "Some Knowledge", "Very Familiar")
            
        Case STATE_BEGINNER, STATE_INTERMEDIATE, STATE_ADVANCED
            GetResponseOptions = Array("Continue", "Have Questions")
            
        Case STATE_QUALIFYING
            GetResponseOptions = Array("Continue to Vehicle", "Continue to Financial", "Have Objection")
            
        Case STATE_VEHICLE_NEEDS
            GetResponseOptions = Array("Continue", "Need More Information")
            
        Case STATE_FINANCIAL
            GetResponseOptions = Array("Continue", "Have Budget Questions")
            
        Case STATE_BENEFITS
            GetResponseOptions = Array("Continue", "Questions About Benefits")
            
        Case STATE_LEASE_OPTIONS
            GetResponseOptions = Array("Continue", "Questions About End Options", "Have Objection")
            
        Case STATE_OBJECTIONS
            GetResponseOptions = Array("Objection Resolved", "Need More Information", "Not Interested")
            
        Case STATE_QUOTE_INFO
            GetResponseOptions = Array("Continue", "Need to Check Details")
            
        Case STATE_CLOSING
            GetResponseOptions = Array("Complete Call", "Schedule Follow-up", "Send Quote")
            
        Case STATE_FOLLOW_UP
            GetResponseOptions = Array("Confirm Follow-up", "Change Details")
            
        Case Else
            GetResponseOptions = Array("Continue")
    End Select
End Function

' Handle response button click
Public Sub HandleResponseClick()
    Dim btnName As String
    Dim response As String
    Dim ws As Worksheet
    
    ' Get button name
    btnName = Application.Caller
    
    ' Get modern script sheet
    Set ws = ThisWorkbook.Sheets("ModernScript")
    
    ' Get button caption (response)
    response = ws.Buttons(btnName).caption
    
    ' Handle the response
    HandleResponse response
End Sub

' Go back to previous script state
Public Sub NavigateBack()
    ' Check if we have history to go back to
    If scriptHistory.count = 0 Then
        Exit Sub
    End If
    
    ' Get previous state
    currentScriptState = scriptHistory(scriptHistory.count)
    
    ' Remove from history
    If scriptHistory.count > 0 Then
        scriptHistory.Remove scriptHistory.count
    End If
    
    ' Update script content
    UpdateScriptContent
End Sub

' Get script content for each state
Private Function GetInitialGreeting() As String
    Dim greeting As String
    
    greeting = "Hi " & GetActiveCustomerName() & ", my name is " & Application.userName & ". Before I start, just a reminder that our calls are recorded for training purposes." & vbCrLf & vbCrLf
    greeting = greeting & "I understand you're interested in a novated lease. Is that correct?" & vbCrLf & vbCrLf
    greeting = greeting & "Great! I'd like to first understand your familiarity with novated leasing to make sure I explain things at the right level for you."
    
    GetInitialGreeting = greeting
End Function

Private Function GetKnowledgeAssessment() As String
    Dim content As String
    
    content = "How familiar are you with novated leasing?" & vbCrLf & vbCrLf
    content = content & "• Is this your first time looking into novated leasing?" & vbCrLf
    content = content & "• Do you have some knowledge about how it works?" & vbCrLf
    content = content & "• Are you very familiar with novated leasing?"
    
    GetKnowledgeAssessment = content
End Function

Private Function GetBeginnerEducation() As String
    Dim content As String
    
    content = "Thank you for letting me know. I'll make sure to explain novated leasing clearly." & vbCrLf & vbCrLf
    content = content & "A novated lease is a three-way agreement between you, your employer, and a finance company. It's a way to pay for a car and its running costs using a combination of your pre-tax and post-tax salary." & vbCrLf & vbCrLf
    content = content & "The main benefits include:" & vbCrLf
    content = content & "• Tax savings on part of your car payments" & vbCrLf
    content = content & "• GST savings on the purchase price and running costs" & vbCrLf
    content = content & "• Simplified budgeting with one regular payment" & vbCrLf
    content = content & "• No large upfront payment required" & vbCrLf & vbCrLf
    content = content & "Does this make sense so far?"
    
    GetBeginnerEducation = content
End Function

Private Function GetIntermediateEducation() As String
    Dim content As String
    
    content = "Great, since you have some familiarity with novated leasing, I'll focus on the key points." & vbCrLf & vbCrLf
    content = content & "A novated lease combines pre-tax and post-tax salary deductions to cover your vehicle and running costs, resulting in potential tax savings." & vbCrLf & vbCrLf
    content = content & "Key aspects include:" & vbCrLf
    content = content & "• GST savings on the purchase price and running costs" & vbCrLf
    content = content & "• Coverage of registration, insurance, servicing, and fuel" & vbCrLf
    content = content & "• A residual payment at the end of the lease" & vbCrLf
    content = content & "• Flexible end-of-lease options" & vbCrLf & vbCrLf
    content = content & "Would you like more details on any of these aspects?"
    
    GetIntermediateEducation = content
End Function

Private Function GetAdvancedEducation() As String
    Dim content As String
    
    content = "Excellent! Since you're already familiar with novated leasing, let's focus on your specific needs and how we can tailor a solution for you." & vbCrLf & vbCrLf
    content = content & "Just to confirm, you understand that a novated lease involves:" & vbCrLf
    content = content & "• Pre-tax and post-tax salary contributions" & vbCrLf
    content = content & "• GST benefits on purchase and running costs" & vbCrLf
    content = content & "• Comprehensive running cost management" & vbCrLf
    content = content & "• End-of-lease options like payout, refinance, or upgrade" & vbCrLf & vbCrLf
    content = content & "Is there any specific aspect of novated leasing you'd like me to elaborate on?"
    
    GetAdvancedEducation = content
End Function

Private Function GetQualifyingQuestions() As String
    Dim content As String
    
    content = "Now, to help me understand your needs better, I'd like to ask you a few questions:" & vbCrLf & vbCrLf
    content = content & "1. Have you ever had a novated lease before?" & vbCrLf
    content = content & "2. Do you have a specific car in mind? (new/used, electric or petrol)" & vbCrLf
    content = content & "3. Have you test driven the car you're interested in?" & vbCrLf
    content = content & "4. Do you have a budget in mind?" & vbCrLf
    content = content & "5. When would you like to be in the new vehicle?" & vbCrLf
    content = content & "6. How are you paying for your current car? (cash, finance, home loan)" & vbCrLf & vbCrLf
    content = content & "Which aspect would you like to discuss first?"
    
    GetQualifyingQuestions = content
End Function

Private Function GetVehicleNeedsQuestions() As String
    Dim content As String
    
    content = "Let's talk about your vehicle requirements:" & vbCrLf & vbCrLf
    content = content & "• What type of vehicle are you looking for? (SUV, sedan, electric, etc.)" & vbCrLf
    content = content & "• Do you have a specific make and model in mind?" & vbCrLf
    content = content & "• What features are most important to you?" & vbCrLf
    content = content & "• How many kilometers do you drive annually?" & vbCrLf
    content = content & "• Do you need any specific accessories? (tow bar, roof racks, etc.)" & vbCrLf & vbCrLf
    content = content & "This information will help me prepare an accurate quote for you."
    
    GetVehicleNeedsQuestions = content
End Function

Private Function GetFinancialQuestions() As String
    Dim content As String
    
    content = "Now I need to gather some financial information to create an accurate quote:" & vbCrLf & vbCrLf
    content = content & "1. What is your gross annual income?" & vbCrLf
    content = content & "2. For insurance purposes, can I confirm your postcode and suburb?" & vbCrLf
    content = content & "3. In the last 3 years, have you had 2 or more at-fault claims?" & vbCrLf
    content = content & "4. In the last 3 years, have you been charged with a DUI or negligent driving?" & vbCrLf
    content = content & "5. How many years would you like the lease to be? (typically 1-5 years)" & vbCrLf & vbCrLf
    content = content & "This information is used to calculate your potential savings and prepare your quote."
    
    GetFinancialQuestions = content
End Function

Private Function GetBenefitsExplanation() As String
    Dim content As String
    
    content = "Let me explain the benefits of a novated lease in more detail:" & vbCrLf & vbCrLf
    content = content & "The vehicle is fully maintained, meaning you'll get:" & vbCrLf
    content = content & "• Fuel cards for your fuel purchases" & vbCrLf
    content = content & "• Direct billing for servicing and maintenance" & vbCrLf
    content = content & "• Coverage for tires and registration renewals" & vbCrLf & vbCrLf
    content = content & "Other than tolls and fines, you shouldn't need to spend any additional money on your vehicle beyond the regular lease payments." & vbCrLf & vbCrLf
    content = content & "Your deductions go to an account held at SG Fleet. As you spend money on the vehicle, it deducts from that account, which you can monitor through our online app." & vbCrLf & vbCrLf
    content = content & "Throughout the lease, you're paying the vehicle down to a residual or balloon amount set by the ATO."
    
    GetBenefitsExplanation = content
End Function

Private Function GetLeaseEndOptions() As String
    Dim content As String
    
    content = "At the end of the lease, you'll have several options:" & vbCrLf & vbCrLf
    content = content & "1. If you love the car, you can pay out the residual amount and keep the car" & vbCrLf & vbCrLf
    content = content & "2. If you love the car and the lease arrangement, you can refinance and extend the lease" & vbCrLf & vbCrLf
    content = content & "3. You can sell the car yourself - any money you make above the residual amount is yours tax-free" & vbCrLf & vbCrLf
    content = content & "4. The most popular option is to trade in the car and upgrade to a new vehicle with a new lease" & vbCrLf & vbCrLf
    content = content & "We can also source the vehicle for you, with fleet discounts that we pass on to you. All our new quotes include an Eco Protection Pack and Minor Damage Repair membership."
    
    GetLeaseEndOptions = content
End Function

Private Function GetObjectionsResponses() As String
    Dim content As String
    
    content = "I understand you have some concerns. Let me address some common questions:" & vbCrLf & vbCrLf
    content = content & "If you're concerned about cost - While there are regular payments, the tax benefits and GST savings often make novated leasing more cost-effective than traditional car ownership." & vbCrLf & vbCrLf
    content = content & "If you're worried about changing jobs - If you change employers, you have options: transfer the lease to your new employer (if they participate), convert to a consumer loan, or pay it out." & vbCrLf & vbCrLf
    content = content & "If you're uncertain about the tax benefits - We can prepare a personalized calculation showing exactly how much you could save based on your salary and vehicle choice." & vbCrLf & vbCrLf
    content = content & "What specific concerns do you have that I can address?"
    
    GetObjectionsResponses = content
End Function

Private Function GetQuoteInfoCollection() As String
    Dim content As String
    
    content = "To generate an accurate quote, I need to confirm these details:" & vbCrLf & vbCrLf
    content = content & "Vehicle Information:" & vbCrLf
    content = content & "• Brand: " & vbCrLf
    content = content & "• Model: " & vbCrLf
    content = content & "• Variant: " & vbCrLf
    content = content & "• Color preference: " & vbCrLf
    content = content & "• Additional accessories: " & vbCrLf & vbCrLf
    content = content & "Lease Details:" & vbCrLf
    content = content & "• Lease term (years): " & vbCrLf
    content = content & "• Annual kilometers: " & vbCrLf & vbCrLf
    content = content & "Please confirm these details so I can prepare your quote."
    
    GetQuoteInfoCollection = content
End Function

Private Function GetClosingScript() As String
    Dim content As String
    
    content = "Thank you for providing all the information. Here's what happens next:" & vbCrLf & vbCrLf
    content = content & "I'll send you an indicative quote to give you an idea of how the lease will look. Along with the quote, you'll receive:" & vbCrLf
    content = content & "• A link to complete the finance pre-approval (valid for 6 months)" & vbCrLf
    content = content & "• Information about our trade-in program if applicable" & vbCrLf & vbCrLf
    content = content & "Most customers choose to get their finance pre-approved while we find a car for them, as the approval lasts for 6 months." & vbCrLf & vbCrLf
    content = content & "Once I receive the pricing (expected within 24 hours), I'll call you to discuss the next steps." & vbCrLf & vbCrLf
    content = content & "Do you have any final questions I can answer for you today?"
    
    GetClosingScript = content
End Function

Private Function GetFollowUpScript() As String
    Dim content As String
    
    content = "Let's schedule a follow-up to continue our conversation:" & vbCrLf & vbCrLf
    content = content & "• When would be a good time for me to call you back?" & vbCrLf
    content = content & "• Is there any specific information you'd like me to prepare before our next call?" & vbCrLf
    content = content & "• Would you prefer I email the information to you instead?" & vbCrLf & vbCrLf
    content = content & "I'll make a note of your preferences and ensure we follow up appropriately."
    
    GetFollowUpScript = content
End Function
