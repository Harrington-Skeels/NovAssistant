Attribute VB_Name = "ReportGenerator"
Sub SetupTemplatesSheet(ws As Worksheet)
    ' Setup the basic structure for Templates sheet
    
    ' Add headers
    ws.Range("A1").Value = "Template Type"
    ws.Range("B1").Value = "Template Name"
    ws.Range("C1").Value = "Subject"
    ws.Range("D1").Value = "Content"
    
    ' Format headers
    ws.Range("A1:D1").Font.Bold = True
    ws.Range("A1:D1").Interior.Color = RGB(0, 66, 37) ' SG Fleet green
    ws.Range("A1:D1").Font.Color = RGB(255, 255, 255) ' White
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 15
    ws.Columns("B").ColumnWidth = 25
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 60
    
    ' Add some basic email templates
    ws.Range("A2").Value = "EmailTemplate"
    ws.Range("B2").Value = "Initial Contact"
    ws.Range("C2").Value = "Novated Lease Information for [Customer Name]"
    ws.Range("D2").Value = "Dear [Customer Name]," & vbCrLf & vbCrLf & _
                          "Thank you for your interest in novated leasing with SG Fleet." & vbCrLf & vbCrLf & _
                          "A novated lease offers significant benefits including:" & vbCrLf & _
                          "• Potential tax savings" & vbCrLf & _
                          "• GST savings on the purchase price" & vbCrLf & _
                          "• GST-free running costs" & vbCrLf & _
                          "• Simplified budgeting with one regular payment" & vbCrLf & vbCrLf & _
                          "I'd be happy to prepare a personalized quote to show you how much you could save." & vbCrLf & vbCrLf & _
                          "Please let me know if you have any questions." & vbCrLf & vbCrLf & _
                          "Kind regards,"
    
    ws.Range("A3").Value = "EmailTemplate"
    ws.Range("B3").Value = "Quote Follow-Up"
    ws.Range("C3").Value = "Follow-up on your Novated Lease Quote - [Vehicle]"
    ws.Range("D3").Value = "Dear [Customer Name]," & vbCrLf & vbCrLf & _
                          "I hope this email finds you well." & vbCrLf & vbCrLf & _
                          "I'm following up on the novated lease quote I sent through for the [Vehicle]. I wanted to check if you've had a chance to review it and if you have any questions." & vbCrLf & vbCrLf & _
                          "If you'd like to proceed or need any adjustments to the quote, please let me know." & vbCrLf & vbCrLf & _
                          "Kind regards,"
End Sub

Sub SetupSettingsSheet(ws As Worksheet)
    ' Setup the basic structure for Settings sheet
    
    ' Add headers
    ws.Range("A1").Value = "Setting"
    ws.Range("B1").Value = "Value"
    
    ' Format headers
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Interior.Color = RGB(0, 66, 37) ' SG Fleet green
    ws.Range("A1:B1").Font.Color = RGB(255, 255, 255) ' White
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 40
    
    ' Add basic settings
    ws.Range("A2").Value = "DynamicsURL"
    ws.Range("B2").Value = "https://sgfleet.crm.dynamics.com"
    
    ws.Range("A3").Value = "DynamicsUser"
    ws.Range("B3").Value = Application.userName
    
    ws.Range("A4").Value = "CallTarget"
    ws.Range("B4").Value = "50"
    
    ws.Range("A5").Value = "AutoSyncInterval"
    ws.Range("B5").Value = "15" ' Minutes
    
    ' Hide the sheet to prevent accidental changes
    ws.Visible = xlSheetVeryHidden
End Sub

Sub SetupBasicData()
    ' Add some sample data to make the system look functional
    
    ' Add sample customers
    AddSampleCustomers
    
    ' Add sample calls
    AddSampleCalls
    
    ' Add sample contact history
    AddSampleContactHistory
    
    ' Add sample quotes
    AddSampleQuotes
End Sub

Sub AddSampleCustomers()
    ' Add sample customers to CustomerTracker
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CustomerTracker")
    
    ' Check if sheet already has data
    If Not IsEmpty(ws.Range("A2").Value) Then Exit Sub
    
    ' Sample customer data
    Dim customerData(1 To 5, 1 To 15) As Variant
    
    ' CustomerID, Name, Email, Phone, Stage, LastContact, NextAction, NextActionDate,
    ' Vehicle, LeaseTerm, Notes, Documents, Source, Status, CRMID
    
    customerData(1, 1) = "CUST-25001"
    customerData(1, 2) = "John Smith"
    customerData(1, 3) = "john.smith@example.com"
    customerData(1, 4) = "(02) 9876 5432"
    customerData(1, 5) = "Quote Sent"
    customerData(1, 6) = DateValue("5-Mar-2025")
    customerData(1, 7) = "Follow-up Call"
    customerData(1, 8) = DateValue("8-Mar-2025")
    customerData(1, 9) = "VOLVO EX30 Ultra"
    customerData(1, 10) = "36"
    customerData(1, 11) = "- Initially interested in Polestar 4 but found Volvo EX30 better value" & vbCrLf & _
                         "- Wants to upgrade from current Toyota Camry, finance ending in 2 months"
    customerData(1, 12) = ""
    customerData(1, 13) = "Website"
    customerData(1, 14) = "Hot"
    customerData(1, 15) = ""
    
    customerData(2, 1) = "CUST-25002"
    customerData(2, 2) = "Jane Doe"
    customerData(2, 3) = "jane.doe@example.com"
    customerData(2, 4) = "(02) 1234 5678"
    customerData(2, 5) = "Initial Call"
    customerData(2, 6) = DateValue("4-Mar-2025")
    customerData(2, 7) = "Prepare Quote"
    customerData(2, 8) = DateValue("6-Mar-2025")
    customerData(2, 9) = "Toyota RAV4 Hybrid"
    customerData(2, 10) = ""
    customerData(2, 11) = "- Looking for a family SUV" & vbCrLf & _
                         "- Current car is 6 years old and starting to have issues"
    customerData(2, 12) = ""
    customerData(2, 13) = "Referral"
    customerData(2, 14) = "Warm"
    customerData(2, 15) = ""
    
    customerData(3, 1) = "CUST-25003"
    customerData(3, 2) = "Bob Johnson"
    customerData(3, 3) = "bob.johnson@example.com"
    customerData(3, 4) = "(02) 2345 6789"
    customerData(3, 5) = "Finance Application"
    customerData(3, 6) = DateValue("3-Mar-2025")
    customerData(3, 7) = "Check Application Status"
    customerData(3, 8) = DateValue("7-Mar-2025")
    customerData(3, 9) = "BMW X3 M Sport"
    customerData(3, 10) = "48"
    customerData(3, 11) = "- Very specific about features" & vbCrLf & _
                         "- Wants the M Sport package" & vbCrLf & _
                         "- Finance application submitted on 3-Mar"
    customerData(3, 12) = ""
    customerData(3, 13) = "CRM"
    customerData(3, 14) = "Hot"
    customerData(3, 15) = "CRM-12345"
    
    customerData(4, 1) = "CUST-25004"
    customerData(4, 2) = "Sarah Williams"
    customerData(4, 3) = "sarah.williams@example.com"
    customerData(4, 4) = "(02) 3456 7890"
    customerData(4, 5) = "Vehicle Procurement"
    customerData(4, 6) = DateValue("2-Mar-2025")
    customerData(4, 7) = "Confirm Delivery Date"
    customerData(4, 8) = DateValue("9-Mar-2025")
    customerData(4, 9) = "Kia EV6 GT-Line"
    customerData(4, 10) = "36"
    customerData(4, 11) = "- Very excited about EV" & vbCrLf & _
                         "- Dealer confirmed order, awaiting delivery date"
    customerData(4, 12) = ""
    customerData(4, 13) = "Website"
    customerData(4, 14) = "Hot"
    customerData(4, 15) = ""
    
    customerData(5, 1) = "CUST-25005"
    customerData(5, 2) = "Michael Brown"
    customerData(5, 3) = "michael.brown@example.com"
    customerData(5, 4) = "(02) 4567 8901"
    customerData(5, 5) = "Settlement"
    customerData(5, 6) = DateValue("1-Mar-2025")
    customerData(5, 7) = "Schedule Delivery"
    customerData(5, 8) = DateValue("10-Mar-2025")
    customerData(5, 9) = "Mazda CX-5 Touring"
    customerData(5, 10) = "24"
    customerData(5, 11) = "- Paperwork completed" & vbCrLf & _
                         "- Finance approved" & vbCrLf & _
                         "- Vehicle ready for delivery"
    customerData(5, 12) = ""
    customerData(5, 13) = "Referral"
    customerData(5, 14) = "Hot"
    customerData(5, 15) = ""
    
    ' Add data to sheet
    ws.Range("A2:O6").Value = customerData
End Sub

Sub AddSampleCalls()
    ' Add sample calls to CallPlanner
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CallPlanner")
    
    ' Check if sheet already has data
    If Not IsEmpty(ws.Range("A2").Value) Then Exit Sub
    
    ' Sample call data
    Dim callData(1 To 5, 1 To 8) As Variant
    
    ' Time, Customer, Phone, Purpose, Stage, Status, Outcome, Notes
    
    callData(1, 1) = "9:00 AM"
    callData(1, 2) = "John Smith"
    callData(1, 3) = "(02) 9876 5432"
    callData(1, 4) = "Quote Follow-up"
    callData(1, 5) = "Quote Sent"
    callData(1, 6) = "Hot"
    callData(1, 7) = "Pending"
    callData(1, 8) = ""
    
    callData(2, 1) = "10:15 AM"
    callData(2, 2) = "Jane Doe"
    callData(2, 3) = "(02) 1234 5678"
    callData(2, 4) = "Prepare Quote"
    callData(2, 5) = "Initial Call"
    callData(2, 6) = "Warm"
    callData(2, 7) = "Pending"
    callData(2, 8) = ""
    
    callData(3, 1) = "11:30 AM"
    callData(3, 2) = "Bob Johnson"
    callData(3, 3) = "(02) 2345 6789"
    callData(3, 4) = "Check Application Status"
    callData(3, 5) = "Finance Application"
    callData(3, 6) = "Hot"
    callData(3, 7) = "Pending"
    callData(3, 8) = ""
    
    callData(4, 1) = "1:45 PM"
    callData(4, 2) = "Sarah Williams"
    callData(4, 3) = "(02) 3456 7890"
    callData(4, 4) = "Confirm Delivery Date"
    callData(4, 5) = "Vehicle Procurement"
    callData(4, 6) = "Hot"
    callData(4, 7) = "Pending"
    callData(4, 8) = ""
    
    callData(5, 1) = "3:00 PM"
    callData(5, 2) = "Michael Brown"
    callData(5, 3) = "(02) 4567 8901"
    callData(5, 4) = "Schedule Delivery"
    callData(5, 5) = "Settlement"
    callData(5, 6) = "Hot"
    callData(5, 7) = "Pending"
    callData(5, 8) = ""
    
    ' Add data to sheet
    ws.Range("A2:H6").Value = callData
End Sub

Sub AddSampleContactHistory()
    ' Add sample contact history
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ContactHistory")
    
    ' Check if sheet already has data
    If Not IsEmpty(ws.Range("A2").Value) Then Exit Sub
    
    ' Sample contact data
    Dim contactData(1 To 8, 1 To 5) As Variant
    
    ' Customer, ContactType, Details, DateTime, User
    
    contactData(1, 1) = "John Smith"
    contactData(1, 2) = "Outbound Call"
    contactData(1, 3) = "Sent quote for Volvo EX30 with monthly payment of $1,269.49. Customer was interested, mentioned they will review with spouse tonight."
    contactData(1, 4) = Now() - 3 - TimeValue("2:00:00")
    contactData(1, 5) = Application.userName
    
    contactData(2, 1) = "John Smith"
    contactData(2, 2) = "Email Sent"
    contactData(2, 3) = "Quote for Polestar 4 with monthly payment of $1,441.60."
    contactData(2, 4) = Now() - 3 - TimeValue("2:45:00")
    contactData(2, 5) = Application.userName
    
    contactData(3, 1) = "John Smith"
    contactData(3, 2) = "Outbound Call"
    contactData(3, 3) = "Initial consultation. Customer interested in EV options. Gathered salary information and employer details. Educated customer on novated lease benefits."
    contactData(3, 4) = Now() - 4 - TimeValue("9:15:00")
    contactData(3, 5) = Application.userName
    
    contactData(4, 1) = "Jane Doe"
    contactData(4, 2) = "Outbound Call"
    contactData(4, 3) = "Initial consultation. Customer looking for a hybrid SUV. Gathered basic information and explained novated lease benefits."
    contactData(4, 4) = Now() - 2 - TimeValue("11:30:00")
    contactData(4, 5) = Application.userName
    
    contactData(5, 1) = "Bob Johnson"
    contactData(5, 2) = "Outbound Call"
    contactData(5, 3) = "Confirmed receipt of finance application. Explained next steps and timeline for approval."
    contactData(5, 4) = Now() - 1 - TimeValue("10:00:00")
    contactData(5, 5) = Application.userName
    
    contactData(6, 1) = "Bob Johnson"
    contactData(6, 2) = "Email Sent"
    contactData(6, 3) = "Finance application confirmation and details of next steps."
    contactData(6, 4) = Now() - 1 - TimeValue("10:15:00")
    contactData(6, 5) = Application.userName
    
    contactData(7, 1) = "Sarah Williams"
    contactData(7, 2) = "Outbound Call"
    contactData(7, 3) = "Checked status of vehicle order. Dealer confirmed vehicle has been ordered and is expected in 3-4 weeks."
    contactData(7, 4) = Now() - 2 - TimeValue("14:00:00")
    contactData(7, 5) = Application.userName
    
    contactData(8, 1) = "Michael Brown"
    contactData(8, 2) = "Outbound Call"
    contactData(8, 3) = "All paperwork complete. Vehicle ready for delivery. Customer confirmed availability for delivery next Tuesday."
    contactData(8, 4) = Now() - 1 - TimeValue("15:30:00")
    contactData(8, 5) = Application.userName
    
    ' Add data to sheet
    ws.Range("A2:E9").Value = contactData
End Sub

Sub AddSampleQuotes()
    ' Add sample quotes to QuoteHistory
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("QuoteHistory")
    
    ' Check if sheet already has data
    If Not IsEmpty(ws.Range("A2").Value) Then Exit Sub
    
    ' Sample quote data
    Dim quoteData(1 To 5, 1 To 8) As Variant
    
    ' QuoteID, Customer, Vehicle, Date, MonthlyTotal, AnnualSavings, LeaseTerm, PDFPath
    
    quoteData(1, 1) = "250305-8BTRPR"
    quoteData(1, 2) = "John Smith"
    quoteData(1, 3) = "VOLVO EX30 Ultra"
    quoteData(1, 4) = DateValue("5-Mar-2025")
    quoteData(1, 5) = 1269.49
    quoteData(1, 6) = 11574.23
    quoteData(1, 7) = "36"
    quoteData(1, 8) = "C:\Quotes\John_Smith_VOLVO_EX30_Ultra.pdf"
    
    quoteData(2, 1) = "250305-7JKLMN"
    quoteData(2, 2) = "John Smith"
    quoteData(2, 3) = "Polestar 4 Dual Motor"
    quoteData(2, 4) = DateValue("5-Mar-2025")
    quoteData(2, 5) = 1441.6
    quoteData(2, 6) = 12845.78
    quoteData(2, 7) = "36"
    quoteData(2, 8) = "C:\Quotes\John_Smith_Polestar_4.pdf"
    
    quoteData(3, 1) = "250304-6GHJKL"
    quoteData(3, 2) = "Bob Johnson"
    quoteData(3, 3) = "BMW X3 M Sport"
    quoteData(3, 4) = DateValue("4-Mar-2025")
    quoteData(3, 5) = 1512.75
    quoteData(3, 6) = 13245.32
    quoteData(3, 7) = "48"
    quoteData(3, 8) = "C:\Quotes\Bob_Johnson_BMW_X3.pdf"
    
    quoteData(4, 1) = "250304-5FGHJK"
    quoteData(4, 2) = "Sarah Williams"
    quoteData(4, 3) = "Kia EV6 GT-Line"
    quoteData(4, 4) = DateValue("4-Mar-2025")
    quoteData(4, 5) = 1328.95
    quoteData(4, 6) = 12452.64
    quoteData(4, 7) = "36"
    quoteData(4, 8) = "C:\Quotes\Sarah_Williams_Kia_EV6.pdf"
    
    quoteData(5, 1) = "250303-4EFGHI"
    quoteData(5, 2) = "Michael Brown"
    quoteData(5, 3) = "Mazda CX-5 Touring"
    quoteData(5, 4) = DateValue("3-Mar-2025")
    quoteData(5, 5) = 986.45
    quoteData(5, 6) = 9245.36
    quoteData(5, 7) = "24"
    quoteData(5, 8) = "C:\Quotes\Michael_Brown_Mazda_CX5.pdf"
    
    ' Add data to sheet
    ws.Range("A2:H6").Value = quoteData
    
    ' Format currency columns
    ws.Range("E2:F6").NumberFormat = "$#,##0.00"
End Sub

Sub CreateMenuSystem()
    ' Create a simple menu system for all NovAssistant features
    Dim menuSheet As Worksheet
    Dim sheetExists As Boolean
    sheetExists = False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Menu" Then
            sheetExists = True
            Set menuSheet = ws
            Exit For
        End If
    Next ws
    
    If Not sheetExists Then
        ' Create new sheet
        Set menuSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        menuSheet.Name = "Menu"
    Else
        ' Clear existing content
        menuSheet.Cells.Clear
    End If
    
    ' Set up color variables
    Dim primaryColor As Long, secondaryColor As Long, accentColor As Long, textColor As Long
    primaryColor = RGB(0, 66, 37)      ' Dark green - SG Fleet color
    secondaryColor = RGB(245, 245, 245) ' Light gray
    accentColor = RGB(255, 152, 0)     ' Orange
    textColor = RGB(51, 51, 51)        ' Dark gray
    
    ' Create header
    With menuSheet.Range("B1:K1")
        .Merge
        .Value = "NOVASSISTANT MENU"
        .Font.Size = 20
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = primaryColor
        .Font.Color = RGB(255, 255, 255) ' White text
        .RowHeight = 40
    End With
    
    ' Add menu items
    CreateMenuItem menuSheet, "B3", "G5", "Dashboard", "Navigate to the Dashboard to see today's calls and performance metrics", "ViewDashboard", primaryColor
    
    CreateMenuItem menuSheet, "H3", "K5", "Customer Profile", "View detailed customer information and contact history", "ViewCustomerProfile", primaryColor
    
    CreateMenuItem menuSheet, "B6", "G8", "Call Assistant", "Use the phone script navigator for better call handling", "ViewCallAssistant", primaryColor
    
    CreateMenuItem menuSheet, "H6", "K8", "Create Quote", "Create a new novated lease quote", "CreateNewQuote", primaryColor
    
    CreateMenuItem menuSheet, "B9", "G11", "Customer Tracker", "View and manage all customer records", "ViewCustomerTracker", primaryColor
    
    CreateMenuItem menuSheet, "H9", "K11", "Call Planner", "View today's scheduled calls", "ViewCallPlanner", primaryColor
    
    CreateMenuItem menuSheet, "B12", "G14", "Import from CRM", "Import leads from Dynamics CRM", "SyncFromCRM", primaryColor
    
    CreateMenuItem menuSheet, "H12", "K14", "Prepare CRM Export", "Prepare data for export to CRM", "PrepareExportToCRM", primaryColor
    
    CreateMenuItem menuSheet, "B15", "G17", "Email Templates", "View and edit email templates", "ViewTemplates", primaryColor
    
    CreateMenuItem menuSheet, "H15", "K17", "Settings", "Configure NovAssistant settings", "ViewSettings", primaryColor
    
    ' Add usage instructions
    With menuSheet.Range("B19:K22")
        .Merge
        .Value = "HOW TO USE NOVASSISTANT" & vbCrLf & vbCrLf & _
                "Click on any menu item above to navigate to that feature. The Dashboard is your starting point for daily activities." & vbCrLf & _
                "The system will automatically track your calls, quotes, and customer interactions to help you reach your target of 50 calls per day."
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = secondaryColor
    End With
    
    ' Activate menu sheet
    menuSheet.Activate
    menuSheet.Range("A1").Select
End Sub

Sub CreateMenuItem(ws As Worksheet, startCell As String, endCell As String, title As String, description As String, macroName As String, buttonColor As Long)
    ' Creates a menu item as a button with title and description
    
    ' Merge range for item
    ws.Range(startCell & ":" & endCell).Merge
    
    ' Add button over the merged range
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range(startCell).left, ws.Range(startCell).top, _
                           ws.Range(endCell).left + ws.Range(endCell).width - ws.Range(startCell).left, _
                           ws.Range(endCell).top + ws.Range(endCell).height - ws.Range(startCell).top)
    
    With btn
        .Name = "menuBtn" & title
        .caption = title & vbCrLf & vbCrLf & description
        .OnAction = macroName
        .Font.Size = 11
        .Font.Bold = True
    End With
End Sub

' Navigation functions for menu system
Sub ViewDashboard()
    ThisWorkbook.Sheets("Dashboard").Activate
End Sub

Sub ViewCustomerProfile()
    ThisWorkbook.Sheets("CustomerProfile").Activate
End Sub

Sub ViewCallAssistant()
    ThisWorkbook.Sheets("CallAssistant").Activate
End Sub

Sub ViewCustomerTracker()
    ThisWorkbook.Sheets("CustomerTracker").Activate
End Sub

Sub ViewCallPlanner()
    ThisWorkbook.Sheets("CallPlanner").Activate
End Sub

Sub ViewTemplates()
    ThisWorkbook.Sheets("Templates").Activate
End Sub

Sub ViewSettings()
    ' Check if user has permission to view settings
    Dim pwd As String
    pwd = InputBox("Enter admin password to view settings:", "Settings Access")
    
    If pwd = "admin" Then ' Simple password for example
        ' Unhide settings sheet
        ThisWorkbook.Sheets("Settings").Visible = xlSheetVisible
        ThisWorkbook.Sheets("Settings").Activate
    Else
        MsgBox "Incorrect password. Settings access denied.", vbExclamation
    End If
End Sub
Sub SetupCompleteNovAssistant()
    ' One-click setup for the entire NovAssistant system
    
    ' Show welcome message
    MsgBox "Welcome to NovAssistant Setup! This will create a complete system for managing your novated lease sales.", vbInformation
    
    Application.StatusBar = "Setting up NovAssistant... Creating required sheets"
    
    ' Create required sheets
    EnsureRequiredSheets
    
    ' Add sample data
    SetupBasicData
    
    ' Add Dynamics buttons to Dashboard
    Dim dashboardSheet As Worksheet
    Set dashboardSheet = ThisWorkbook.Sheets("Dashboard")
    
    ' Add CRM buttons
    AddSafeCRMButtons dashboardSheet
    
    ' Show completion message
    Application.StatusBar = False
    MsgBox "NovAssistant setup complete! You can now use all features from your Dashboard.", vbInformation
End Sub
Sub EnsureRequiredSheets()
    ' Make sure all required sheets exist
    Dim sheetsToCreate As Variant
    sheetsToCreate = Array("CustomerTracker", "CallPlanner", "ContactHistory", "QuoteHistory", "Templates", "Settings")
    
    For Each sheetName In sheetsToCreate
        Dim sheetExists As Boolean
        sheetExists = False
        
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = sheetName Then
                sheetExists = True
                Exit For
            End If
        Next ws
        
        If Not sheetExists Then
            ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)).Name = sheetName
            
            ' Setup basic structure for each sheet
            Select Case sheetName
                Case "CustomerTracker"
                    SetupCustomerTrackerSheet ThisWorkbook.Sheets(sheetName)
                Case "CallPlanner"
                    SetupCallPlannerSheet ThisWorkbook.Sheets(sheetName)
                Case "ContactHistory"
                    SetupContactHistorySheet ThisWorkbook.Sheets(sheetName)
                Case "QuoteHistory"
                    SetupQuoteHistorySheet ThisWorkbook.Sheets(sheetName)
                Case "Templates"
                    SetupTemplatesSheet ThisWorkbook.Sheets(sheetName)
                Case "Settings"
                    SetupSettingsSheet ThisWorkbook.Sheets(sheetName)
            End Select
        End If
    Next sheetName
End Sub
Sub AddSafeCRMButtons(dashboardSheet As Worksheet)
    Dim crmPosition As Double
    crmPosition = 250 ' Starting position for CRM buttons
    
    ' Create Import from CRM button
    CreateActionButton dashboardSheet, "ImportFromCRMBtn", "Import from CRM", "SyncFromCRM", crmPosition
    crmPosition = crmPosition + 30
    
    ' Create Compare with CRM button
    CreateActionButton dashboardSheet, "CompareCRMBtn", "Compare with CRM", "CompareWithCRM", crmPosition
    crmPosition = crmPosition + 30
    
    ' Create Prepare Export button
    CreateActionButton dashboardSheet, "PrepareExportBtn", "Prepare CRM Export", "PrepareExportToCRM", crmPosition
End Sub
Sub SetupCustomerTrackerSheet(ws As Worksheet)
    ' Add headers
    ws.Range("A1").Value = "Customer ID"
    ws.Range("B1").Value = "Customer Name"
    ws.Range("C1").Value = "Email"
    ws.Range("D1").Value = "Phone"
    ws.Range("E1").Value = "Stage"
    ws.Range("F1").Value = "Last Contact"
    ws.Range("G1").Value = "Next Action"
    ws.Range("H1").Value = "Next Action Date"
    ws.Range("I1").Value = "Vehicle"
    ws.Range("J1").Value = "Lease Term"
    ws.Range("K1").Value = "Notes"
    ws.Range("L1").Value = "Documents"
    ws.Range("M1").Value = "Source"
    ws.Range("N1").Value = "Status"
    ws.Range("O1").Value = "CRM ID"
    
    ' Format headers
    ws.Range("A1:O1").Font.Bold = True
    ws.Range("A1:O1").Interior.Color = RGB(0, 66, 37)
    ws.Range("A1:O1").Font.Color = RGB(255, 255, 255)
End Sub

Sub SetupCallPlannerSheet(ws As Worksheet)
    ' Add headers
    ws.Range("A1").Value = "Time"
    ws.Range("B1").Value = "Customer"
    ws.Range("C1").Value = "Phone"
    ws.Range("D1").Value = "Purpose"
    ws.Range("E1").Value = "Stage"
    ws.Range("F1").Value = "Status"
    ws.Range("G1").Value = "Outcome"
    ws.Range("H1").Value = "Notes"
    
    ' Format headers
    ws.Range("A1:H1").Font.Bold = True
    ws.Range("A1:H1").Interior.Color = RGB(0, 66, 37)
    ws.Range("A1:H1").Font.Color = RGB(255, 255, 255)
End Sub

Sub SetupContactHistorySheet(ws As Worksheet)
    ' Add headers
    ws.Range("A1").Value = "Customer"
    ws.Range("B1").Value = "Contact Type"
    ws.Range("C1").Value = "Details"
    ws.Range("D1").Value = "Date/Time"
    ws.Range("E1").Value = "User"
    
    ' Format headers
    ws.Range("A1:E1").Font.Bold = True
    ws.Range("A1:E1").Interior.Color = RGB(0, 66, 37)
    ws.Range("A1:E1").Font.Color = RGB(255, 255, 255)
End Sub

Sub SetupQuoteHistorySheet(ws As Worksheet)
    ' Add headers
    ws.Range("A1").Value = "Quote ID"
    ws.Range("B1").Value = "Customer"
    ws.Range("C1").Value = "Vehicle"
    ws.Range("D1").Value = "Date"
    ws.Range("E1").Value = "Monthly Total"
    ws.Range("F1").Value = "Annual Savings"
    ws.Range("G1").Value = "Lease Term"
    ws.Range("H1").Value = "PDF Path"
    
    ' Format headers
    ws.Range("A1:H1").Font.Bold = True
    ws.Range("A1:H1").Interior.Color = RGB(0, 66, 37)
    ws.Range("A1:H1").Font.Color = RGB(255, 255, 255)
End Sub
