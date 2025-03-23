Attribute VB_Name = "Core"
Sub EnhanceDashboard()
    ' Enhance existing Dashboard
    Dim dashboardSheet As Worksheet
    
    ' Reference existing Dashboard sheet
    Set dashboardSheet = ThisWorkbook.Sheets("Dashboard")
    
    ' Clear any existing buttons that might conflict
    ClearExistingDashboardControls dashboardSheet
    
    ' Add quick action buttons
    AddQuickActionButtons dashboardSheet
    
    ' Create today's calls view with auto-refresh
    CreateTodaysCallsView dashboardSheet
    
    ' Create priority customers widget
    CreatePriorityCustomersWidget dashboardSheet
    
    ' Set up automatic refresh timer (every 15 minutes)
    Application.OnTime Now + TimeValue("00:15:00"), "RefreshDashboard"
    
    MsgBox "Dashboard enhanced successfully! It will auto-refresh every 15 minutes.", vbInformation
End Sub

Sub ClearExistingDashboardControls(dashboardSheet As Worksheet)
    ' Remove any existing buttons to prevent conflicts
    Dim btn As Button
    On Error Resume Next
    For Each btn In dashboardSheet.Buttons
        btn.Delete
    Next btn
    On Error GoTo 0
End Sub

Sub AddQuickActionButtons(dashboardSheet As Worksheet)
    Dim topPosition As Double
    topPosition = 75 ' Starting position for first button
    
    ' Create New Call button
    CreateActionButton dashboardSheet, "NewCallBtn", "Start New Call", "StartNewCall", topPosition
    topPosition = topPosition + 30
    
    ' Create New Quote button
    CreateActionButton dashboardSheet, "NewQuoteBtn", "Create Quote", "CreateNewQuote", topPosition
    topPosition = topPosition + 30
    
    ' Create Follow-up button
    CreateActionButton dashboardSheet, "FollowUpBtn", "Schedule Follow-up", "ScheduleFollowUp", topPosition
    topPosition = topPosition + 30
    
    ' Create Sync button
    CreateActionButton dashboardSheet, "SyncBtn", "Sync Outlook", "SyncWithOutlook", topPosition
End Sub

Sub CreateActionButton(dashboardSheet As Worksheet, buttonName As String, buttonCaption As String, macroName As String, topPosition As Double)
    Dim newButton As Object
    
    ' Create new button using Form Controls (more reliable)
    Set newButton = dashboardSheet.Shapes.AddFormControl(xlButtonControl, _
                                           left:=10, _
                                           top:=topPosition, _
                                           width:=120, _
                                           height:=25)
    With newButton
        .Name = buttonName
        .OnAction = macroName
        ' Get the TextFrame to set caption
        With .TextFrame
            .Characters.text = buttonCaption
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With
    End With
End Sub
Sub CreateTodaysCallsView(dashboardSheet As Worksheet)
    ' Set up area for today's calls
    dashboardSheet.Range("B5").Value = "TODAY'S CALLS"
    dashboardSheet.Range("B5").Font.Bold = True
    dashboardSheet.Range("B5").Font.Size = 14
    
    ' Add column headers
    dashboardSheet.Range("B6").Value = "Time"
    dashboardSheet.Range("C6").Value = "Customer"
    dashboardSheet.Range("D6").Value = "Phone"
    dashboardSheet.Range("E6").Value = "Status"
    dashboardSheet.Range("B6:E6").Font.Bold = True
    
    ' Add formulas to pull data from CallPlanner sheet
    For i = 1 To 10
        dashboardSheet.Range("B" & (6 + i)).Formula = "=IF(ROW()-6<=COUNTA(CallPlanner!A:A)-1,INDEX(CallPlanner!A:A,ROW()-5),"""")"
        dashboardSheet.Range("C" & (6 + i)).Formula = "=IF(B" & (6 + i) & "="""","""",INDEX(CallPlanner!B:B,ROW()-5))"
        dashboardSheet.Range("D" & (6 + i)).Formula = "=IF(B" & (6 + i) & "="""","""",INDEX(CallPlanner!C:C,ROW()-5))"
        dashboardSheet.Range("E" & (6 + i)).Formula = "=IF(B" & (6 + i) & "="""","""",INDEX(CallPlanner!G:G,ROW()-5))"
    Next i
    
    ' Add call buttons
    AddCallButtons dashboardSheet
End Sub

Sub AddCallButtons(dashboardSheet As Worksheet)
    ' Add call buttons next to each customer
    For i = 1 To 10
        Dim btn As Button
        Set btn = dashboardSheet.Buttons.Add(dashboardSheet.Range("F" & (6 + i)).left, _
                                      dashboardSheet.Range("F" & (6 + i)).top, _
                                      40, 18)
        With btn
            .Name = "CallBtn" & i
            .caption = "Call"
            .OnAction = "CallCustomerFromDashboard"
            .Visible = True
        End With
    Next i
End Sub

Sub CreatePriorityCustomersWidget(dashboardSheet As Worksheet)
    ' Set up area for priority customers
    dashboardSheet.Range("G5").Value = "PRIORITY CUSTOMERS"
    dashboardSheet.Range("G5").Font.Bold = True
    dashboardSheet.Range("G5").Font.Size = 14
    
    ' Add column headers
    dashboardSheet.Range("G6").Value = "Customer"
    dashboardSheet.Range("H6").Value = "Stage"
    dashboardSheet.Range("I6").Value = "Due Date"
    dashboardSheet.Range("J6").Value = "Status"
    dashboardSheet.Range("G6:J6").Font.Bold = True
    
    ' Add formulas to show priority customers
    ' These will display customers with upcoming follow-ups
    For i = 1 To 8
        dashboardSheet.Range("G" & (6 + i)).Formula = "=IF(ROW()-6<=COUNTA(CustomerTracker!B:B),INDEX(CustomerTracker!B:B,MATCH(LARGE(IF(CustomerTracker!H:H<=TODAY()+3,ROW(CustomerTracker!A:A)-1),ROW()-6),ROW(CustomerTracker!A:A)-1,0)),"""")"
        dashboardSheet.Range("H" & (6 + i)).Formula = "=IF(G" & (6 + i) & "="""","""",INDEX(CustomerTracker!E:E,MATCH(G" & (6 + i) & ",CustomerTracker!B:B,0)))"
        dashboardSheet.Range("I" & (6 + i)).Formula = "=IF(G" & (6 + i) & "="""","""",TEXT(INDEX(CustomerTracker!H:H,MATCH(G" & (6 + i) & ",CustomerTracker!B:B,0)),""d-mmm""))"
        dashboardSheet.Range("J" & (6 + i)).Formula = "=IF(G" & (6 + i) & "="""","""",IF(INDEX(CustomerTracker!H:H,MATCH(G" & (6 + i) & ",CustomerTracker!B:B,0))<TODAY(),""Overdue"",""Due""))"
    Next i
    
    ' Add conditional formatting for status column
    With dashboardSheet.Range("J7:J14").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Overdue""")
        .Interior.Color = RGB(255, 200, 200) ' Light red for overdue
    End With
    
    With dashboardSheet.Range("J7:J14").FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Due""")
        .Interior.Color = RGB(200, 255, 200) ' Light green for due
    End With
End Sub

Sub RefreshDashboard()
    ' Auto-refresh dashboard data
    Application.ScreenUpdating = False
    
    ' Update today's date on dashboard
    ThisWorkbook.Sheets("Dashboard").Range("K2").Value = "Last Updated: " & Format(Now, "dd-mmm-yyyy hh:mm")
    
    ' Check for any new emails or calendar items
    If InitializeOutlook() Then
        SyncCustomerEmails
    End If
    
    ' Update call status indicators
    UpdateCallStatusIndicators
    
    Application.ScreenUpdating = True
    
    ' Schedule next refresh
    Application.OnTime Now + TimeValue("00:15:00"), "RefreshDashboard"
End Sub

Sub UpdateCallStatusIndicators()
    ' Update status indicators for calls
    Dim dashboardSheet As Worksheet
    Set dashboardSheet = ThisWorkbook.Sheets("Dashboard")
    
    ' Update call progress
    Dim totalCalls As Integer
    Dim completedCalls As Integer
    
    totalCalls = Application.CountA(ThisWorkbook.Sheets("CallPlanner").Range("A:A")) - 1
    completedCalls = Application.CountIf(ThisWorkbook.Sheets("CallPlanner").Range("G:G"), "Completed")
    
    dashboardSheet.Range("L5").Value = "CALL PROGRESS"
    dashboardSheet.Range("L6").Value = "Completed: " & completedCalls & " / 50"
    
    ' Create visual indicator
    Dim progressPct As Double
    progressPct = completedCalls / 50
    
    ' Color coding based on progress
    If progressPct < 0.3 Then
        dashboardSheet.Range("L6").Interior.Color = RGB(255, 200, 200) ' Red if behind
    ElseIf progressPct < 0.7 Then
        dashboardSheet.Range("L6").Interior.Color = RGB(255, 255, 200) ' Yellow if on track
    Else
        dashboardSheet.Range("L6").Interior.Color = RGB(200, 255, 200) ' Green if ahead
    End If
End Sub
