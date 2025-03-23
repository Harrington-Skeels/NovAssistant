Attribute VB_Name = "DynamicScriptUI"
' ====================================================================
' DynamicScriptUI Module - Version #1
' ====================================================================
' This module creates an improved user interface for the dynamic script system
' with a more intuitive, form-like experience within Excel

Option Explicit

' Script UI constants
Private Const THEME_PRIMARY = RGB(0, 66, 37)     ' Dark green
Private Const THEME_SECONDARY = RGB(39, 123, 77) ' Medium green
Private Const THEME_ACCENT = RGB(255, 153, 0)    ' Orange
Private Const THEME_LIGHT = RGB(242, 242, 242)   ' Light gray
Private Const THEME_WHITE = RGB(255, 255, 255)   ' White
Private Const THEME_TEXT_DARK = RGB(51, 51, 51)  ' Dark gray text
Private Const THEME_TEXT_LIGHT = RGB(255, 255, 255) ' White text

' Improved script creation function
Public Sub CreateModernScriptUI()
    Dim scriptSheet As Worksheet
    
    ' Check if script sheet exists
    On Error Resume Next
    Set scriptSheet = ThisWorkbook.Sheets("ModernScript")
    On Error GoTo 0
    
    ' Create new sheet if it doesn't exist
    If scriptSheet Is Nothing Then
        Set scriptSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
        scriptSheet.Name = "ModernScript"
    End If
    
    ' Clear sheet
    scriptSheet.Cells.Clear
    
    ' Configure column widths for layout
    scriptSheet.Columns("A").ColumnWidth = 2  ' Left margin
    scriptSheet.Columns("B:J").ColumnWidth = 10 ' Content columns
    scriptSheet.Columns("K").ColumnWidth = 2  ' Right margin
    
    ' Create header section
    With scriptSheet.Range("B1:J2")
        .Merge
        .Value = "NOVATED LEASE CONVERSATION ASSISTANT"
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = THEME_PRIMARY
        .Font.Color = THEME_TEXT_LIGHT
        .RowHeight = 30
    End With
    
    ' Create customer info panel
    CreateCustomerPanel scriptSheet
    
    ' Create script view panel
    CreateScriptViewPanel scriptSheet
    
    ' Create response panel
    CreateResponsePanel scriptSheet
    
    ' Create notes panel
    CreateNotesPanel scriptSheet
    
    ' Create action buttons
    CreateActionButtons scriptSheet
    
    ' Add keyboard shortcut handler
    AddKeyboardHandler scriptSheet
    
    ' Select cell A1 to avoid accidental editing
    scriptSheet.Range("A1").Select
    
    ' Apply protection to prevent accidental changes
    ApplySheetProtection scriptSheet
End Sub

' Create customer information panel
Private Sub CreateCustomerPanel(ws As Worksheet)
    ' Header
    With ws.Range("B4:J4")
        .Merge
        .Value = "CUSTOMER INFORMATION"
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Interior.Color = THEME_SECONDARY
        .Font.Color = THEME_TEXT_LIGHT
    End With
    
    ' Create two-column layout for customer details
    ws.Range("B5").Value = "Name:"
    ws.Range("B6").Value = "Phone:"
    ws.Range("B7").Value = "Email:"
    ws.Range("B8").Value = "Status:"
    
    ' Make labels bold
    ws.Range("B5:B8").Font.Bold = True
    ws.Range("B5:B8").IndentLevel = 1
    
    ' Create input fields
    With ws.Range("C5:E5")
        .Merge
        .Interior.Color = THEME_LIGHT
        .Name = "CustomerName"
    End With
    
    With ws.Range("C6:E6")
        .Merge
        .Interior.Color = THEME_LIGHT
        .Name = "CustomerPhone"
    End With
    
    With ws.Range("C7:E7")
        .Merge
        .Interior.Color = THEME_LIGHT
        .Name = "CustomerEmail"
    End With
    
    With ws.Range("C8:E8")
        .Merge
        .Interior.Color = THEME_LIGHT
        .Name = "CustomerStatus"
    End With
    
    ' Right side customer details
    ws.Range("F5").Value = "Stage:"
    ws.Range("F6").Value = "Duration:"
    ws.Range("F7").Value = "Next Action:"
    ws.Range("F8").Value = "Due Date:"
    
    ' Make labels bold
    ws.Range("F5:F8").Font.Bold = True
    ws.Range("F5:F8").IndentLevel = 1
    
    ' Create input fields
    With ws.Range("G5:J5")
        .Merge
        .Interior.Color = THEME_LIGHT
        .Name = "CustomerStage"
    End With
    
    With ws.Range("G6:J6")
        .Merge
        .Interior.Color = THEME_LIGHT
        .Name = "CallDuration"
        .Value = "00:00:00"
    End With
    
    With ws.Range("G7:J7")
        .Merge
        .Interior.Color = THEME_LIGHT
        .Name = "NextAction"
    End With
    
    With ws.Range("G8:J8")
        .Merge
        .Interior.Color = THEME_LIGHT
        .Name = "DueDate"
    End With
    
    ' Add border to panel
    ws.Range("B4:J8").BorderAround Weight:=xlMedium, ColorIndex:=1
End Sub

' Create script view panel
Private Sub CreateScriptViewPanel(ws As Worksheet)
    ' Header
    With ws.Range("B10:J10")
        .Merge
        .Value = "SCRIPT VIEW"
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Interior.Color = THEME_SECONDARY
        .Font.Color = THEME_TEXT_LIGHT
    End With
    
    ' Script navigation breadcrumb
    With ws.Range("B11:J11")
        .Merge
        .Value = "Current Path: Initial Greeting"
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .Interior.Color = THEME_LIGHT
        .Name = "ScriptPath"
    End With
    
    ' Script content area
    With ws.Range("B12:J19")
        .Merge
        .Font.Size = 11
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Interior.Color = THEME_WHITE
        .Name = "ScriptContent"
        .Value = "Script content will appear here when you start a call."
    End With
    
    ' Add border to panel
    ws.Range("B10:J19").BorderAround Weight:=xlMedium, ColorIndex:=1
End Sub

' Create response options panel
Private Sub CreateResponsePanel(ws As Worksheet)
    ' Header
    With ws.Range("B21:J21")
        .Merge
        .Value = "CUSTOMER RESPONSE"
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Interior.Color = THEME_SECONDARY
        .Font.Color = THEME_TEXT_LIGHT
    End With
    
    ' Response buttons area
    With ws.Range("B22:J27")
        .Merge
        .Interior.Color = THEME_LIGHT
        .Name = "ResponseArea"
    End With
    
    ' Add border to panel
    ws.Range("B21:J27").BorderAround Weight:=xlMedium, ColorIndex:=1
End Sub

' Create notes panel
Private Sub CreateNotesPanel(ws As Worksheet)
    ' Header
    With ws.Range("B29:J29")
        .Merge
        .Value = "CALL NOTES"
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .Interior.Color = THEME_SECONDARY
        .Font.Color = THEME_TEXT_LIGHT
    End With
    
    ' Notes area
    With ws.Range("B30:J36")
        .Merge
        .WrapText = True
        .Font.Size = 11
        .VerticalAlignment = xlTop
        .Interior.Color = THEME_WHITE
        .Name = "NotesArea"
    End With
    
    ' Add border to panel
    ws.Range("B29:J36").BorderAround Weight:=xlMedium, ColorIndex:=1
End Sub

' Create action buttons
Private Sub CreateActionButtons(ws As Worksheet)
    Dim startBtn As Button, endBtn As Button, saveBtn As Button
    Dim followUpBtn As Button, updateBtn As Button
    
    ' Row for buttons
    Dim btnRow As Long
    btnRow = 38
    
    ' Create Start Call button
    Set startBtn = ws.Buttons.Add(ws.Range("B" & btnRow).left, ws.Range("B" & btnRow).top, 80, 25)
    With startBtn
        .caption = "Start Call"
        .Name = "StartCallBtn"
        .OnAction = "StartModernCall"
        .Font.Size = 9
        .Font.Bold = True
    End With
    
    ' Create End Call button
    Set endBtn = ws.Buttons.Add(ws.Range("D" & btnRow).left, ws.Range("D" & btnRow).top, 80, 25)
    With endBtn
        .caption = "End Call"
        .Name = "EndCallBtn"
        .OnAction = "EndModernCall"
        .Font.Size = 9
        .Font.Bold = True
    End With
    
    ' Create Save Notes button
    Set saveBtn = ws.Buttons.Add(ws.Range("F" & btnRow).left, ws.Range("F" & btnRow).top, 80, 25)
    With saveBtn
        .caption = "Save Notes"
        .Name = "SaveNotesBtn"
        .OnAction = "SaveCallNotes"
        .Font.Size = 9
        .Font.Bold = True
    End With
    
    ' Create Follow-up button
    Set followUpBtn = ws.Buttons.Add(ws.Range("H" & btnRow).left, ws.Range("H" & btnRow).top, 80, 25)
    With followUpBtn
        .caption = "Schedule Follow-up"
        .Name = "FollowUpBtn"
        .OnAction = "ScheduleFollowUp"
        .Font.Size = 9
        .Font.Bold = True
    End With
    
    ' Add help text
    With ws.Range("B" & (btnRow + 2) & ":J" & (btnRow + 2))
        .Merge
        .Value = "Press Alt+1 through Alt+6 for quick response selection. Press Alt+S to save notes."
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

' Add keyboard handler
Private Sub AddKeyboardHandler(ws As Worksheet)
    ' Add code to worksheet module later
    ' (This will require adding code to the worksheet's code module)
End Sub

' Apply sheet protection
Private Sub ApplySheetProtection(ws As Worksheet)
    ' Allow users to select unlocked cells and use buttons
    ws.Protect password:="", UserInterfaceOnly:=True, _
        AllowFormattingCells:=False, _
        AllowFormattingColumns:=False, _
        AllowFormattingRows:=False, _
        AllowInsertingColumns:=False, _
        AllowInsertingRows:=False, _
        AllowInsertingHyperlinks:=False, _
        AllowDeletingColumns:=False, _
        AllowDeletingRows:=False, _
        AllowSorting:=False, _
        AllowFiltering:=False, _
        AllowUsingPivotTables:=False
    
    ' Unlock input cells
    ws.Range("CustomerName").Locked = False
    ws.Range("CustomerPhone").Locked = False
    ws.Range("CustomerEmail").Locked = False
    ws.Range("NotesArea").Locked = False
End Sub

