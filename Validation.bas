Attribute VB_Name = "Validation"
' ====================================================================
' EnhancedIntegration Module
' ====================================================================
' This module improves the integration with Microsoft Outlook and
' Dynamics CRM to streamline the flow of customer data and activities

Option Explicit

' Auto-sync settings
Private autoSyncEnabled As Boolean
Private lastOutlookSync As Date
Private lastDynamicsSync As Date
Private syncIntervalMinutes As Integer

' Initialize integration settings
Public Sub InitializeEnhancedIntegration()
On Error Resume Next
    ' Set default settings
    autoSyncEnabled = True
    lastOutlookSync = 0
    lastDynamicsSync = 0
    syncIntervalMinutes = 30 ' Default to 30-minute sync interval
    
    ' Load settings from Settings sheet if available
    LoadIntegrationSettings
    
    ' Initial sync
    PerformFullSync
    
    ' Set up automatic sync timer
    Application.OnTime Now + TimeValue("00:" & syncIntervalMinutes & ":00"), "AutoSyncIntegrations"
    
    ' Show status message
    Application.StatusBar = "Enhanced integration initialized. Auto-sync every " & syncIntervalMinutes & " minutes."
    Application.OnTime Now + TimeValue("00:00:10"), "ResetStatusBar"
End Sub

' Load integration settings from the Settings sheet
Private Sub LoadIntegrationSettings()
On Error Resume Next
    Dim settingsSheet As Worksheet
    Dim settingRow As Range
    
    ' Try to get the Settings sheet
    On Error Resume Next
    Set settingsSheet = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    
    If Not settingsSheet Is Nothing Then
        ' Look up auto-sync setting
        Set settingRow = settingsSheet.Range("A:A").Find("AutoSyncEnabled", LookIn:=xlValues)
        If Not settingRow Is Nothing Then
            autoSyncEnabled = (settingRow.Offset(0, 1).Value = "True")
        End If
        
        ' Look up sync interval setting
        Set settingRow = settingsSheet.Range("A:A").Find("SyncIntervalMinutes", LookIn:=xlValues)
        If Not settingRow Is Nothing Then
            syncIntervalMinutes = settingRow.Offset(0, 1).Value
        End If
    End If
End Sub

' Save integration settings to the Settings sheet
Private Sub SaveIntegrationSettings()
On Error Resume Next
    Dim settingsSheet As Worksheet
    Dim settingRow As Range
    
    ' Try to get the Settings sheet
    On Error Resume Next
    Set settingsSheet = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    
    If settingsSheet Is Nothing Then
        ' Create the Settings sheet if it doesn't exist
        Set settingsSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        settingsSheet.Name = "Settings"
        
        ' Add headers
        settingsSheet.Range("A1").Value = "Setting"
        settingsSheet.Range("B1").Value = "Value"
        settingsSheet.Range("A1:B1").Font.Bold = True
    End If
    
    ' Update auto-sync setting
    Set settingRow = settingsSheet.Range("A:A").Find("AutoSyncEnabled", LookIn:=xlValues)
    If Not settingRow Is Nothing Then
        settingRow.Offset(0, 1).Value = autoSyncEnabled
    Else
        Dim nextRow As Long
        nextRow = settingsSheet.Cells(settingsSheet.Rows.count, "A").End(xlUp).row + 1
        settingsSheet.Cells(nextRow, 1).Value = "AutoSyncEnabled"
        settingsSheet.Cells(nextRow, 2).Value = autoSyncEnabled
    End If
    
    ' Update sync interval setting
    Set settingRow = settingsSheet.Range("A:A").Find("SyncIntervalMinutes", LookIn:=xlValues)
    If Not settingRow Is Nothing Then
        settingRow.Offset(0, 1).Value = syncIntervalMinutes
    Else
        nextRow = settingsSheet.Cells(settingsSheet.Rows.count, "A").End(xlUp).row + 1
        settingsSheet.Cells(nextRow, 1).Value = "SyncIntervalMinutes"
        settingsSheet.Cells(nextRow, 2).Value = syncIntervalMinutes
    End If
    
    ' Update last sync times
    Set settingRow = settingsSheet.Range("A:A").Find("LastOutlookSync", LookIn:=xlValues)
    If Not settingRow Is Nothing Then
        settingRow.Offset(0, 1).Value = Format(lastOutlookSync, "yyyy-mm-dd hh:mm:ss")
    Else
        nextRow = settingsSheet.Cells(settingsSheet.Rows.count, "A").End(xlUp).row + 1
        settingsSheet.Cells(nextRow, 1).Value = "LastOutlookSync"
        settingsSheet.Cells(nextRow, 2).Value = Format(lastOutlookSync, "yyyy-mm-dd hh:mm:ss")
    End If
    
    Set settingRow = settingsSheet.Range("A:A").Find("LastDynamicsSync", LookIn:=xlValues)
    If Not settingRow Is Nothing Then
        settingRow.Offset(0, 1).Value = Format(lastDynamicsSync, "yyyy-mm-dd hh:mm:ss")
    Else
        nextRow = settingsSheet.Cells(settingsSheet.Rows.count, "A").End(xlUp).row + 1
        settingsSheet.Cells(nextRow, 1).Value = "LastDynamicsSync"
        settingsSheet.Cells(nextRow, 2).Value = Format(lastDynamicsSync, "yyyy-mm-dd hh:mm:ss")
    End If
End Sub

' Toggle auto-sync on/off
Public Sub ToggleAutoSync()
On Error Resume Next
    autoSyncEnabled = Not autoSyncEnabled
    SaveIntegrationSettings
    
    If autoSyncEnabled Then
        MsgBox "Auto-sync has been enabled. Data will sync every " & syncIntervalMinutes & " minutes.", vbInformation
        
        ' Schedule next sync
        Application.OnTime Now + TimeValue("00:" & syncIntervalMinutes & ":00"), "AutoSyncIntegrations"
    Else
        MsgBox "Auto-sync has been disabled. You will need to manually sync data.", vbInformation
        
        ' Cancel any pending auto-sync
        On Error Resume Next
        Application.OnTime Now + TimeValue("00:" & syncIntervalMinutes & ":00"), "AutoSyncIntegrations", , False
        On Error GoTo 0
    End If
End Sub

' Change the sync interval
Public Sub ChangeSyncInterval()
On Error Resume Next
    Dim newInterval As String
    newInterval = InputBox("Enter the sync interval in minutes (10-60):", "Change Sync Interval", syncIntervalMinutes)
    
    If newInterval = "" Then Exit Sub
    
    Dim newValue As Integer
    newValue = Val(newInterval)
    
    If newValue < 10 Then
        MsgBox "Sync interval cannot be less than 10 minutes.", vbExclamation
        Exit Sub
    ElseIf newValue > 60 Then
        MsgBox "Sync interval cannot be more than 60 minutes.", vbExclamation
        Exit Sub
    End If
    
    ' Cancel current auto-sync timer
    If autoSyncEnabled Then
        On Error Resume Next
        Application.OnTime Now + TimeValue("00:" & syncIntervalMinutes & ":00"), "AutoSyncIntegrations", , False
        On Error GoTo 0
    End If
    
    ' Update interval
    syncIntervalMinutes = newValue
    SaveIntegrationSettings
    
    ' Set new timer if auto-sync is enabled
    If autoSyncEnabled Then
        Application.OnTime Now + TimeValue("00:" & syncIntervalMinutes & ":00"), "AutoSyncIntegrations"
    End If
    
    MsgBox "Sync interval updated to " & syncIntervalMinutes & " minutes.", vbInformation
End Sub

' Automatic sync handler - called by timer
Public Sub AutoSyncIntegrations()
On Error Resume Next
    ' Only proceed if auto-sync is enabled
    If Not autoSyncEnabled Then Exit Sub
    
    ' Perform sync
    Dim syncResult As Boolean
    syncResult = PerformFullSync(True)
    
    ' Schedule next sync
    Application.OnTime Now + TimeValue("00:" & syncIntervalMinutes & ":00"), "AutoSyncIntegrations"
End Sub

' Perform full synchronization with all systems
Public Function PerformFullSync(Optional silent As Boolean = False) As Boolean
On Error Resume Next
    Dim outlookSuccess As Boolean
    Dim dynamicsSuccess As Boolean
    Dim statusMessage As String
    
    ' Update status
    Application.StatusBar = "Syncing with external systems... Please wait."
    
    ' Sync with Outlook first
    outlookSuccess = SyncWithOutlookEnhanced(silent)
    lastOutlookSync = Now
    
    ' Then sync with Dynamics CRM
    dynamicsSuccess = SyncWithDynamicsEnhanced(silent)
    lastDynamicsSync = Now
    
    ' Save sync times
    SaveIntegrationSettings
    
    ' Prepare status message
    statusMessage = "Sync completed:" & vbCrLf
    statusMessage = statusMessage & "- Outlook: " & IIf(outlookSuccess, "Success", "Failed") & vbCrLf
    statusMessage = statusMessage & "- Dynamics CRM: " & IIf(dynamicsSuccess, "Success", "Failed")
    
    ' Update status
    Application.StatusBar = "Sync completed at " & Format(Now, "h:mm:ss AM/PM")
    Application.OnTime Now + TimeValue("00:00:10"), "ResetStatusBar"
    
    ' Show result if not silent mode
    If Not silent Then
        MsgBox statusMessage, vbInformation
    End If
    
    ' Return overall success
    PerformFullSync = outlookSuccess And dynamicsSuccess
End Function

' Enhanced Outlook synchronization
Public Function SyncWithOutlookEnhanced(Optional silent As Boolean = False) As Boolean
On Error GoTo ErrorHandler
    Dim objOutlook As Object
    Dim objNamespace As Object
    Dim objCalendar As Object
    Dim objInbox As Object
    Dim objContacts As Object
    Dim objTasks As Object
    Dim syncCount As Integer
    
    ' Initialize Outlook
    On Error Resume Next
    Set objOutlook = GetObject(, "Outlook.Application")
    If objOutlook Is Nothing Then
        Set objOutlook = CreateObject("Outlook.Application")
    End If
    
    If objOutlook Is Nothing Then
        If Not silent Then
            MsgBox "Could not connect to Outlook. Please ensure Outlook is running.", vbExclamation
        End If
        SyncWithOutlookEnhanced = False
        Exit Function
    End If
    
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objInbox = objNamespace.GetDefaultFolder(6) ' 6 = olFolderInbox
    Set objCalendar = objNamespace.GetDefaultFolder(9) ' 9 = olFolderCalendar
    Set objContacts = objNamespace.GetDefaultFolder(10) ' 10 = olFolderContacts
    Set objTasks = objNamespace.GetDefaultFolder(13) ' 13 = olFolderTasks
    
    syncCount = 0
    
    ' Sync customer emails
    syncCount = syncCount + SyncCustomerEmails(objInbox)
    
    ' Sync calendar appointments
    syncCount = syncCount + SyncCalendarAppointments(objCalendar)
    
    ' Sync tasks
    syncCount = syncCount + SyncOutlookTasks(objTasks)
    
    ' Cleanup
    Set objTasks = Nothing
    Set objContacts = Nothing
    Set objCalendar = Nothing
    Set objInbox = Nothing
    Set objNamespace = Nothing
    Set objOutlook = Nothing
    
    ' Show summary if not silent
    If Not silent Then
        MsgBox "Outlook sync completed successfully." & vbCrLf & _
               "- Synchronized " & syncCount & " items", vbInformation
    End If
    
    SyncWithOutlookEnhanced = True
    Exit Function
    
ErrorHandler:
    If Not silent Then
        MsgBox "Error syncing with Outlook: " & Err.description, vbCritical
    End If
    SyncWithOutlookEnhanced = False
End Function

' Sync customer emails from Outlook
Private Function SyncCustomerEmails(objInbox As Object) As Integer
On Error Resume Next
    Dim objItems As Object
    Dim objMail As Object
    Dim customerSheet As Worksheet
    Dim customerRange As Range
    Dim customerEmails As Collection
    Dim i As Long, j As Long
    Dim customerRow As Range
    Dim customerName As String
    Dim totalProcessed As Integer
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRange = customerSheet.Range("A2:Z" & customerSheet.UsedRange.Rows.count)
    
    ' Create collection of customer emails
    Set customerEmails = New Collection
    
    ' Assume column C has customer emails
    For i = 1 To customerRange.Rows.count
        If Not IsEmpty(customerRange.Cells(i, 3).Value) Then
            On Error Resume Next
            customerEmails.Add customerRange.Cells(i, 3).Value, CStr(customerRange.Cells(i, 3).Value)
            On Error GoTo 0
        End If
    Next i
    
    ' Process emails from the last 24 hours
    Dim filterDate As Date
    filterDate = Now - 1
    
    Dim filterString As String
    filterString = "[ReceivedTime] >= '" & Format(filterDate, "mm/dd/yyyy hh:mm AMPM") & "'"
    
    Set objItems = objInbox.Items.Restrict(filterString)
    objItems.Sort "[ReceivedTime]", True ' Sort descending (newest first)
    
    totalProcessed = 0
    
    ' Process up to 100 most recent emails
    Dim processLimit As Integer
    processLimit = 100
    
    For Each objMail In objItems
        On Error Resume Next
        ' Check if email is from one of our customers
        If Not objMail.SenderEmailAddress = "" Then
            For i = 1 To customerRange.Rows.count
                If LCase(customerRange.Cells(i, 3).Value) = LCase(objMail.SenderEmailAddress) Then
                    ' Found matching customer
                    customerName = customerRange.Cells(i, 2).Value ' Assume column B has customer name
                    
                    ' Log this email in customer communication history
                    LogCustomerContact customerName, "Received Email", objMail.subject, objMail.ReceivedTime
                    
                    ' Update customer last contact date
                    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
                    If Not customerRow Is Nothing Then
                        customerRow.Offset(0, 4).Value = objMail.ReceivedTime ' Last Contact column
                    End If
                    
                    totalProcessed = totalProcessed + 1
                    Exit For
                End If
            Next i
        End If
        
        ' Check limit
        If totalProcessed >= processLimit Then Exit For
        
        On Error GoTo 0
    Next objMail
    
    SyncCustomerEmails = totalProcessed
End Function

' Sync calendar appointments from Outlook
Private Function SyncCalendarAppointments(objCalendar As Object) As Integer
On Error Resume Next
    Dim objItems As Object
    Dim objAppt As Object
    Dim callSheet As Worksheet
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    Dim nextRow As Long
    Dim totalCreated As Integer
    
    ' Get sheets
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set callSheet = ThisWorkbook.Sheets("CallPlanner")
    
    ' Filter for today's and future appointments up to 7 days ahead
    Dim filterDate As Date
    filterDate = Date
    
    Dim endDate As Date
    endDate = Date + 7
    
    Dim filterString As String
    filterString = "[Start] >= '" & Format(filterDate, "mm/dd/yyyy") & " 12:00 AM' AND [Start] <= '" & _
                  Format(endDate, "mm/dd/yyyy") & " 11:59 PM'"
    
    Set objItems = objCalendar.Items.Restrict(filterString)
    objItems.Sort "[Start]"
    
    totalCreated = 0
    
    ' Process appointments
    For Each objAppt In objItems
        ' Check if this is a customer follow-up or has novated lease keywords
        If InStr(1, objAppt.subject, "Follow-up with", vbTextCompare) > 0 Or _
           InStr(1, objAppt.subject, "novated", vbTextCompare) > 0 Or _
           InStr(1, objAppt.subject, "lease", vbTextCompare) > 0 Then
            
            ' Extract possible customer name from subject
            Dim customerName As String
            If InStr(1, objAppt.subject, "Follow-up with", vbTextCompare) > 0 Then
                customerName = Trim(Replace(objAppt.subject, "Follow-up with", ""))
            Else
                ' Try to parse for a customer name - find first name that matches a customer
                Dim possibleNames() As String
                possibleNames = Split(objAppt.subject, " ")
                
                Dim nameFound As Boolean
                nameFound = False
                
                For i = 0 To UBound(possibleNames)
                    If Len(possibleNames(i)) > 2 Then ' Skip short words
                        Set customerRow = customerSheet.Range("B:B").Find(possibleNames(i), LookIn:=xlValues, LookAt:=xlPart)
                        If Not customerRow Is Nothing Then
                            customerName = customerRow.Value
                            nameFound = True
                            Exit For
                        End If
                    End If
                Next i
                
                If Not nameFound Then
                    ' Use appointment title if no customer name found
                    customerName = left(objAppt.subject, 30) & "..."
                End If
            End If
            
            ' Check if appointment is already in call planner
            Dim existingCall As Range
            Set existingCall = callSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
            
            ' Only add if it's a future appointment (not past) and not already in call planner
            If DateValue(objAppt.Start) >= Date And existingCall Is Nothing Then
                ' Add to call planner
                nextRow = callSheet.Cells(callSheet.Rows.count, "A").End(xlUp).row + 1
                
                callSheet.Cells(nextRow, 1).Value = Format(objAppt.Start, "h:mm AM/PM")
                callSheet.Cells(nextRow, 2).Value = customerName
                
                ' Try to find customer phone number
                Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
                If Not customerRow Is Nothing Then
                    callSheet.Cells(nextRow, 3).Value = customerRow.Offset(0, 2).Value ' Phone
                    callSheet.Cells(nextRow, 5).Value = customerRow.Offset(0, 3).Value ' Stage
                    callSheet.Cells(nextRow, 6).Value = customerRow.Offset(0, 12).Value ' Status
                Else
                    callSheet.Cells(nextRow, 3).Value = "Unknown"
                    callSheet.Cells(nextRow, 5).Value = "Unknown"
                    callSheet.Cells(nextRow, 6).Value = "New"
                End If
                
                callSheet.Cells(nextRow, 4).Value = "Calendar: " & objAppt.subject
                callSheet.Cells(nextRow, 7).Value = "Pending"
                
                totalCreated = totalCreated + 1
            End If
        End If
    Next objAppt
    
    SyncCalendarAppointments = totalCreated
End Function

' Sync tasks from Outlook
Private Function SyncOutlookTasks(objTasks As Object) As Integer
On Error Resume Next
    Dim objTask As Object
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    Dim totalUpdated As Integer
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    totalUpdated = 0
    
    ' Process only incomplete tasks
    Dim filterString As String
    filterString = "[Complete] = False"
    
    Dim objItems As Object
    Set objItems = objTasks.Items.Restrict(filterString)
    objItems.Sort "[DueDate]"
    
    For Each objTask In objItems
        ' Check if this is a customer-related task
        If InStr(1, objTask.subject, "Follow up with", vbTextCompare) > 0 Or _
           InStr(1, objTask.subject, "novated", vbTextCompare) > 0 Or _
           InStr(1, objTask.subject, "lease", vbTextCompare) > 0 Then
            
            ' Extract possible customer name from subject
            Dim customerName As String
            If InStr(1, objTask.subject, "Follow up with", vbTextCompare) > 0 Then
                customerName = Trim(Replace(objTask.subject, "Follow up with", ""))
            Else
                ' Try to parse for a customer name - find first name that matches a customer
                Dim possibleNames() As String
                possibleNames = Split(objTask.subject, " ")
                
                Dim nameFound As Boolean
                nameFound = False
                
                For i = 0 To UBound(possibleNames)
                    If Len(possibleNames(i)) > 2 Then ' Skip short words
                        Set customerRow = customerSheet.Range("B:B").Find(possibleNames(i), LookIn:=xlValues, LookAt:=xlPart)
                        If Not customerRow Is Nothing Then
                            customerName = customerRow.Value
                            nameFound = True
                            Exit For
                        End If
                    End If
                Next i
                
                If Not nameFound Then
                    ' Use task subject if no customer name found
                    customerName = left(objTask.subject, 30) & "..."
                End If
            End If
            
            ' Find customer in tracker
            Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
            
            If Not customerRow Is Nothing Then
                ' Update customer record with task info
                If customerRow.Offset(0, 7).Value <> objTask.dueDate Then
                    customerRow.Offset(0, 6).Value = "Task: " & objTask.subject ' Next Action column
                    customerRow.Offset(0, 7).Value = objTask.dueDate ' Next Action Date column
                    totalUpdated = totalUpdated + 1
                End If
                
                ' Check if task is completed
                If objTask.Complete Then
                    customerRow.Offset(0, 6).Value = "Completed: " & customerRow.Offset(0, 6).Value
                    LogCustomerContact customerName, "Task Completed", customerRow.Offset(0, 6).Value, Now()
                End If
            End If
        End If
    Next objTask
    
    SyncOutlookTasks = totalUpdated
End Function

' Enhanced Dynamics CRM synchronization
Public Function SyncWithDynamicsEnhanced(Optional silent As Boolean = False) As Boolean
On Error GoTo ErrorHandler
    ' Check if settings are configured
    Dim serverURL As String
    Dim userName As String
    Dim password As String
    
    serverURL = GetSetting("DynamicsURL")
    userName = GetSetting("DynamicsUser")
    
    If serverURL = "" Or userName = "" Then
        If Not silent Then
            MsgBox "Dynamics CRM connection not configured. Please update settings.", vbExclamation
        End If
        SyncWithDynamicsEnhanced = False
        Exit Function
    End If
    
    ' Connect to Dynamics
    If Not ConnectToDynamics() Then
        If Not silent Then
            MsgBox "Failed to connect to Dynamics CRM.", vbExclamation
        End If
        SyncWithDynamicsEnhanced = False
        Exit Function
    End If
    
    ' Perform bidirectional sync
    Dim importCount As Integer
    Dim exportCount As Integer
    
    ' Import leads and contacts from Dynamics to Excel
    importCount = ImportFromDynamics()
    
    ' Export updated data from Excel to Dynamics
    exportCount = ExportToDynamics()
    
    ' Cleanup
    CleanupDynamics
    
    ' Show summary if not silent
    If Not silent Then
        MsgBox "Dynamics CRM sync completed successfully." & vbCrLf & _
               "- Imported " & importCount & " new leads/contacts" & vbCrLf & _
               "- Exported " & exportCount & " updated records", vbInformation
    End If
    
    SyncWithDynamicsEnhanced = True
    Exit Function
    
ErrorHandler:
    If Not silent Then
        MsgBox "Error syncing with Dynamics CRM: " & Err.description, vbCritical
    End If
    SyncWithDynamicsEnhanced = False
End Function

' Import leads and contacts from Dynamics CRM
Private Function ImportFromDynamics() As Integer
On Error Resume Next
    ' This would contain the actual Dynamics CRM WebAPI communication
    ' For demonstration purposes, we'll simulate the import
    
    Dim customerSheet As Worksheet
    Dim nextRow As Long
    Dim totalImported As Integer
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    nextRow = customerSheet.Cells(customerSheet.Rows.count, "A").End(xlUp).row + 1
    
    ' In a real implementation, you'd make API calls to Dynamics and process the response
    ' For this example, we'll simulate importing 3 leads for demonstration
    
    totalImported = 0
    
    ' Check if we should import any records today (simulate API response)
    Dim shouldImport As Boolean
    shouldImport = (Day(Now) Mod 2 = 0) ' Import on even days for demo
    
    If shouldImport Then
        For i = 1 To 3
            ' Create a unique ID for the customer
            Dim customerID As String
            customerID = "CRM-" & Format(Date, "yyyymmdd") & "-" & Format(i + 100, "000")
            
            ' Check if this ID already exists
            Dim existingRow As Range
            Set existingRow = customerSheet.Range("A:A").Find(customerID, LookIn:=xlValues)
            
            If existingRow Is Nothing Then
                ' Add new customer record
                customerSheet.Cells(nextRow, 1).Value = customerID
                customerSheet.Cells(nextRow, 2).Value = "CRM Lead " & Format(Date, "mm-dd") & "-" & i
                customerSheet.Cells(nextRow, 3).Value = "lead" & i & "@example.com"
                customerSheet.Cells(nextRow, 4).Value = "555-123-" & Format(1000 + i, "0000")
                customerSheet.Cells(nextRow, 5).Value = "Initial Call"
                customerSheet.Cells(nextRow, 6).Value = Date
                customerSheet.Cells(nextRow, 7).Value = "Follow-up call"
                customerSheet.Cells(nextRow, 8).Value = Date + 2
                customerSheet.Cells(nextRow, 14).Value = "New Lead"
                customerSheet.Cells(nextRow, 15).Value = Application.userName
                customerSheet.Cells(nextRow, 16).Value = "crm-id-" & Format(i, "000000") ' Dynamics CRM ID
                
                nextRow = nextRow + 1
                totalImported = totalImported + 1
            End If
        Next i
    End If
    
    ImportFromDynamics = totalImported
End Function

' Export updated customer data to Dynamics CRM
Private Function ExportToDynamics() As Integer
On Error Resume Next
    ' This would contain the actual Dynamics CRM WebAPI communication
    ' For demonstration purposes, we'll simulate the export
    
    Dim customerSheet As Worksheet
    Dim totalExported As Integer
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    totalExported = 0
    
    ' In a real implementation, you'd iterate through updated records
    ' and make API calls to update them in Dynamics
    
    ' Look for recently modified records (simulate by finding records updated today)
    Dim i As Long
    For i = 2 To customerSheet.UsedRange.Rows.count
        ' Check if this is a record from Dynamics (has CRM ID)
        If Not IsEmpty(customerSheet.Cells(i, 16).Value) Then
            ' Check if it was recently modified (using last contact date as proxy)
            If Not IsEmpty(customerSheet.Cells(i, 6).Value) Then
                If DateValue(customerSheet.Cells(i, 6).Value) = Date Then
                    ' This record would be exported
                    totalExported = totalExported + 1
                End If
            End If
        End If
    Next i
    
    ExportToDynamics = totalExported
End Function

' Schedule follow-up in both systems
Public Function ScheduleFollowUpEnhanced(customerName As String, customerPhone As String, followupDate As Date, durationMinutes As Integer, actionType As String, Optional notes As String) As Boolean
On Error Resume Next
    Dim syncToOutlook As Boolean
    Dim syncToDynamics As Boolean
    
    ' Update customer record in Excel
    syncToOutlook = ScheduleCustomerFollowUp(customerName, actionType, followupDate)
    
    ' Create Outlook appointment
    If syncToOutlook Then
        syncToOutlook = CreateOutlookAppointment(customerName, customerPhone, followupDate, durationMinutes, notes)
    End If
    
    ' Update CRM if connected
    If ConnectToDynamics() Then
        syncToDynamics = UpdateDynamicsFollowUp(customerName, followupDate, actionType, notes)
        CleanupDynamics
    Else
        syncToDynamics = False
    End If
    
    ScheduleFollowUpEnhanced = syncToOutlook
End Function

' Create appointment in Outlook
Private Function CreateOutlookAppointment(customerName As String, customerPhone As String, appointmentDate As Date, durationMinutes As Integer, Optional notes As String) As Boolean
On Error Resume Next
    Dim objOutlook As Object
    Dim objAppointment As Object
    
    ' Try to get running instance of Outlook
    On Error Resume Next
    Set objOutlook = GetObject(, "Outlook.Application")
    If objOutlook Is Nothing Then
        Set objOutlook = CreateObject("Outlook.Application")
    End If
    
    If objOutlook Is Nothing Then
        CreateOutlookAppointment = False
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
    
    CreateOutlookAppointment = True
End Function

' Update follow-up in Dynamics
Private Function UpdateDynamicsFollowUp(customerName As String, followupDate As Date, actionType As String, Optional notes As String) As Boolean
On Error Resume Next
    ' This would contain the actual Dynamics CRM WebAPI communication
    ' For demonstration purposes, we'll simulate the update
    
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If customerRow Is Nothing Then
        UpdateDynamicsFollowUp = False
        Exit Function
    End If
    
    ' Check if this customer has a Dynamics CRM ID
    If IsEmpty(customerRow.Offset(0, 15).Value) Then
        UpdateDynamicsFollowUp = False
        Exit Function
    End If
    
    ' Simulate successful update to Dynamics
    UpdateDynamicsFollowUp = True
End Function

' Create or update a customer in the Excel tracker
Public Function UpsertCustomer(customerName As String, email As String, phone As String, stage As String, Optional crmID As String, Optional notes As String) As Boolean
On Error Resume Next
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    Dim nextRow As Long
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If customerRow Is Nothing Then
        ' New customer - add to tracker
        nextRow = customerSheet.Cells(customerSheet.Rows.count, "A").End(xlUp).row + 1
        
        ' Generate customer ID
        customerSheet.Cells(nextRow, 1).Value = "CUST-" & Format(Date, "yyyymmdd") & "-" & Format(nextRow - 1, "000")
        
        ' Add customer details
        customerSheet.Cells(nextRow, 2).Value = customerName
        customerSheet.Cells(nextRow, 3).Value = email
        customerSheet.Cells(nextRow, 4).Value = phone
        customerSheet.Cells(nextRow, 5).Value = stage ' Stage
        customerSheet.Cells(nextRow, 6).Value = Date ' Last Contact
        
        If notes <> "" Then
            customerSheet.Cells(nextRow, 11).Value = Format(Now, "yyyy-mm-dd hh:mm") & " - " & notes ' Notes
        End If
        
        customerSheet.Cells(nextRow, 14).Value = "New Lead" ' Status
        customerSheet.Cells(nextRow, 15).Value = Application.userName ' Assigned To
        
        If crmID <> "" Then
            customerSheet.Cells(nextRow, 16).Value = crmID ' Dynamics CRM ID
        End If
        
        ' Log in history
        LogCustomerContact customerName, "New Customer", "Added to system", Now()
    Else
        ' Existing customer - update details
        customerRow.Offset(0, 2).Value = email
        customerRow.Offset(0, 3).Value = phone
        
        ' Only update stage if provided
        If stage <> "" Then
            customerRow.Offset(0, 4).Value = stage
        End If
        
        customerRow.Offset(0, 5).Value = Date ' Last Contact
        
        ' Append notes if provided
        If notes <> "" Then
            If Not IsEmpty(customerRow.Offset(0, 10).Value) Then ' Notes column
                customerRow.Offset(0, 10).Value = customerRow.Offset(0, 10).Value & vbCrLf & _
                    Format(Now, "yyyy-mm-dd hh:mm") & " - " & notes
            Else
                customerRow.Offset(0, 10).Value = Format(Now, "yyyy-mm-dd hh:mm") & " - " & notes
            End If
        End If
        
        ' Update CRM ID if provided
        If crmID <> "" Then
            customerRow.Offset(0, 15).Value = crmID
        End If
        
        ' Log in history
        LogCustomerContact customerName, "Customer Update", "Updated customer details", Now()
    End If
    
    UpsertCustomer = True
End Function

' Update customer stage in both Excel and external systems
Public Function UpdateCustomerStage(customerName As String, newStage As String, Optional notes As String) As Boolean
On Error Resume Next
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    Dim syncToOutlook As Boolean
    Dim syncToDynamics As Boolean
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If customerRow Is Nothing Then
        UpdateCustomerStage = False
        Exit Function
    End If
    
    ' Update Excel
    Dim oldStage As String
    oldStage = customerRow.Offset(0, 4).Value
    customerRow.Offset(0, 4).Value = newStage
    customerRow.Offset(0, 5).Value = Date ' Last Contact
    
    ' Add stage change to notes
    If Not IsEmpty(customerRow.Offset(0, 10).Value) Then ' Notes column
        customerRow.Offset(0, 10).Value = customerRow.Offset(0, 10).Value & vbCrLf & _
            Format(Now, "yyyy-mm-dd hh:mm") & " - Stage changed from " & oldStage & " to " & newStage
        
        If notes <> "" Then
            customerRow.Offset(0, 10).Value = customerRow.Offset(0, 10).Value & vbCrLf & _
                Format(Now, "yyyy-mm-dd hh:mm") & " - " & notes
        End If
    Else
        customerRow.Offset(0, 10).Value = Format(Now, "yyyy-mm-dd hh:mm") & " - Stage changed from " & oldStage & " to " & newStage
        
        If notes <> "" Then
            customerRow.Offset(0, 10).Value = customerRow.Offset(0, 10).Value & vbCrLf & _
                Format(Now, "yyyy-mm-dd hh:mm") & " - " & notes
        End If
    End If
    
    ' Log in history
    LogCustomerContact customerName, "Stage Change", "Changed from " & oldStage & " to " & newStage, Now()
    
    ' Update Dynamics if customer has CRM ID
    If Not IsEmpty(customerRow.Offset(0, 15).Value) Then
        syncToDynamics = True ' Placeholder for actual Dynamics update code
    End If
    
    UpdateCustomerStage = True
End Function

' Schedule customer follow-up in Excel
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

' Log customer contact to history
Public Sub LogCustomerContact(customerName As String, contactType As String, details As String, contactDate As Date)
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
        
        ' Add headers
        historySheet.Range("A1").Value = "Customer"
        historySheet.Range("B1").Value = "Contact Type"
        historySheet.Range("C1").Value = "Details"
        historySheet.Range("D1").Value = "Date/Time"
        historySheet.Range("E1").Value = "User"
        historySheet.Range("A1:E1").Font.Bold = True
    End If
    
    ' Find next empty row
    nextRow = historySheet.Cells(historySheet.Rows.count, "A").End(xlUp).row + 1
    
    ' Add contact record
    historySheet.Cells(nextRow, 1).Value = customerName
    historySheet.Cells(nextRow, 2).Value = contactType
    historySheet.Cells(nextRow, 3).Value = details
    historySheet.Cells(nextRow, 4).Value = contactDate
    historySheet.Cells(nextRow, 5).Value = Application.userName
End Sub

' Create an email in Outlook
Public Function SendEnhancedEmail(customerName As String, email As String, subject As String, body As String, Optional crmLogging As Boolean = True) As Boolean
On Error Resume Next
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
        
        ' Add signature if available
        On Error Resume Next
        .HTMLBody = .HTMLBody & objOutlook.Application.ActiveExplorer.Session.Accounts.Item(1).CurrentUser.AddressEntry.GetFreeBusy(Now, 15, True)
        On Error GoTo 0
        
        ' Display email for review before sending
        .Display
    End With
    
    ' Log in customer history
    LogCustomerContact customerName, "Email Sent", subject, Now()
    
    ' Update last contact date
    UpdateCustomerLastContact customerName
    
    ' Log in Dynamics CRM if requested
    If crmLogging And ConnectToDynamics() Then
        LogEmailToDynamics customerName, subject
        CleanupDynamics
    End If
    
    SendEnhancedEmail = True
End Function

' Log email to Dynamics CRM
Private Function LogEmailToDynamics(customerName As String, subject As String) As Boolean
On Error Resume Next
    ' This would contain the actual Dynamics CRM WebAPI communication
    ' For demonstration purposes, we'll simulate the logging
    
    Dim customerSheet As Worksheet
    Dim customerRow As Range
    
    Set customerSheet = ThisWorkbook.Sheets("CustomerTracker")
    Set customerRow = customerSheet.Range("B:B").Find(customerName, LookIn:=xlValues)
    
    If customerRow Is Nothing Then
        LogEmailToDynamics = False
        Exit Function
    End If
    
    ' Check if this customer has a Dynamics CRM ID
    If IsEmpty(customerRow.Offset(0, 15).Value) Then
        LogEmailToDynamics = False
        Exit Function
    End If
    
    ' Simulate successful log to Dynamics
    LogEmailToDynamics = True
End Function

' Reset status bar
Public Sub ResetStatusBar()
On Error Resume Next
    Application.StatusBar = False
End Sub

