Attribute VB_Name = "MSIntegration"
' ====================================================================
' MSIntegration Module - Version #1
' ====================================================================
' This module enhances integration with Microsoft services
' while maintaining Excel as the central hub

Option Explicit

' Constants for service types
Private Const SERVICE_OUTLOOK = "Outlook"
Private Const SERVICE_ONEDRIVE = "OneDrive"
Private Const SERVICE_SHAREPOINT = "SharePoint"
Private Const SERVICE_TEAMS = "Teams"

' Integration status tracking
Private integrationStatus As Object ' Dictionary

' Initialize integration
Public Sub InitializeIntegration()
    ' Create integration status dictionary
    Set integrationStatus = CreateObject("Scripting.Dictionary")
    
    ' Set default status
    integrationStatus(SERVICE_OUTLOOK) = CheckOutlookAvailability()
    integrationStatus(SERVICE_ONEDRIVE) = CheckOneDriveAvailability()
    integrationStatus(SERVICE_SHAREPOINT) = False ' Default to false until configured
    integrationStatus(SERVICE_TEAMS) = CheckTeamsAvailability()
    
    ' Load integration settings
    LoadIntegrationSettings
End Sub

' Check if a specific integration is available
Public Function IsIntegrationAvailable(serviceName As String) As Boolean
    ' Initialize if needed
    If integrationStatus Is Nothing Then
        InitializeIntegration
    End If
    
    ' Check if service exists in dictionary
    If integrationStatus.Exists(serviceName) Then
        IsIntegrationAvailable = integrationStatus(serviceName)
    Else
        IsIntegrationAvailable = False
    End If
End Function

' Check Outlook availability
Private Function CheckOutlookAvailability() As Boolean
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
        CheckOutlookAvailability = False
    Else
        ' Clean up
        Set objOutlook = Nothing
        CheckOutlookAvailability = True
    End If
    
    On Error GoTo 0
End Function

' Check OneDrive availability
Private Function CheckOneDriveAvailability() As Boolean
    On Error Resume Next
    
    ' Check for OneDrive folder
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Look for common OneDrive paths
    If fso.FolderExists(Environ("USERPROFILE") & "\OneDrive") Then
        CheckOneDriveAvailability = True
    ElseIf fso.FolderExists(Environ("USERPROFILE") & "\OneDrive - " & GetCompanyName()) Then
        CheckOneDriveAvailability = True
    Else
        CheckOneDriveAvailability = False
    End If
    
    On Error GoTo 0
End Function

' Get company name from user email
Private Function GetCompanyName() As String
    On Error Resume Next
    
    ' Try to get from Outlook
    If IsIntegrationAvailable(SERVICE_OUTLOOK) Then
        Dim objOutlook As Object
        Dim objNamespace As Object
        
        Set objOutlook = GetObject(, "Outlook.Application")
        Set objNamespace = objOutlook.GetNamespace("MAPI")
        
        ' Get current user email
        Dim userEmail As String
        userEmail = objNamespace.CurrentUser.Address
        
        ' Extract domain from email
        If InStr(userEmail, "@") > 0 Then
            GetCompanyName = Mid(userEmail, InStr(userEmail, "@") + 1)
            
            ' Remove domain extension
            If InStr(GetCompanyName, ".") > 0 Then
                GetCompanyName = left(GetCompanyName, InStr(GetCompanyName, ".") - 1)
            End If
        Else
            GetCompanyName = "Company"
        End If
        
        ' Clean up
        Set objNamespace = Nothing
        Set objOutlook = Nothing
    Else
        GetCompanyName = "Company"
    End If
    
    On Error GoTo 0
End Function

' Check Teams availability
Private Function CheckTeamsAvailability() As Boolean
    On Error Resume Next
    
    ' Check for Teams executable
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Look for common Teams paths
    If fso.FileExists(Environ("LOCALAPPDATA") & "\Microsoft\Teams\current\Teams.exe") Then
        CheckTeamsAvailability = True
    Else
        CheckTeamsAvailability = False
    End If
    
    On Error GoTo 0
End Function

' Load integration settings
Private Sub LoadIntegrationSettings()
    On Error Resume Next
    
    Dim settingsSheet As Worksheet
    
    ' Try to get settings sheet
    Set settingsSheet = ThisWorkbook.Sheets("Settings")
    
    If settingsSheet Is Nothing Then Exit Sub
    
    ' Load settings from sheet
    Dim settingRow As Range
    
    ' OneDrive/SharePoint settings
    Set settingRow = settingsSheet.Range("A:A").Find("SharePointURL", LookIn:=xlValues)
    If Not settingRow Is Nothing Then
        If settingRow.Offset(0, 1).Value <> "" Then
            integrationStatus(SERVICE_SHAREPOINT) = True
        End If
    End If
    
    ' Other settings as needed
    
    On Error GoTo 0
End Sub

' Create Outlook appointment for follow-up
Public Function CreateOutlookAppointment(customerName As String, customerPhone As String, _
    appointmentDate As Date, duration As Integer, subject As String, Optional notes As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Check if Outlook is available
    If Not IsIntegrationAvailable(SERVICE_OUTLOOK) Then
        CreateOutlookAppointment = False
        Exit Function
    End If
    
    ' Create Outlook objects
    Dim objOutlook As Object
    Dim objAppointment As Object
    
    Set objOutlook = GetObject(, "Outlook.Application")
    Set objAppointment = objOutlook.CreateItem(1) ' 1 = olAppointmentItem
    
    ' Set appointment properties
    With objAppointment
        .subject = subject
        .Location = "Phone: " & customerPhone
        .Start = appointmentDate
        .duration = duration
        .ReminderSet = True
        .ReminderMinutesBeforeStart = 15
        .body = "Follow-up call with " & customerName & vbCrLf & vbCrLf & _
                "Phone: " & customerPhone & vbCrLf & vbCrLf & _
                IIf(notes <> "", "Notes: " & notes, "")
                
        ' Add category if supported
        On Error Resume Next
        .Categories = "Novated Lease"
        On Error GoTo ErrorHandler
                
        .Save
    End With
    
    ' Clean up
    Set objAppointment = Nothing
    Set objOutlook = Nothing
    
    CreateOutlookAppointment = True
    Exit Function
    
ErrorHandler:
    CreateOutlookAppointment = False
End Function

' Send email to customer
Public Function SendCustomerEmail(customerName As String, customerEmail As String, _
    subject As String, body As String, Optional attachmentPath As String) As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Check if Outlook is available
    If Not IsIntegrationAvailable(SERVICE_OUTLOOK) Then
        SendCustomerEmail = False
        Exit Function
    End If
    
    ' Create Outlook objects
    Dim objOutlook As Object
    Dim objMail As Object
    
    Set objOutlook = GetObject(, "Outlook.Application")
    Set objMail = objOutlook.CreateItem(0) ' 0 = olMailItem
    
    ' Set email properties
    With objMail
        .to = customerEmail
        .subject = subject
        .HTMLBody = body
        
        ' Add attachment if specified
        If attachmentPath <> "" Then
            If Dir(attachmentPath) <> "" Then
                .Attachments.Add attachmentPath
            End If
        End If
        
        ' Display email for review before sending
        .Display
    End With
    
    ' Clean up
    Set objMail = Nothing
    Set objOutlook = Nothing
    
    SendCustomerEmail = True
    Exit Function
    
ErrorHandler:
    SendCustomerEmail = False
End Function

' Save document to OneDrive/SharePoint
Public Function SaveDocumentToCloud(localPath As String, customerName As String, _
    documentType As String) As String
    
    On Error GoTo ErrorHandler
    
    ' Check if OneDrive or SharePoint is available
    If Not (IsIntegrationAvailable(SERVICE_ONEDRIVE) Or IsIntegrationAvailable(SERVICE_SHAREPOINT)) Then
        SaveDocumentToCloud = ""
        Exit Function
    End If
    
    ' Determine destination path
    Dim destinationPath As String
    destinationPath = GetCloudDocumentPath(customerName, documentType)
    
    If destinationPath = "" Then
        SaveDocumentToCloud = ""
        Exit Function
    End If
    
    ' Create folder if it doesn't exist
    CreateFolderIfNeeded destinationPath
    
    ' Create destination file path
    Dim fileName As String
    fileName = GetSafeFileName(customerName & "_" & documentType & "_" & Format(Date, "yyyymmdd"))
    
    Dim destinationFile As String
    destinationFile = destinationPath & "\" & fileName & ".xlsx"
    
    ' Copy file
    FileCopy localPath, destinationFile
    
    ' Return cloud path
    SaveDocumentToCloud = destinationFile
    Exit Function
    
ErrorHandler:
    SaveDocumentToCloud = ""
End Function

' Get cloud document path
Private Function GetCloudDocumentPath(customerName As String, documentType As String) As String
    On Error Resume Next
    
    Dim settingsSheet As Worksheet
    Dim basePath As String
    
    ' Try to get settings sheet
    Set settingsSheet = ThisWorkbook.Sheets("Settings")
    
    If settingsSheet Is Nothing Then
        GetCloudDocumentPath = ""
        Exit Function
    End If
    
    ' Try to get base path from settings
    Dim settingRow As Range
    
    If IsIntegrationAvailable(SERVICE_SHAREPOINT) Then
        ' Use SharePoint path
        Set settingRow = settingsSheet.Range("A:A").Find("SharePointDocumentsPath", LookIn:=xlValues)
        If Not settingRow Is Nothing Then
            basePath = settingRow.Offset(0, 1).Value
        End If
    ElseIf IsIntegrationAvailable(SERVICE_ONEDRIVE) Then
        ' Use OneDrive path
        Set settingRow = settingsSheet.Range("A:A").Find("OneDriveDocumentsPath", LookIn:=xlValues)
        If Not settingRow Is Nothing Then
            basePath = settingRow.Offset(0, 1).Value
        End If
        
        ' If not specified, use default OneDrive path
        If basePath = "" Then
            basePath = Environ("USERPROFILE") & "\OneDrive\Documents\Novated Lease Quotes"
        End If
    End If
    
    ' Check if base path is specified
    If basePath = "" Then
        GetCloudDocumentPath = ""
        Exit Function
    End If
    
    ' Build customer-specific path
    GetCloudDocumentPath = basePath & "\" & GetSafeFileName(customerName)
    
    On Error GoTo 0
End Function

' Create folder if it doesn't exist
Private Sub CreateFolderIfNeeded(folderPath As String)
    On Error Resume Next
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create folder path recursively
    Dim parentPath As String
    Dim folders As Variant
    Dim currentPath As String
    Dim i As Integer
    
    ' Split path into folders
    folders = Split(folderPath, "\")
    
    ' Build path progressively and create folders as needed
    currentPath = folders(0) & "\"
    
    For i = 1 To UBound(folders)
        currentPath = currentPath & folders(i)
        
        If Not fso.FolderExists(currentPath) Then
            fso.CreateFolder currentPath
        End If
        
        currentPath = currentPath & "\"
    Next i
    
    On Error GoTo 0
End Sub

' Get safe filename (remove invalid characters)
Private Function GetSafeFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim result As String
    
    invalidChars = "/\:*?""<>|"
    result = fileName
    
    ' Replace invalid characters with underscore
    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i
    
    GetSafeFileName = result
End Function

' Generate Teams meeting link
Public Function GenerateTeamsMeetingLink(appointmentDate As Date, subject As String) As String
    On Error GoTo ErrorHandler
    
    ' Check if Outlook is available
    If Not IsIntegrationAvailable(SERVICE_OUTLOOK) Then
        GenerateTeamsMeetingLink = ""
        Exit Function
    End If
    
    ' Create Outlook objects
    Dim objOutlook As Object
    Dim objAppointment As Object
    Dim meetingLink As String
    
    Set objOutlook = GetObject(, "Outlook.Application")
    Set objAppointment = objOutlook.CreateItem(1) ' 1 = olAppointmentItem
    
    ' Set appointment properties
    With objAppointment
        .subject = subject
        .Start = appointmentDate
        .duration = 30
        
        ' Make it a Teams meeting
        On Error Resume Next
        .MeetingStatus = 1 ' 1 = olMeeting
        
        ' Use appropriate method based on version
        ' This will vary depending on your Office version
        Dim teamsAddin As Object
        For Each teamsAddin In objOutlook.COMAddIns
            If InStr(1, teamsAddin.description, "Teams", vbTextCompare) > 0 Then
                If teamsAddin.Connect Then
                    teamsAddin.Object.CreateTeamsMeeting objAppointment
                End If
                Exit For
            End If
        Next teamsAddin
        
        .Display ' Display to get the Teams link
        On Error GoTo ErrorHandler
        
        ' Try to extract Teams link from body (this is a simplified approach)
        meetingLink = ExtractTeamsLink(.body)
    End With
    
    ' Don't save this temporary appointment
    ' objAppointment.Close 2 ' 2 = olDiscard
    
    ' Clean up
    Set objAppointment = Nothing
    Set objOutlook = Nothing
    
    GenerateTeamsMeetingLink = meetingLink
    Exit Function
    
ErrorHandler:
    GenerateTeamsMeetingLink = ""
End Function

' Extract Teams meeting link from text
Private Function ExtractTeamsLink(text As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Look for common Teams meeting link pattern
    startPos = InStr(1, text, "https://teams.microsoft.com/l/meetup-join/", vbTextCompare)
    
    If startPos > 0 Then
        ' Find end of URL (space or line break)
        endPos = InStr(startPos, text, " ")
        If endPos = 0 Then
            endPos = InStr(startPos, text, vbCrLf)
        End If
        
        If endPos = 0 Then
            ' If no end marker found, take the rest of the string
            ExtractTeamsLink = Mid(text, startPos)
        Else
            ExtractTeamsLink = Mid(text, startPos, endPos - startPos)
        End If
    Else
        ExtractTeamsLink = ""
    End If
End Function

' Configure SharePoint integration
Public Function ConfigureSharePointIntegration(siteURL As String, documentLibrary As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim settingsSheet As Worksheet
    
    ' Try to get settings sheet
    On Error Resume Next
    Set settingsSheet = ThisWorkbook.Sheets("Settings")
    On Error GoTo ErrorHandler
    
    If settingsSheet Is Nothing Then
        ' Create settings sheet
        Set settingsSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        settingsSheet.Name = "Settings"
        
        ' Add headers
        settingsSheet.Range("A1").Value = "Setting"
        settingsSheet.Range("B1").Value = "Value"
        settingsSheet.Range("A1:B1").Font.Bold = True
    End If
    
    ' Find or create SharePoint settings
    Dim settingRow As Range
    Dim nextRow As Long
    
    ' SharePoint URL
    Set settingRow = settingsSheet.Range("A:A").Find("SharePointURL", LookIn:=xlValues)
    If settingRow Is Nothing Then
        nextRow = settingsSheet.Cells(settingsSheet.Rows.count, "A").End(xlUp).row + 1
        settingsSheet.Cells(nextRow, "A").Value = "SharePointURL"
        settingsSheet.Cells(nextRow, "B").Value = siteURL
    Else
        settingRow.Offset(0, 1).Value = siteURL
    End If
    
    ' Document library
    Set settingRow = settingsSheet.Range("A:A").Find("SharePointDocumentsPath", LookIn:=xlValues)
    If settingRow Is Nothing Then
        nextRow = settingsSheet.Cells(settingsSheet.Rows.count, "A").End(xlUp).row + 1
        settingsSheet.Cells(nextRow, "A").Value = "SharePointDocumentsPath"
        settingsSheet.Cells(nextRow, "B").Value = documentLibrary
    Else
        settingRow.Offset(0, 1).Value = documentLibrary
    End If
    
    ' Update integration status
    integrationStatus(SERVICE_SHAREPOINT) = True
    
    ConfigureSharePointIntegration = True
    Exit Function
    
ErrorHandler:
    ConfigureSharePointIntegration = False
End Function

