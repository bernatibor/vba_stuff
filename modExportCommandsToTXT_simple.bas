Attribute VB_Name = "modExportCommandsToTXT"
'@Author @tberna
'@version 1.0 build 1

Option Explicit

'
' Parse output of IONIX "Editor->Command" in the currently open email's body
' Output to new txt of device name
'
Sub exportCommandsToTxt()
    
    '
    ' Outlook Variables
    '
    Dim olMailItem As Outlook.MailItem
    Set olMailItem = Application.ActiveInspector.CurrentItem
    
    'On Error GoTo ErrorHandler

    Dim deviceName As String
    Dim deviceNameLine As Long
    Dim outputFolderPath As String
    Dim countryFolderPath As String
    Dim countryCode As String
    Dim currentFilePath As String
    Dim endOfSection As Long
    Dim deviceCount As Integer
    Dim fileAlreadyExistsCount As Integer
    Dim failedTaskCount As Integer
    
    '~~>stats
    fileAlreadyExistsCount = 0
    deviceCount = 0
    failedTaskCount = 0
    
    outputFolderPath = "C:\Users\611514880\Documents\Security\WORK\DHL\output"
    
    Dim lines() As String
    Dim lineParts() As String
    lines = Split(olMailItem.Body, vbNewLine)

    Dim i As Long
    Dim j As Long
    
    i = 0
    Do While i < UBound(lines)
        '~~>start of task section
        i = findNextLineNum("- Task", lines(), i)
        
        If i = UBound(lines) Then Exit Do    '~~>exit the loop if there are no more tasks in the mail
        deviceCount = deviceCount + 1
        
        '~~>get device name from line
        deviceNameLine = findNextLineNum("Device Name", lines(), i)
        lineParts = Split(lines(deviceNameLine), ": ")
        deviceName = lineParts(1)
        
        '~~>if the task failed, log and skip to next task
        If isTaskFailed(lines(deviceNameLine + 1)) Then
            failedTaskCount = failedTaskCount + 1         ' if we need to log device names that failed, this is where to begin
        Else
            '~~>check if country folder exists
            'countryCode = getCountryCode(deviceName)
            'countryFolderPath = outputFolderPath + "\" + countryCode           ' outputFolder / country folder / hostname.txt
            'Call createFolderIfNotExists(countryFolderPath)
    
            currentFilePath = outputFolderPath + "\" + deviceName + ".txt"
            If fileExists(currentFilePath) Then fileAlreadyExistsCount = fileAlreadyExistsCount + 1
            
            Open currentFilePath For Output As #1
            
            '~~>start of command output
            i = findNextLineNum("Enable Mode Results", lines(), i) + 1
            endOfSection = findNextLineNum("- Task", lines(), i) - 3
            
            For j = i To endOfSection
                Print #1, lines(j)
            Next j
            
            Close #1
            '~~>end of command output
            i = endOfSection
        End If
        i = i + 1
    Loop
    
ExitPoint:
    MsgBox "Number of devices processed in mail: " & deviceCount & vbNewLine & "Number of files overwritten: " & fileAlreadyExistsCount & vbNewLine & "Failed tasks: " & failedTaskCount
    Exit Sub

ErrorHandler:
    MsgBox Err.Number & " - " & Err.Description
    Resume ExitPoint

     '// Debug only
    Resume

End Sub

'
' Increment line index while the current line does not contain the searchString or until the array runs out
'
' @param String searchText the text we are searching for
' @param stringArray lines() the array containing the lines of the report
' @param Long i index of current line
'
' @return Long
'
Function findNextLineNum(searchText As String, lines() As String, ByVal i As Long) As Long
    Do While i < UBound(lines())
        If InStr(lines(i), searchText) > 0 Then
            Exit Do
        End If
        i = i + 1
    Loop
    findNextLineNum = i
End Function

' Check if a file exists
Function fileExists(fullPath As String) As Boolean

    Dim obj_fso As Object

    Set obj_fso = CreateObject("Scripting.FileSystemObject")
    fileExists = obj_fso.fileExists(fullPath)

End Function

Sub createFolderIfNotExists(folderPath As String)
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(folderPath) Then .CreateFolder folderPath
    End With
End Sub

Function getCountryCode(hostname As String) As String
    Dim temp, firstDash, secondDash As String
    
    firstDash = InStr(hostname, "-")
    secondDash = InStr(firstDash + 1, hostname, "-")
    temp = Right(hostname, Len(hostname) - secondDash)
    
    getCountryCode = Left(temp, 2)
End Function

Function isTaskFailed(line As String) As Boolean
    Dim lineParts() As String
    
    lineParts = Split(line, ": ")
    
    If lineParts(1) = "Completed" Then
        isTaskFailed = False
    Else
        isTaskFailed = True
    End If
End Function
