Attribute VB_Name = "FalLog"
' **************************************************************************************
' Module    : FalLog
' Author    : Forent ALBANY
' Website   :
' Purpose   : Provides a flexible logging utility with configurable levels and destinations.
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2025-07-29              Initial Release
'---------------------------------------------------------------------------------------
' Dependencies: Microsoft Scripting Runtime (for FileSystemObject)
' **************************************************************************************
' @usage
' To use the logger, first initialize it at the start of your application.
' Then, call LogMessage throughout your code to record events.
'
' Sub Example_LoggerUsage()
'     ' --- Configuration ---
'     ' Configure the logger once at the start of your application.
'     ' This example logs everything up to the DEBUG level to both the
'     ' Immediate Window and a file.
'     FalLog.InitializeLogger Level:=llDebug, _
'                          Destination:=ldBoth, _
'                          FilePath:="C:\Temp\MyApp.log"
'
'     ' --- Example Log Calls ---
'     FalLog.LogMessage llInfo, "MainApp.Startup", "Application starting up."
'
'     Dim x As Long, y As Long
'     x = 10
'     y = 0
'
'     FalLog.LogMessage llDebug, "MainApp.Calculate", "Preparing to divide " & x & " by " & y
'
'     On Error Resume Next
'     Dim result As Double
'     result = x / y
'     If Err.Number <> 0 Then
'         ' Log the error with all details.
'         FalLog.LogMessage llError, "MainApp.Calculate", "Division failed. Error " & Err.Number & ": " & Err.Description
'         Err.Clear
'     End If
'     On Error GoTo 0
'
'     FalLog.LogMessage llWarning, "MainApp.Shutdown", "User has not saved their work."
'     FalLog.LogMessage llInfo, "MainApp.Shutdown", "Application closing."
' End Sub
' **************************************************************************************

Option Explicit

' Enum for defining the severity of a log message.
Public Enum LogLevel
    llOff = 0
    llError = 1
    llWarning = 2
    llInfo = 3
    llDebug = 4
End Enum

' Enum for defining where the log output should be sent.
Public Enum LogDestination
    ldNone = 0
    ldImmediate = 1
    ldFile = 2
    ldBoth = 3
End Enum

' --- Module-level configuration variables ---
Public CurrentLogLevel As LogLevel
Public CurrentLogDestination As LogDestination
Public LogFilePath As String

' Private helper to get the string representation of a log level.
Private Function GetLogLevelString(ByVal Level As LogLevel) As String
    Select Case Level
        Case llError: GetLogLevelString = "ERROR"
        Case llWarning: GetLogLevelString = "WARNING"
        Case llInfo: GetLogLevelString = "INFO"
        Case llDebug: GetLogLevelString = "DEBUG"
        Case Else: GetLogLevelString = "UNKNOWN"
    End Select
End Function

' Private helper to write a message to a file.
Private Sub WriteToFile(ByVal formattedMessage As String)
    If Trim(LogFilePath) = "" Then Exit Sub

    On Error GoTo ErrorHandler
    Dim fileNumber As Integer
    fileNumber = FreeFile
    
    ' Ensure the directory exists before writing
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim parentFolder As String
    parentFolder = fso.GetParentFolderName(LogFilePath)
    If Not fso.FolderExists(parentFolder) Then
        fso.CreateFolder parentFolder
    End If
    Set fso = Nothing

    Open LogFilePath For Append As #fileNumber
    Print #fileNumber, formattedMessage
    Close #fileNumber
    Exit Sub

ErrorHandler:
    Debug.Print "FalLog Error: Failed to write to log file '" & LogFilePath & "'. Error: " & Err.Description
End Sub

' The main logging procedure.
Public Sub LogMessage(ByVal Level As LogLevel, ByVal Source As String, ByVal Message As String)
    ' @brief Logs a message if its severity level is at or below the module's CurrentLogLevel.
    ' @param Level The severity level of the message (from the LogLevel enum).
    ' @param Source The source of the message (e.g., "ModuleName.FunctionName").
    ' @param Message The content of the log message.
    ' @details The output destination is controlled by the module-level CurrentLogDestination variable.
    '          The log file path is controlled by the module-level LogFilePath variable.
    
    ' 1. Check if logging is completely off or if the message level is too high to be logged.
    If CurrentLogLevel = llOff Or Level > CurrentLogLevel Then Exit Sub
    
    ' 2. Format the log message.
    Dim formattedMessage As String
    formattedMessage = Format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & _
                       Left(GetLogLevelString(Level) & Space(7), 7) & " | " & _
                       Source & " | " & _
                       Message
                       
    ' 3. Send the message to the configured destination(s).
    If CurrentLogDestination = ldImmediate Or CurrentLogDestination = ldBoth Then
        Debug.Print formattedMessage
    End If
    
    If CurrentLogDestination = ldFile Or CurrentLogDestination = ldBoth Then
        WriteToFile formattedMessage
    End If
End Sub

' An initialization routine to set up the logger easily.
Public Sub InitializeLogger(Optional ByVal Level As LogLevel = llInfo, Optional ByVal Destination As LogDestination = ldImmediate, Optional ByVal FilePath As String = "")
    ' @brief Configures the logger's settings.
    ' @param Level (Optional) The maximum log level to record. Defaults to llInfo.
    ' @param Destination (Optional) Where to send the log output. Defaults to the Immediate Window.
    ' @param FilePath (Optional) The full path for the log file. Required if Destination includes ldFile.
    
    CurrentLogLevel = Level
    CurrentLogDestination = Destination
    LogFilePath = FilePath
    
    If (Destination = ldFile Or Destination = ldBoth) And Trim(FilePath) = "" Then
        Debug.Print "FalLog Warning: Log destination includes a file, but LogFilePath is not set."
    End If
End Sub
