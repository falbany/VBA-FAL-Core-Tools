Attribute VB_Name = "VBAProjectImporter"

Option Explicit

' **************************************************************************************
' Module    : VBAProjectImporter
' Author    : Florent ALBANY
' Date      : 2025-08-25
' Purpose   : To simplify the process of importing multiple VBA modules and classes.
' **************************************************************************************

Public Sub ImportVBAFiles()
    ' @brief Prompts the user to select a folder and imports all .bas and .cls files from it.
    ' @details Requires the "Trust access to the VBA project object model" setting to be enabled in Excel's Trust Center.

    On Error GoTo ErrHandler

    ' --- Check for programmatic access to the VBA project ---
    If Application.VBE.ActiveVBProject Is Nothing Then
        MsgBox "Error: Programmatic access to the VBA project is not enabled." & vbCrLf & vbCrLf & _
               "Please go to File > Options > Trust Center > Trust Center Settings > Macro Settings, " & _
               "and check 'Trust access to the VBA project object model'.", vbCritical, "Access Denied"
        Exit Sub
    End If

    ' --- Get folder path from user ---
    Dim folderPath As String
    folderPath = GetFolder()
    If folderPath = "" Then Exit Sub ' User cancelled

    ' --- Import files ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim file As Object
    Dim importedCount As Long
    importedCount = 0

    For Each file In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(file.Name)) = "bas" Or LCase(fso.GetExtensionName(file.Name)) = "cls" Then
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import file.Path
            If Err.Number = 0 Then
                importedCount = importedCount + 1
            Else
                ' Optional: Log or display error for individual file import failure
                Debug.Print "Could not import file: " & file.Name & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
    Next file

    MsgBox importedCount & " file(s) imported successfully.", vbInformation, "Import Complete"

    Exit Sub

ErrHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetFolder() As String
    ' @brief Opens a dialog for the user to select a folder.
    ' @return The path of the selected folder, or an empty string if cancelled.

    Dim fldr As Object ' FileDialog
    Set fldr = Application.FileDialog(4) ' msoFileDialogFolderPicker = 4

    With fldr
        .Title = "Select a Folder to Import VBA Files From"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            GetFolder = ""
        Else
            GetFolder = .SelectedItems(1)
        End If
    End With

End Function
