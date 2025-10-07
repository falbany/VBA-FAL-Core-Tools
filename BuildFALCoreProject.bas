Attribute VB_Name = "BuildFALCoreProject"

Option Explicit

Public Sub BuildProject()
    '@brief Imports all necessary source files (.bas, .cls) to build the FALCore project.
    '@details This macro must be run from a host workbook (e.g., a new .xlsm file).
    '         It requires "Trust access to the VBA project object model" to be enabled.

    On Error GoTo ErrHandler

    ' --- 1. Check for programmatic access ---
    If Application.VBE.ActiveVBProject Is Nothing Then
        MsgBox "Error: Programmatic access to the VBA project is not enabled." & vbCrLf & vbCrLf & _
               "Please go to File > Options > Trust Center > Trust Center Settings > Macro Settings, " & _
               "and check 'Trust access to the VBA project object model'.", vbCritical, "Access Denied"
        Exit Sub
    End If

    ' --- 2. Get the project root folder ---
    Dim projectRoot As String
    projectRoot = GetProjectRootFolder()
    If projectRoot = "" Then Exit Sub ' User cancelled

    ' --- 3. Define all files to be imported ---
    Dim filesToImport As Variant
    filesToImport = Array( _
        "FALCore/Modules/FalArray.bas", _
        "FALCore/Modules/FalCSV.bas", _
        "FALCore/Modules/FalFile.bas", _
        "FALCore/Modules/FalLang.bas", _
        "FALCore/Modules/FalLog.bas", _
        "FALCore/Modules/FalUtils.bas", _
        "FALCore/Modules/FalWork.bas", _
        "FALCore/Modules/FalXls.bas", _
        "FALCore/Modules/Plot.bas", _
        "FALCore/Classes/FalPlot.cls", _
        "FALCore/Classes/FalPlot2.cls", _
        "FALCore/Classes/clsMdmInputParameter.cls", _
        "FALCore/Classes/clsMdmOutputParameter.cls", _
        "FALCore/Classes/clsMdmDataBlock.cls", _
        "FALCore/Classes/clsMDM.cls", _
        "FALCore/demo/demoMDM.bas", _
        "modules/VBA-Dictionary/Dictionary.cls", _
        "modules/VBA-JSON/JsonConverter.bas", _
        "modules/VBA-StringBuilder/StringBuilder.cls", _
        "modules/VBA-Better-Array/BetterArray.cls", _
        "modules/VBA-Log/Logger.bas" _
    )

    ' --- 4. Import files ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim i As Long
    Dim importedCount As Long
    Dim failedCount As Long
    Dim filePath As String

    For i = LBound(filesToImport) To UBound(filesToImport)
        filePath = fso.BuildPath(projectRoot, filesToImport(i))

        If fso.FileExists(filePath) Then
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Import filePath
            If Err.Number = 0 Then
                importedCount = importedCount + 1
                Debug.Print "Successfully imported: " & filesToImport(i)
            Else
                failedCount = failedCount + 1
                Debug.Print "FAILED to import: " & filesToImport(i) & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            failedCount = failedCount + 1
            Debug.Print "File not found: " & filePath
        End If
    Next i

    ' --- 5. Final Report ---
    Dim report As String
    report = "Project Build Complete." & vbCrLf & vbCrLf & _
             "Successfully imported: " & importedCount & " files." & vbCrLf
    If failedCount > 0 Then
        report = report & "Failed to import: " & failedCount & " files." & vbCrLf & vbCrLf & _
                 "Please check the Immediate Window (Ctrl+G) for details."
    End If

    MsgBox report, IIf(failedCount > 0, vbExclamation, vbInformation), "Build Report"

    Exit Sub

ErrHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetProjectRootFolder() As String
    ' @brief Opens a dialog for the user to select the project's root folder.
    ' @return The path of the selected folder, or an empty string if cancelled.

    Dim fldr As Object ' FileDialog
    Set fldr = Application.FileDialog(4) ' msoFileDialogFolderPicker = 4

    With fldr
        .Title = "Select the Project Root Folder (the one containing FALCore and modules)"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            GetProjectRootFolder = ""
        Else
            GetProjectRootFolder = .SelectedItems(1)
        End If
    End With

End Function