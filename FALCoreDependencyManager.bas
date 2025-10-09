Attribute VB_Name = "FALCoreDependencyManager"

Option Explicit

Private DEPENDENCIES As Variant

Private Sub InitializeDEPENDENCIES()
    DEPENDENCIES = Array( _
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
        "FALCore/Modules/FalMdmEnums.bas", _
        "modules/VBA-Dictionary/Dictionary.cls", _
        "modules/VBA-JSON/JsonConverter.bas", _
        "modules/VBA-StringBuilder/src/StringBuilder.cls", _
        "modules/VBA-Better-Array/src/BetterArray.cls", _
        "modules/VBA-Log/Logger.bas" _
    )
End Sub

Public Sub ImportDependencies()
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
    If projectRoot = "" Then
        MsgBox "Could not determine project root folder. Ensure the Excel file is saved in the project's root directory.", vbCritical
        Exit Sub
    End If
    
    ' --- 3. Define all files to be imported ---
    InitializeDEPENDENCIES

    ' --- 4. Import files ---
    Dim i As Long
    Dim importedCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim filePath As String
    Dim relativePath As String
    Dim componentName As String
    Dim sep As String
    sep = Application.PathSeparator
    Dim vbComp As Object ' VBComponent

    For i = LBound(DEPENDENCIES) To UBound(DEPENDENCIES)
        componentName = GetComponentNameFromFilePath(DEPENDENCIES(i))

        ' Check if component already exists
        Set vbComp = Nothing
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(componentName)
        On Error GoTo 0 ' Reset error handling

        If vbComp Is Nothing Then
            ' Component does not exist, try to import it
            relativePath = Replace(DEPENDENCIES(i), "/", sep)
            filePath = projectRoot & sep & relativePath

            If Dir(filePath) <> "" Then
                On Error Resume Next
                ThisWorkbook.VBProject.VBComponents.Import filePath
                If Err.Number = 0 Then
                    importedCount = importedCount + 1
                    Debug.Print "Successfully imported: " & DEPENDENCIES(i)
                Else
                    failedCount = failedCount + 1
                    Debug.Print "FAILED to import: " & DEPENDENCIES(i) & " | " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrHandler
            Else
                failedCount = failedCount + 1
                Debug.Print "File not found: " & filePath
            End If
        Else
            ' Component already exists, skip it
            skippedCount = skippedCount + 1
            Debug.Print "Component '" & componentName & "' already exists. Skipping."
        End If
    Next i

    ' --- 5. Final Report ---
    Dim report As String
    report = "Dependencies import complete." & vbCrLf & vbCrLf & _
             "Successfully imported: " & importedCount & " files." & vbCrLf & _
             "Skipped: " & skippedCount & " files (already exist)." & vbCrLf

    If failedCount > 0 Then
        report = report & "Failed to import: " & failedCount & " files." & vbCrLf & vbCrLf & _
                 "Please check the Immediate Window (Ctrl+G) for details."
    End If

    MsgBox report, IIf(failedCount > 0, vbExclamation, vbInformation), "Build Report"

    Exit Sub

ErrHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub ExportDependencies()
    ' @brief Exports all source files from the project to their respective locations.
    On Error GoTo ErrHandler

    Dim projectRoot As String
    projectRoot = GetProjectRootFolder()
    If projectRoot = "" Then Exit Sub

    InitializeDEPENDENCIES

    Dim i As Long
    Dim componentName As String
    Dim exportPath As String
    Dim sep As String
    sep = Application.PathSeparator
    Dim vbComp As Object ' VBComponent

    For i = LBound(DEPENDENCIES) To UBound(DEPENDENCIES)
        componentName = GetComponentNameFromFilePath(DEPENDENCIES(i))
        
        Set vbComp = Nothing
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(componentName)
        On Error GoTo 0

        If Not vbComp Is Nothing Then
            exportPath = projectRoot & sep & Replace(DEPENDENCIES(i), "/", sep)
            vbComp.Export exportPath
            Debug.Print "Successfully exported: " & componentName & " to " & exportPath
        Else
            Debug.Print "Component '" & componentName & "' not found in project. Skipping export."
        End If
    Next i
    
    MsgBox "Export complete. Check the immediate window for details.", vbInformation, "Export Report"
    Exit Sub

ErrHandler:
    MsgBox "An unexpected error occurred: " & err.Description, vbCritical, "Error"
End Sub

Public Sub RemoveDependencies()
    ' @brief Removes all imported source files from the project, ensuring only files declared in DEPENDENCIES are removed.
    On Error GoTo ErrHandler

    InitializeDEPENDENCIES

    Dim i As Long
    Dim componentName As String
    Dim vbComp As Object ' VBComponent

    For i = LBound(DEPENDENCIES) To UBound(DEPENDENCIES)
        componentName = GetComponentNameFromFilePath(DEPENDENCIES(i))
        
        Set vbComp = Nothing
        On Error Resume Next
        Set vbComp = ThisWorkbook.VBProject.VBComponents(componentName)
        On Error GoTo 0 ' Reset error handling

        If Not vbComp Is Nothing Then
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
            Debug.Print "Successfully removed: " & componentName
        Else
            Debug.Print "Component not found, skipping: " & componentName
        End If
    Next i

    MsgBox "Removal complete. Check the immediate window for details.", vbInformation, "Removal Report"
    Exit Sub

ErrHandler:
    MsgBox "An unexpected error occurred: " & err.Description, vbCritical, "Error"
End Sub

Public Sub RemoveAllDependencies()
    ' @brief Removes all components from the project except for this module.
    On Error GoTo ErrHandler

    Dim i As Long
    Dim count As Long
    count = ThisWorkbook.VBProject.VBComponents.count

    For i = count To 1 Step -1
        Dim vbComp As Object ' VBComponent
        Set vbComp = ThisWorkbook.VBProject.VBComponents(i)
        
        If vbComp.Name <> "FALCoreDependencyManager" Then
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
            If Err.Number = 0 Then
                Debug.Print "Successfully removed: " & vbComp.Name
            Else
                Debug.Print "FAILED to remove: " & vbComp.Name & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
    Next i

    MsgBox "All other dependencies removed.", vbInformation, "Removal Report"
    Exit Sub

ErrHandler:
    MsgBox "An unexpected error occurred: " & Err.Description, vbCritical, "Error"
End Sub

Private Function GetComponentNameFromFilePath(filePath As Variant) As String
    ' @brief Extracts the component name from a file path (e.g., "path/to/MyModule.bas" -> "MyModule").
    Dim arr As Variant
    arr = Split(filePath, "/")
    Dim fileName As String
    fileName = arr(UBound(arr))
    
    Dim dotIndex As Long
    dotIndex = InStrRev(fileName, ".")
    If dotIndex > 0 Then
        GetComponentNameFromFilePath = Left(fileName, dotIndex - 1)
    Else
        GetComponentNameFromFilePath = fileName
    End If
End Function

Private Function GetProjectRootFolder() As String
    ' @brief Returns the path of the workbook, assuming it's at the project root.
    GetProjectRootFolder = ThisWorkbook.Path
End Function

