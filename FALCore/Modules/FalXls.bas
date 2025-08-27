Attribute VB_Name = "FalXls"
Option Explicit

'---
' @ModuleDescription: A module for Excel-specific functions, such as creating a summary worksheet.
'---

'---
' @Procedure: CreateSummarySheet
' @Description: Creates a worksheet that serves as a summary for all other worksheets in the current workbook.
'               It includes hyperlinks to facilitate navigation between worksheets and a button to refresh the summary.
'---
Public Sub CreateSummarySheet()
    Dim sh As Object ' Can be a Worksheet or a Chart
    Dim summarySheet As Worksheet
    Dim i As Long
    Dim btn As Button
    Const summarySheetName As String = "Sheet Summary"

    ' Turn off screen updating to speed up the process
    Application.ScreenUpdating = False

    ' Delete the summary sheet if it already exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(summarySheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Add a new worksheet as the first sheet
    Set summarySheet = Worksheets.Add(Before:=Worksheets(1))
    summarySheet.Name = summarySheetName

    ' Set column widths
    summarySheet.Columns("B:C").ColumnWidth = 30

    ' Add headers
    summarySheet.Cells(1, 2) = "Sheet Name"
    summarySheet.Cells(1, 3) = "Sheet Type"
    summarySheet.Cells(1, 4) = "Go to Sheet"

    ' Loop through all sheets (worksheets and charts) and create a hyperlink for each
    i = 2
    For Each sh In ThisWorkbook.Sheets
        If sh.Name <> summarySheetName Then
            summarySheet.Cells(i, 2) = sh.Name
            summarySheet.Cells(i, 3) = TypeName(sh)
            summarySheet.Hyperlinks.Add Anchor:=summarySheet.Cells(i, 4), _
                                        Address:="", _
                                        SubAddress:="'" & sh.Name & "'!A1", _
                                        TextToDisplay:="Link"
            i = i + 1
        End If
    Next sh

    ' Add a button to refresh the summary
    Set btn = summarySheet.Buttons.Add(summarySheet.Range("D" & i + 1).Left, _
                                      summarySheet.Range("D" & i + 1).Top, _
                                      100, _
                                      30)
    With btn
        .OnAction = "RefreshSummary"
        .Caption = "Refresh Summary"
        .Name = "RefreshButton"
    End With

    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub

Public Sub ExportVBAComponents()
    ' @brief Exports all standard and class modules from the active VBA project to subfolders.
    ' @details Creates 'vba/modules' and 'vba/classes' subfolders in the workbook's directory.
    '          Requires "Trust access to the VBA project object model" to be enabled.

    On Error GoTo ErrHandler

    ' --- Check for programmatic access ---
    If Application.VBE.ActiveVBProject Is Nothing Then
        MsgBox "Error: Programmatic access to the VBA project is not enabled." & vbCrLf & vbCrLf & _
               "Please go to File > Options > Trust Center > Trust Center Settings > Macro Settings, " & _
               "and check 'Trust access to the VBA project object model'.", vbCritical, "Access Denied"
        Exit Sub
    End If

    ' --- Define paths ---
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    If wbPath = "" Then
        MsgBox "Please save the workbook before exporting components.", vbExclamation, "Save Workbook"
        Exit Sub
    End If

    Dim basePath As String, modulesPath As String, classesPath As String
    basePath = wbPath & "\vba"
    modulesPath = basePath & "\modules"
    classesPath = basePath & "\classes"

    ' --- Create directories ---
    On Error Resume Next
    MkDir basePath
    MkDir modulesPath
    MkDir classesPath
    On Error GoTo ErrHandler

    ' --- Export components ---
    Dim vbComp As Object 'VBComponent
    Dim exportedCount As Long
    exportedCount = 0

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Dim exportPath As String
        Dim fileExt As String

        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule
                fileExt = ".bas"
                exportPath = modulesPath & "\" & vbComp.Name & fileExt
            Case 2 ' vbext_ct_ClassModule
                fileExt = ".cls"
                exportPath = classesPath & "\" & vbComp.Name & fileExt
            Case Else
                ' Skip other types like UserForms, ThisWorkbook, etc.
                fileExt = ""
        End Select

        If fileExt <> "" Then
            vbComp.Export exportPath
            exportedCount = exportedCount + 1
        End If
    Next vbComp

    MsgBox exportedCount & " component(s) exported successfully to:" & vbCrLf & basePath, vbInformation, "Export Complete"

    Exit Sub

ErrHandler:
    MsgBox "An unexpected error occurred during export: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub ImportVBAComponents()
    ' @brief Imports or updates all modules and classes from the default /vba subfolders.
    ' @details Scans /vba/modules and /vba/classes, removes existing components with the same name, and then imports the file.
    '          Requires "Trust access to the VBA project object model" to be enabled.

    On Error GoTo ErrHandler

    ' --- Check for programmatic access ---
    If Application.VBE.ActiveVBProject Is Nothing Then
        MsgBox "Error: Programmatic access to the VBA project is not enabled." & vbCrLf & vbCrLf & _
               "Please go to File > Options > Trust Center > Trust Center Settings > Macro Settings, " & _
               "and check 'Trust access to the VBA project object model'.", vbCritical, "Access Denied"
        Exit Sub
    End If

    ' --- Define paths ---
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    If wbPath = "" Then
        MsgBox "Please save the workbook first.", vbExclamation, "Save Workbook"
        Exit Sub
    End If

    Dim modulesPath As String, classesPath As String
    modulesPath = wbPath & "\vba\modules"
    classesPath = wbPath & "\vba\classes"

    ' --- Import components ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim importedCount As Long
    importedCount = 0

    ' Import modules
    If fso.FolderExists(modulesPath) Then
        Dim file As Object
        For Each file In fso.GetFolder(modulesPath).Files
            If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
                RemoveComponent file.Name
                ThisWorkbook.VBProject.VBComponents.Import file.Path
                importedCount = importedCount + 1
            End If
        Next file
    End If

    ' Import classes
    If fso.FolderExists(classesPath) Then
        For Each file In fso.GetFolder(classesPath).Files
            If LCase(fso.GetExtensionName(file.Name)) = "cls" Then
                RemoveComponent file.Name
                ThisWorkbook.VBProject.VBComponents.Import file.Path
                importedCount = importedCount + 1
            End If
        Next file
    End If

    MsgBox importedCount & " component(s) imported/updated successfully.", vbInformation, "Import Complete"

    Exit Sub

ErrHandler:
    MsgBox "An unexpected error occurred during import: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub UpdateVBAComponents()
    ' @brief Re-imports all VBA components from the default /vba subfolders, updating the project.
    ' @details This is an alias for the ImportVBAComponents sub.

    Dim response As VbMsgBoxResult
    response = MsgBox("This will overwrite any existing modules and classes in your project with the files from the /vba folder. Are you sure you want to continue?", _
                      vbQuestion + vbYesNo, "Confirm Update")

    If response = vbYes Then
        ImportVBAComponents
    End If
End Sub

Private Sub RemoveComponent(ByVal componentName As String)
    ' @brief Removes a component from the VBA project if it exists.
    ' @param componentName The name of the component to remove (including extension).

    Dim vbComp As Object 'VBComponent
    Dim baseName As String
    baseName = Left(componentName, InStrRev(componentName, ".") - 1)

    On Error Resume Next
    Set vbComp = ThisWorkbook.VBProject.VBComponents(baseName)
    If Not vbComp Is Nothing Then
        ThisWorkbook.VBProject.VBComponents.Remove vbComp
    End If
    On Error GoTo 0
End Sub

'---
' @Procedure: ImportXml
' @Description: Imports data from an XML file to a worksheet.
' @param FilePath The full path of the XML file to import.
' @param TargetRange The top-left cell of the destination range.
'---
Public Sub ImportXml(FilePath As String, TargetRange As Range)
    On Error GoTo ErrHandler

    ActiveWorkbook.XmlImport URL:=FilePath, ImportMap:=Nothing, Overwrite:=True, Destination:=TargetRange

    Exit Sub
ErrHandler:
    MsgBox "Failed to import XML file. " & Err.Description, vbCritical
End Sub

'---
' @Procedure: ExportXml
' @Description: Exports data from a worksheet to an XML file.
' @param SourceRange The range to export.
' @param FilePath The full path of the XML file to create.
'---
Public Sub ExportXml(SourceRange As Range, FilePath As String)
    On Error GoTo ErrHandler

    Dim xmlMap As XmlMap

    ' Add a new XML map to the workbook
    Set xmlMap = ThisWorkbook.XmlMaps.Add(SourceRange.Parent.Parent.Path & "\Schema.xsd", "Root")

    ' Export the data to the XML file
    SourceRange.Parent.Parent.XmlMaps(xmlMap.Name).Export URL:=FilePath

    Exit Sub
ErrHandler:
    MsgBox "Failed to export to XML. " & Err.Description, vbCritical
End Sub

'---
' @Procedure: ExportToPdf
' @Description: Exports a worksheet or a range to a PDF file.
' @param Target The worksheet or range to export.
' @param FilePath The full path of the PDF file to create.
'---
Public Sub ExportToPdf(Target As Object, FilePath As String)
    On Error GoTo ErrHandler

    If TypeOf Target Is Worksheet Then
        Target.ExportAsFixedFormat Type:=xlTypePDF, Filename:=FilePath
    ElseIf TypeOf Target Is Range Then
        Target.ExportAsFixedFormat Type:=xlTypePDF, Filename:=FilePath
    Else
        MsgBox "The target must be a worksheet or a range.", vbCritical
    End If

    Exit Sub
ErrHandler:
    MsgBox "Failed to export to PDF. " & Err.Description, vbCritical
End Sub

'---
' @Procedure: CreatePivotTable
' @Description: Creates a new PivotTable.
' @param SourceRange The data source for the PivotTable.
' @param DestinationRange The top-left cell of the PivotTable report.
' @param TableName The name of the new PivotTable.
'---
Public Sub CreatePivotTable(SourceRange As Range, DestinationRange As Range, TableName As String)
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable

    ' Create the PivotCache
    Set pvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=SourceRange)

    ' Create the PivotTable
    Set pvt = pvtCache.CreatePivotTable(TableDestination:=DestinationRange, TableName:=TableName)
End Sub

'---
' @Procedure: CreateConditionalFormat
' @Description: Creates a conditional formatting rule for a given range.
' @param TargetRange The range to apply the conditional formatting to.
' @param FormatConditionType The type of conditional formatting.
' @param Operator The operator for the conditional formatting.
' @param Formula1 The first formula for the conditional formatting.
' @param Formula2 The second formula for the conditional formatting (optional).
'---
Public Sub CreateConditionalFormat(TargetRange As Range, FormatConditionType As XlFormatConditionType, Operator As XlFormatConditionOperator, Formula1 As String, Optional Formula2 As String)
    With TargetRange.FormatConditions.Add(Type:=FormatConditionType, Operator:=Operator, Formula1:=Formula1, Formula2:=Formula2)
        ' Customize the formatting as needed
        .Interior.Color = RGB(255, 0, 0)
    End With
End Sub

'---
' @Procedure: CreateDataValidation
' @Description: Creates a data validation rule for a given range.
' @param TargetRange The range to apply the data validation to.
' @param ValidationType The type of data validation.
' @param Formula1 The first formula for the data validation.
' @param Formula2 The second formula for the data validation (optional).
'---
Public Sub CreateDataValidation(TargetRange As Range, ValidationType As XlDVType, Formula1 As String, Optional Formula2 As String)
    With TargetRange.Validation
        .Delete
        .Add Type:=ValidationType, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=Formula1, Formula2:=Formula2
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

'---
' @Procedure: RefreshSummary
' @Description: Refreshes the summary worksheet by calling the CreateSummarySheet procedure.
'---
Public Sub RefreshSummary()
    CreateSummarySheet
End Sub

'---
' @Procedure: CreateNamedRange
' @Description: Creates a new named range in the workbook.
' @param RangeName The name of the new named range.
' @param RefersTo The formula or range that the named range refers to.
'---
Public Sub CreateNamedRange(RangeName As String, RefersTo As String)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=RangeName, RefersTo:=RefersTo
    If Err.Number <> 0 Then
        MsgBox "Failed to create named range '" & RangeName & "'.", vbCritical
    End If
End Sub

'---
' @Procedure: DeleteNamedRange
' @Description: Deletes an existing named range from the workbook.
' @param RangeName The name of the named range to delete.
'---
Public Sub DeleteNamedRange(RangeName As String)
    On Error Resume Next
    ThisWorkbook.Names(RangeName).Delete
    If Err.Number <> 0 Then
        MsgBox "Failed to delete named range '" & RangeName & "'.", vbCritical
    End If
End Sub

'---
' @Procedure: ListNamedRanges
' @Description: Creates a new worksheet and lists all named ranges in the workbook.
'---
Public Sub ListNamedRanges()
    Dim summarySheet As Worksheet
    Dim i As Long
    Dim nm As Name

    ' Turn off screen updating to speed up the process
    Application.ScreenUpdating = False

    ' Delete the summary sheet if it already exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Named Ranges").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Add a new worksheet as the first sheet
    Set summarySheet = Worksheets.Add(Before:=Worksheets(1))
    summarySheet.Name = "Named Ranges"

    ' Set column widths
    summarySheet.Columns("B:C").ColumnWidth = 30

    ' Add headers
    summarySheet.Cells(1, 2) = "Named Range"
    summarySheet.Cells(1, 3) = "Refers To"

    ' Loop through all named ranges and list them
    i = 2
    For Each nm In ThisWorkbook.Names
        summarySheet.Cells(i, 2) = nm.Name
        summarySheet.Cells(i, 3) = nm.RefersTo
        i = i + 1
    Next nm

    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub
