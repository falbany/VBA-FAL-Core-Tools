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
    Dim ws As Worksheet
    Dim summarySheet As Worksheet
    Dim i As Long
    Dim btn As Button

    ' Turn off screen updating to speed up the process
    Application.ScreenUpdating = False

    ' Delete the summary sheet if it already exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Add a new worksheet as the first sheet
    Set summarySheet = Worksheets.Add(Before:=Worksheets(1))
    summarySheet.Name = "Summary"

    ' Set column widths
    summarySheet.Columns("B:B").ColumnWidth = 30

    ' Add headers
    summarySheet.Cells(1, 2) = "Worksheet Name"
    summarySheet.Cells(1, 3) = "Go to Sheet"

    ' Loop through all worksheets and create a hyperlink for each
    i = 2
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Summary" Then
            summarySheet.Cells(i, 2) = ws.Name
            summarySheet.Hyperlinks.Add Anchor:=summarySheet.Cells(i, 3), _
                                        Address:="", _
                                        SubAddress:="'" & ws.Name & "'!A1", _
                                        TextToDisplay:="Link"
            i = i + 1
        End If
    Next ws

    ' Add a button to refresh the summary
    Set btn = summarySheet.Buttons.Add(summarySheet.Range("C" & i + 1).Left, _
                                      summarySheet.Range("C" & i + 1).Top, _
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
' @Procedure: CreateChartSummarySheet
' @Description: Creates a worksheet that serves as a summary for all charts in the current workbook.
'               It includes hyperlinks to facilitate navigation between charts and a button to refresh the summary.
'---
Public Sub CreateChartSummarySheet()
    Dim cht As ChartObject
    Dim summarySheet As Worksheet
    Dim i As Long
    Dim btn As Button

    ' Turn off screen updating to speed up the process
    Application.ScreenUpdating = False

    ' Delete the summary sheet if it already exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Chart Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Add a new worksheet as the first sheet
    Set summarySheet = Worksheets.Add(Before:=Worksheets(1))
    summarySheet.Name = "Chart Summary"

    ' Set column widths
    summarySheet.Columns("B:B").ColumnWidth = 30

    ' Add headers
    summarySheet.Cells(1, 2) = "Chart Name"
    summarySheet.Cells(1, 3) = "Go to Chart"

    ' Loop through all chart objects and create a hyperlink for each
    i = 2
    For Each cht In ActiveSheet.ChartObjects
        summarySheet.Cells(i, 2) = cht.Name
        summarySheet.Hyperlinks.Add Anchor:=summarySheet.Cells(i, 3), _
                                    Address:="", _
                                    SubAddress:="'" & cht.Parent.Name & "'!" & cht.TopLeftCell.Address, _
                                    TextToDisplay:="Link"
        i = i + 1
    Next cht

    ' Add a button to refresh the summary
    Set btn = summarySheet.Buttons.Add(summarySheet.Range("C" & i + 1).Left, _
                                      summarySheet.Range("C" & i + 1).Top, _
                                      100, _
                                      30)
    With btn
        .OnAction = "RefreshChartSummary"
        .Caption = "Refresh Summary"
        .Name = "RefreshButton"
    End With

    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub

'---
' @Procedure: RefreshChartSummary
' @Description: Refreshes the chart summary worksheet by calling the CreateChartSummarySheet procedure.
'---
Public Sub RefreshChartSummary()
    CreateChartSummarySheet
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
