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
' @Procedure: RefreshSummary
' @Description: Refreshes the summary worksheet by calling the CreateSummarySheet procedure.
'---
Public Sub RefreshSummary()
    CreateSummarySheet
End Sub
