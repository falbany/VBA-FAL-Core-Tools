Attribute VB_Name = "Plot"
' Module: FalPlot
' Author: Florent ALBANY
' Date: 2025-08-25
' Version: 3.0
'
' Description:
' This module is a backward-compatible wrapper for the FalPlot class.
' It provides a procedural interface to the object-oriented FalPlot class,
' allowing old code that uses the original FalPlot module to continue working.
'
' Dependencies:
' - FalPlot Class

Option Explicit

Public Sub Plot_SelectedRangeWithFormatting()
    ' @brief Creates and formats a chart from the selected cell range, with UI prompts.
    On Error GoTo ErrHandler

    Dim selectedRange As Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to plot.", vbExclamation
        Exit Sub
    End If
    Set selectedRange = Selection

    Dim chartTitle As String
    chartTitle = InputBox("Enter the chart title (leave empty for automatic):", "Chart Title")

    Dim formattingOpts As String
    formattingOpts = InputBox("Enter any additional formatting options (e.g., YTitle=Units;SeriesMarkerStyle=2):", "Formatting Options")

    Dim plot As New FalPlot
    If plot.PlotFromRangeWithFormatting(selectedRange, chartTitle, , , formattingOpts) Then
        MsgBox "The chart '" & plot.Chart.Name & "' has been created and formatted successfully!", vbInformation, "Operation Successful"
    Else
        MsgBox "Failed to create the chart.", vbCritical, "Operation Failed"
    End If

    Exit Sub
ErrHandler:
    MsgBox "An unexpected error occurred in Plot_SelectedRangeWithFormatting.", vbCritical, "Error"
End Sub

Public Sub Create_SmithChart()
    ' @brief Creates a Smith chart based on the selected chart.
    On Error GoTo ErrHandler

    If ActiveChart Is Nothing Then
        MsgBox "Please select a chart first.", vbExclamation
        Exit Sub
    End If

    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart
    plot.CreateSmithChart

    Exit Sub
ErrHandler:
    MsgBox "An unexpected error occurred while creating the Smith Chart.", vbCritical, "Error"
End Sub

Public Sub Format_Chart()
    ' @brief Formats the selected chart with default options.
    On Error GoTo ErrHandler

    If ActiveChart Is Nothing Then
        MsgBox "Please select a chart first.", vbExclamation
        Exit Sub
    End If

    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart
    If Not plot.FormatChart() Then
         MsgBox "An error occurred while formatting the chart.", vbCritical, "Error"
    End If

    Exit Sub
ErrHandler:
    MsgBox "An unexpected error occurred in Format_Chart.", vbCritical, "Error"
End Sub

Public Sub Create_YLog()
    ' @brief Creates a copy of the selected chart with a logarithmic Y-axis.
    On Error GoTo ErrHandler

    If ActiveChart Is Nothing Then
        MsgBox "Please select a chart first.", vbExclamation
        Exit Sub
    End If

    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart
    plot.CreateYLog

    Exit Sub
ErrHandler:
    MsgBox "An unexpected error occurred while creating the Y-Log chart.", vbCritical, "Error"
End Sub

Public Sub Create_Derivative()
    ' @brief Creates a copy of the selected chart with its series' derivatives.
    On Error GoTo ErrHandler

    If ActiveChart Is Nothing Then
        MsgBox "Please select a chart first.", vbExclamation
        Exit Sub
    End If

    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart
    plot.CreateDerivative

    Exit Sub
ErrHandler:
    MsgBox "An unexpected error occurred while creating the derivative chart.", vbCritical, "Error"
End Sub

Public Sub Export_SelectedChartAsImage()
    ' @brief Exports the currently selected chart to an image file.
    On Error GoTo ErrHandler

    If ActiveChart Is Nothing Then
        MsgBox "Please select a chart first.", vbExclamation
        Exit Sub
    End If

    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart

    If plot.ExportAsImage() Then
        MsgBox "The chart was successfully exported!", vbInformation, "Export Successful"
    Else
        ' Error message is handled by the ExportAsImage method if user cancels.
    End If

    Exit Sub
ErrHandler:
    MsgBox "An unexpected error occurred in Export_SelectedChartAsImage.", vbCritical, "Error"
End Sub
