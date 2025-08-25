Attribute VB_Name = "FalPlot"
' Module: FalPlot
' Author: Florent ALBANY
' Date: 2025-08-25
' Version: 2.0
'
' Description:
' This module is a wrapper for the FalPlot class.
' It provides a procedural interface to the object-oriented FalPlot class.
'
' Dependencies:
' - FalPlot Class

Option Explicit

Public Sub Plot_SelectedRangeWithFormatting()
    Dim plot As New FalPlot
    Dim selectedRange As Range
    Dim chartTitle As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range of cells to plot.", vbExclamation
        Exit Sub
    End If
    Set selectedRange = Selection

    chartTitle = InputBox("Enter the chart title (leave empty for automatic):", "Chart Title")

    Set plot.Chart = selectedRange.Parent.Shapes.AddChart2(240, xlXYScatterLines).Chart
    plot.Chart.SetSourceData Source:=selectedRange
    plot.ChartTitle = chartTitle
    plot.FormatChart
End Sub

Public Sub Create_SmithChart()
    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart
    ' Not implemented yet
End Sub

Public Sub Format_Chart()
    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart
    plot.FormatChart
End Sub

Public Sub Create_YLog()
    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart
    ' Not implemented yet
End Sub

Public Sub Create_Derivative()
    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart
    ' Not implemented yet
End Sub

Public Sub Export_SelectedChartAsImage()
    Dim plot As New FalPlot
    Set plot.Chart = ActiveChart
    ' Not implemented yet
End Sub
