VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FalPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************
' Class: FalPlot
' Author: Florent ALBANY
' Date: 2025-08-25
' Version: 1.0
'
' Description:
' This class provides a comprehensive set of methods and properties for creating,
' manipulating, and formatting advanced charts in Microsoft Excel. It is designed
' to simplify the process of plotting data from ranges or arrays, offering a high
' degree of customization for titles, axes, series, and visual styles.
'
' The class encapsulates a Chart object and provides an object-oriented interface
' to its functionalities.
'
' @Example
' Dim plot As New FalPlot
' Dim rng As Range
' Set rng = ActiveSheet.Range("A1:B10")
'
' Set plot.Chart = ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines).Chart
' plot.Chart.SetSourceData Source:=rng
'
' plot.ChartTitle = "My Chart"
' plot.X1Title = "X-Axis"
' plot.Y1Title = "Y-Axis"
' plot.FormatChart
'**************************************************************************************************

Option Explicit

Private m_Chart As Chart
' ... (all the other private members)

' ... (all the other properties with comments)

Private Sub Class_Initialize()
    ' Initialize the class with default values
    ' ... (all the initializations)
End Sub

' ... (all the methods with comments)
