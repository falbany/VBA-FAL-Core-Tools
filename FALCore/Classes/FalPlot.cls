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
Private m_ChartType As XlChartType
Private m_ChartTitle As String
Private m_Y1Title As String
Private m_Y2Title As String
Private m_X1Title As String
Private m_X2Title As String
Private m_X1ScaleType As XlScaleType
Private m_Y1ScaleType As XlScaleType
Private m_X2ScaleType As XlScaleType
Private m_Y2ScaleType As XlScaleType
Private m_HasLegend As Boolean
Private m_HasTitle As Boolean
Private m_X1HasTitle As Boolean
Private m_Y1HasTitle As Boolean
Private m_X2HasTitle As Boolean
Private m_Y2HasTitle As Boolean
Private m_X1ShowAxis As Boolean
Private m_Y1ShowAxis As Boolean
Private m_X1ShowGridLines As Boolean
Private m_Y1ShowGridLines As Boolean
Private m_FontName As String
Private m_TitleFontSize As Integer
Private m_XAxisTitleFontSize As Integer
Private m_YAxisTitleFontSize As Integer
Private m_XAxisTicksFontSize As Integer
Private m_YAxisTicksFontSize As Integer
Private m_LegendFontSize As Integer
Private m_LegendPosition As XlLegendPosition
Private m_LegendInLayout As Boolean
Private m_PlotAreaLineWeight As Double
Private m_SeriesLineWeight As Double
Private m_SeriesLineDashStyle As MsoLineDashStyle
Private m_SeriesMarkerStyle As XlMarkerStyle
Private m_HasDataLabels As Boolean
Private m_X1LabelNumberFormat As String
Private m_Y1LabelNumberFormat As String
Private m_X2LabelNumberFormat As String
Private m_Y2LabelNumberFormat As String
Private m_Y1Min As Variant
Private m_Y1Max As Variant
Private m_X1Min As Variant
Private m_X1Max As Variant
Private m_Y2Min As Variant
Private m_Y2Max As Variant
Private m_X2Min As Variant
Private m_X2Max As Variant
Private m_SquarePlot As Boolean
Private m_CrossesAt As Double
Private m_PlotBy As XlRowCol

Public Property Get Chart() As Chart
    Set Chart = m_Chart
End Property

Public Property Set Chart(ByVal value As Chart)
    Set m_Chart = value
End Property

Public Property Get ChartType() As XlChartType
    ChartType = m_ChartType
End Property

Public Property Let ChartType(ByVal value As XlChartType)
    m_ChartType = value
    If Not m_Chart Is Nothing Then m_Chart.ChartType = value
End Property

Public Property Get ChartTitle() As String
    ChartTitle = m_ChartTitle
End Property

Public Property Let ChartTitle(ByVal value As String)
    m_ChartTitle = value
    If Not m_Chart Is Nothing Then
        If m_Chart.HasTitle Then
            m_Chart.chartTitle.Text = value
        End If
    End If
End Property

Public Property Get Y1Title() As String
    Y1Title = m_Y1Title
End Property

Public Property Let Y1Title(ByVal value As String)
    m_Y1Title = value
    If Not m_Chart Is Nothing Then
        If m_Chart.Axes(xlValue, xlPrimary).HasTitle Then
            m_Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = value
        End If
    End If
End Property

Public Property Get Y2Title() As String
    Y2Title = m_Y2Title
End Property

Public Property Let Y2Title(ByVal value As String)
    m_Y2Title = value
    If Not m_Chart Is Nothing Then
        If m_Chart.Axes(xlValue, xlSecondary).HasTitle Then
            m_Chart.Axes(xlValue, xlSecondary).AxisTitle.Text = value
        End If
    End If
End Property

Public Property Get X1Title() As String
    X1Title = m_X1Title
End Property

Public Property Let X1Title(ByVal value As String)
    m_X1Title = value
    If Not m_Chart Is Nothing Then
        If m_Chart.Axes(xlCategory, xlPrimary).HasTitle Then
            m_Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = value
        End If
    End If
End Property

Public Property Get X2Title() As String
    X2Title = m_X2Title
End Property

Public Property Let X2Title(ByVal value As String)
    m_X2Title = value
    If Not m_Chart Is Nothing Then
        If m_Chart.Axes(xlCategory, xlSecondary).HasTitle Then
            m_Chart.Axes(xlCategory, xlSecondary).AxisTitle.Text = value
        End If
    End If
End Property

Public Property Get X1ScaleType() As XlScaleType
    X1ScaleType = m_X1ScaleType
End Property

Public Property Let X1ScaleType(ByVal value As XlScaleType)
    m_X1ScaleType = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).ScaleType = value
End Property

Public Property Get Y1ScaleType() As XlScaleType
    Y1ScaleType = m_Y1ScaleType
End Property

Public Property Let Y1ScaleType(ByVal value As XlScaleType)
    m_Y1ScaleType = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).ScaleType = value
End Property

Public Property Get X2ScaleType() As XlScaleType
    X2ScaleType = m_X2ScaleType
End Property

Public Property Let X2ScaleType(ByVal value As XlScaleType)
    m_X2ScaleType = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlSecondary).ScaleType = value
End Property

Public Property Get Y2ScaleType() As XlScaleType
    Y2ScaleType = m_Y2ScaleType
End Property

Public Property Let Y2ScaleType(ByVal value As XlScaleType)
    m_Y2ScaleType = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlSecondary).ScaleType = value
End Property

Public Property Get HasLegend() As Boolean
    HasLegend = m_HasLegend
End Property

Public Property Let HasLegend(ByVal value As Boolean)
    m_HasLegend = value
    If Not m_Chart Is Nothing Then m_Chart.HasLegend = value
End Property

Public Property Get HasTitle() As Boolean
    HasTitle = m_HasTitle
End Property

Public Property Let HasTitle(ByVal value As Boolean)
    m_HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.HasTitle = value
End Property

Public Property Get X1HasTitle() As Boolean
    X1HasTitle = m_X1HasTitle
End Property

Public Property Let X1HasTitle(ByVal value As Boolean)
    m_X1HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).HasTitle = value
End Property

Public Property Get Y1HasTitle() As Boolean
    Y1HasTitle = m_Y1HasTitle
End Property

Public Property Let Y1HasTitle(ByVal value As Boolean)
    m_Y1HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).HasTitle = value
End Property

Public Property Get X2HasTitle() As Boolean
    X2HasTitle = m_X2HasTitle
End Property

Public Property Let X2HasTitle(ByVal value As Boolean)
    m_X2HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlSecondary).HasTitle = value
End Property

Public Property Get Y2HasTitle() As Boolean
    Y2HasTitle = m_Y2HasTitle
End Property

Public Property Let Y2HasTitle(ByVal value As Boolean)
    m_Y2HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlSecondary).HasTitle = value
End Property

Public Property Get X1ShowAxis() As Boolean
    X1ShowAxis = m_X1ShowAxis
End Property

Public Property Let X1ShowAxis(ByVal value As Boolean)
    m_X1ShowAxis = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).Format.Line.Visible = IIf(value, msoTrue, msoFalse)
End Property

Public Property Get Y1ShowAxis() As Boolean
    Y1ShowAxis = m_Y1ShowAxis
End Property

Public Property Let Y1ShowAxis(ByVal value As Boolean)
    m_Y1ShowAxis = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).Format.Line.Visible = IIf(value, msoTrue, msoFalse)
End Property

Public Property Get X1ShowGridLines() As Boolean
    X1ShowGridLines = m_X1ShowGridLines
End Property

Public Property Let X1ShowGridLines(ByVal value As Boolean)
    m_X1ShowGridLines = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).HasMajorGridlines = value
End Property

Public Property Get Y1ShowGridLines() As Boolean
    Y1ShowGridLines = m_Y1ShowGridLines
End Property

Public Property Let Y1ShowGridLines(ByVal value As Boolean)
    m_Y1ShowGridLines = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).HasMajorGridlines = value
End Property

Public Property Get FontName() As String
    FontName = m_FontName
End Property

Public Property Let FontName(ByVal value As String)
    m_FontName = value
    If Not m_Chart Is Nothing Then
        ' Apply to all text elements
        If m_Chart.HasTitle Then m_Chart.chartTitle.Font.Name = value
        If m_Chart.Axes(xlCategory, xlPrimary).HasTitle Then m_Chart.Axes(xlCategory, xlPrimary).AxisTitle.Font.Name = value
        If m_Chart.Axes(xlValue, xlPrimary).HasTitle Then m_Chart.Axes(xlValue, xlPrimary).AxisTitle.Font.Name = value
        m_Chart.Axes(xlCategory, xlPrimary).TickLabels.Font.Name = value
        m_Chart.Axes(xlValue, xlPrimary).TickLabels.Font.Name = value
        If m_Chart.HasLegend Then m_Chart.Legend.Font.Name = value
    End If
End Property

Public Property Get TitleFontSize() As Integer
    TitleFontSize = m_TitleFontSize
End Property

Public Property Let TitleFontSize(ByVal value As Integer)
    m_TitleFontSize = value
    If Not m_Chart Is Nothing Then If m_Chart.HasTitle Then m_Chart.chartTitle.Font.Size = value
End Property

Public Property Get XAxisTitleFontSize() As Integer
    XAxisTitleFontSize = m_XAxisTitleFontSize
End Property

Public Property Let XAxisTitleFontSize(ByVal value As Integer)
    m_XAxisTitleFontSize = value
    If Not m_Chart Is Nothing Then If m_Chart.Axes(xlCategory, xlPrimary).HasTitle Then m_Chart.Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = value
End Property

Public Property Get YAxisTitleFontSize() As Integer
    YAxisTitleFontSize = m_YAxisTitleFontSize
End Property

Public Property Let YAxisTitleFontSize(ByVal value As Integer)
    m_YAxisTitleFontSize = value
    If Not m_Chart Is Nothing Then If m_Chart.Axes(xlValue, xlPrimary).HasTitle Then m_Chart.Axes(xlValue, xlPrimary).AxisTitle.Font.Size = value
End Property

Public Property Get XAxisTicksFontSize() As Integer
    XAxisTicksFontSize = m_XAxisTicksFontSize
End Property

Public Property Let XAxisTicksFontSize(ByVal value As Integer)
    m_XAxisTicksFontSize = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).TickLabels.Font.Size = value
End Property

Public Property Get YAxisTicksFontSize() As Integer
    YAxisTicksFontSize = m_YAxisTicksFontSize
End Property

Public Property Let YAxisTicksFontSize(ByVal value As Integer)
    m_YAxisTicksFontSize = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).TickLabels.Font.Size = value
End Property

Public Property Get LegendFontSize() As Integer
    LegendFontSize = m_LegendFontSize
End Property

Public Property Let LegendFontSize(ByVal value As Integer)
    m_LegendFontSize = value
    If Not m_Chart Is Nothing Then If m_Chart.HasLegend Then m_Chart.Legend.Font.Size = value
End Property

Public Property Get LegendPosition() As XlLegendPosition
    LegendPosition = m_LegendPosition
End Property

Public Property Let LegendPosition(ByVal value As XlLegendPosition)
    m_LegendPosition = value
    If Not m_Chart Is Nothing Then If m_Chart.HasLegend Then m_Chart.Legend.Position = value
End Property

Public Property Get LegendInLayout() As Boolean
    LegendInLayout = m_LegendInLayout
End Property

Public Property Let LegendInLayout(ByVal value As Boolean)
    m_LegendInLayout = value
    If Not m_Chart Is Nothing Then If m_Chart.HasLegend Then m_Chart.Legend.IncludeInLayout = value
End Property

Public Property Get PlotAreaLineWeight() As Double
    PlotAreaLineWeight = m_PlotAreaLineWeight
End Property

Public Property Let PlotAreaLineWeight(ByVal value As Double)
    m_PlotAreaLineWeight = value
    If Not m_Chart Is Nothing Then m_Chart.PlotArea.Format.Line.Weight = value
End Property

Public Property Get SeriesLineWeight() As Double
    SeriesLineWeight = m_SeriesLineWeight
End Property

Public Property Let SeriesLineWeight(ByVal value As Double)
    m_SeriesLineWeight = value
    If Not m_Chart Is Nothing Then
        Dim ser As series
        For Each ser In m_Chart.SeriesCollection
            ser.Format.Line.Weight = value
        Next ser
    End If
End Property

Public Property Get SeriesLineDashStyle() As MsoLineDashStyle
    SeriesLineDashStyle = m_SeriesLineDashStyle
End Property

Public Property Let SeriesLineDashStyle(ByVal value As MsoLineDashStyle)
    m_SeriesLineDashStyle = value
    If Not m_Chart Is Nothing Then
        Dim ser As series
        For Each ser In m_Chart.SeriesCollection
            ser.Format.Line.DashStyle = value
        Next ser
    End If
End Property

Public Property Get SeriesMarkerStyle() As XlMarkerStyle
    SeriesMarkerStyle = m_SeriesMarkerStyle
End Property

Public Property Let SeriesMarkerStyle(ByVal value As XlMarkerStyle)
    m_SeriesMarkerStyle = value
    If Not m_Chart Is Nothing Then
        Dim ser As series
        For Each ser In m_Chart.SeriesCollection
            ser.MarkerStyle = value
        Next ser
    End If
End Property

Public Property Get HasDataLabels() As Boolean
    HasDataLabels = m_HasDataLabels
End Property

Public Property Let HasDataLabels(ByVal value As Boolean)
    m_HasDataLabels = value
    If Not m_Chart Is Nothing Then
        Dim ser As series
        For Each ser In m_Chart.SeriesCollection
            ser.HasDataLabels = value
        Next ser
    End If
End Property

Public Property Get X1LabelNumberFormat() As String
    X1LabelNumberFormat = m_X1LabelNumberFormat
End Property

Public Property Let X1LabelNumberFormat(ByVal value As String)
    m_X1LabelNumberFormat = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).TickLabels.NumberFormat = value
End Property

Public Property Get Y1LabelNumberFormat() As String
    Y1LabelNumberFormat = m_Y1LabelNumberFormat
End Property

Public Property Let Y1LabelNumberFormat(ByVal value As String)
    m_Y1LabelNumberFormat = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = value
End Property

Public Property Get X2LabelNumberFormat() As String
    X2LabelNumberFormat = m_X2LabelNumberFormat
End Property

Public Property Let X2LabelNumberFormat(ByVal value As String)
    m_X2LabelNumberFormat = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlSecondary).TickLabels.NumberFormat = value
End Property

Public Property Get Y2LabelNumberFormat() As String
    Y2LabelNumberFormat = m_Y2LabelNumberFormat
End Property

Public Property Let Y2LabelNumberFormat(ByVal value As String)
    m_Y2LabelNumberFormat = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = value
End Property

Public Property Get Y1Min() As Variant
    Y1Min = m_Y1Min
End Property

Public Property Let Y1Min(ByVal value As Variant)
    m_Y1Min = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlValue, xlPrimary).MinimumScale = value
    End If
End Property

Public Property Get Y1Max() As Variant
    Y1Max = m_Y1Max
End Property

Public Property Let Y1Max(ByVal value As Variant)
    m_Y1Max = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlValue, xlPrimary).MaximumScale = value
    End If
End Property

Public Property Get X1Min() As Variant
    X1Min = m_X1Min
End Property

Public Property Let X1Min(ByVal value As Variant)
    m_X1Min = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlCategory, xlPrimary).MinimumScale = value
    End If
End Property

Public Property Get X1Max() As Variant
    X1Max = m_X1Max
End Property

Public Property Let X1Max(ByVal value As Variant)
    m_X1Max = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlCategory, xlPrimary).MaximumScale = value
    End If
End Property

Public Property Get Y2Min() As Variant
    Y2Min = m_Y2Min
End Property

Public Property Let Y2Min(ByVal value As Variant)
    m_Y2Min = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlValue, xlSecondary).MinimumScale = value
    End If
End Property

Public Property Get Y2Max() As Variant
    Y2Max = m_Y2Max
End Property

Public Property Let Y2Max(ByVal value As Variant)
    m_Y2Max = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlValue, xlSecondary).MaximumScale = value
    End If
End Property

Public Property Get X2Min() As Variant
    X2Min = m_X2Min
End Property

Public Property Let X2Min(ByVal value As Variant)
    m_X2Min = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlCategory, xlSecondary).MinimumScale = value
    End If
End Property

Public Property Get X2Max() As Variant
    X2Max = m_X2Max
End Property

Public Property Let X2Max(ByVal value As Variant)
    m_X2Max = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlCategory, xlSecondary).MaximumScale = value
    End If
End Property

Public Property Get SquarePlot() As Boolean
    SquarePlot = m_SquarePlot
End Property

Public Property Let SquarePlot(ByVal value As Boolean)
    m_SquarePlot = value
    If Not m_Chart Is Nothing Then
        If value Then
            With m_Chart.PlotArea
                .Height = .Width
            End With
        End If
    End If
End Property

Public Property Get CrossesAt() As Double
    CrossesAt = m_CrossesAt
End Property

Public Property Let CrossesAt(ByVal value As Double)
    m_CrossesAt = value
    If Not m_Chart Is Nothing Then
        m_Chart.Axes(xlCategory, xlPrimary).CrossesAt = value
        m_Chart.Axes(xlValue, xlPrimary).CrossesAt = value
    End If
End Property

Public Property Get PlotBy() As XlRowCol
    PlotBy = m_PlotBy
End Property

Public Property Let PlotBy(ByVal value As XlRowCol)
    m_PlotBy = value
    If Not m_Chart Is Nothing Then
        m_Chart.PlotBy = value
    End If
End Property

Private Sub Class_Initialize()
    ' Initialize the class with default values
    m_ChartType = xlXYScatterLinesNoMarkers
    m_ChartTitle = "Chart Title"
    m_Y1Title = "Y-Axis"
    m_X1Title = "X-Axis"
    m_HasLegend = True
    m_HasTitle = True
    m_X1HasTitle = True
    m_Y1HasTitle = True
    m_X1ShowAxis = True
    m_Y1ShowAxis = True
    m_X1ShowGridLines = True
    m_Y1ShowGridLines = True
    m_FontName = "Calibri"
    m_TitleFontSize = 14
    m_XAxisTitleFontSize = 12
    m_YAxisTitleFontSize = 12
    m_XAxisTicksFontSize = 10
    m_YAxisTicksFontSize = 10
    m_LegendFontSize = 10
    m_LegendPosition = xlLegendPositionBottom
    m_LegendInLayout = False
    m_PlotAreaLineWeight = 0.75
    m_SeriesLineWeight = 1.5
    m_SeriesLineDashStyle = msoLineSolid
    m_SeriesMarkerStyle = xlMarkerStyleNone
    m_HasDataLabels = False
    m_X1LabelNumberFormat = "General"
    m_Y1LabelNumberFormat = "General"
    m_X2LabelNumberFormat = "General"
    m_Y2LabelNumberFormat = "General"
    m_Y1Min = "Auto"
    m_Y1Max = "Auto"
    m_X1Min = "Auto"
    m_X1Max = "Auto"
    m_Y2Min = "Auto"
    m_Y2Max = "Auto"
    m_X2Min = "Auto"
    m_X2Max = "Auto"
    m_SquarePlot = False
    m_CrossesAt = 0
    m_PlotBy = xlColumns
End Sub

Public Sub FormatChart()
    On Error GoTo ErrHandler

    Dim ser As series
    Dim axNo As Axis

    With m_Chart
        .ChartType = m_ChartType
        .HasTitle = m_HasTitle
        If .HasTitle Then
            .chartTitle.Text = m_ChartTitle
            .chartTitle.Font.Size = m_TitleFontSize
        End If

        .HasLegend = m_HasLegend
        If .HasLegend Then
            With .Legend
                .Position = m_LegendPosition
                .IncludeInLayout = m_LegendInLayout
                .Font.Size = m_LegendFontSize
            End With
        End If

        For Each ser In .SeriesCollection
            With ser
                .format.line.Weight = m_SeriesLineWeight
                .format.line.DashStyle = m_SeriesLineDashStyle
                .MarkerStyle = m_SeriesMarkerStyle
                .HasDataLabels = m_HasDataLabels
            End With
        Next ser
    End With

    Exit Sub
ErrHandler:
    HandleError "Error formatting chart: " & Err.Description
End Sub

Public Sub CreateSmithChart()
    On Error GoTo ErrHandler

    Dim srcChart As Chart
    Dim myChart As Chart

    Set srcChart = GetSelectedChart()
    If srcChart Is Nothing Then Exit Sub

    Set myChart = Copy_Chart(srcChart)
    If myChart Is Nothing Then Call HandleError("Failed to copy the chart."): Exit Sub

    If Not delete_UnMatchingSeries(myChart, ":", False) Then Call HandleError("Failed to clean up Real/Imaginary data."): Exit Sub
    If Not SmithChart_Formatting(myChart, True) Then Call HandleError("Failed to format the Smith chart."): Exit Sub
    If Not Delink_ChartData(myChart) Then Call HandleError("Failed to delink chart data."): Exit Sub

    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in Create_SmithChart.")
End Sub

Public Sub CreateYLog()
    On Error GoTo ErrHandler

    Dim srcChart As Chart
    Dim myChart As Chart

    Set srcChart = GetSelectedChart()
    If srcChart Is Nothing Then Exit Sub

    Set myChart = Copy_Chart(srcChart)
    If myChart Is Nothing Then Call HandleError("Failed to copy the chart."): Exit Sub

    If Not Chart_YLog(myChart) Then Call HandleError("Failed to generate logarithmic Y-axis scale."): Exit Sub
    If Not Delink_ChartData(myChart) Then Call HandleError("Failed to delink chart data."): Exit Sub

    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in Create_YLog.")
End Sub

Public Sub CreateDerivative()
    On Error GoTo ErrHandler

    Dim srcChart As Chart
    Dim myChart As Chart

    Set srcChart = GetSelectedChart()
    If srcChart Is Nothing Then Exit Sub

    Set myChart = Copy_Chart(srcChart)
    If myChart Is Nothing Then Call HandleError("Failed to copy the chart."): Exit Sub

    If Not Chart_Derivate(myChart) Then Call HandleError("Failed to generate the derivative."): Exit Sub
    If Not Delink_ChartData(myChart) Then Call HandleError("Failed to delink chart data."): Exit Sub

    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in Create_Derivative.")
End Sub

Public Sub ExportAsImage()
    On Error GoTo ErrHandler

    Debug.Print "Calling ExportAsImage."

    If Export_ChartAsImage() Then
        MsgBox "The chart was successfully exported!", vbInformation, "Export Successful"
    Else
        MsgBox "The chart export failed or was canceled.", vbExclamation, "Export Failed"
    End If

    Exit Sub

ErrHandler:
    Call HandleError("An unexpected error occurred in ExportAsImage.")
    MsgBox "An internal error occurred during export.", vbCritical, "Critical Error"
End Sub

Private Function HandleError(ByVal ErrMsg As String)
    MsgBox "An error occurred:" & vbCrLf & ErrMsg, vbOKOnly + vbCritical, "Operation impossible"
End Function

Private Function GetSelectedChart() As Chart
    On Error GoTo ErrHandle

    If Not ActiveChart Is Nothing Then
        Set GetSelectedChart = ActiveChart
    Else
        Call HandleError("You must select a chart to perform this operation.")
        Set GetSelectedChart = Nothing
    End If
    Exit Function

ErrHandle:
    Call HandleError("Error getting selected chart.")
    Set GetSelectedChart = Nothing
End Function

Private Function Copy_Chart(myChart As Chart, Optional cpyPosition As String = "Right", Optional offsetPosition As Double = 0) As Variant
    On Error GoTo ifError
    Dim parentType  As String
    Dim newChart    As Chart
    Dim wks_des     As Worksheet

    parentType = TypeName(myChart.parent)
    Select Case TypeName(myChart.parent)
        Case "Workbook"
            Select Case UCase(cpyPosition)
                Case "RIGHT", "AFTER": myChart.Copy After:=myChart
                Case "LEFT", "BEFORE": myChart.Copy Before:=myChart
                Case Else: myChart.Copy After:=myChart
            End Select
            Set newChart = ActiveChart
        Case Else
            Set wks_des = myChart.parent.parent
            myChart.parent.Copy
            wks_des.Paste
            Set newChart = wks_des.ChartObjects(wks_des.ChartObjects.count).Chart
            newChart.parent.Height = myChart.parent.Height
            newChart.parent.Width = myChart.parent.Width
            newChart.parent.Top = myChart.parent.Top
            newChart.parent.Left = myChart.parent.Left
            Select Case UCase(cpyPosition)
                Case "RIGHT": newChart.parent.Left = myChart.parent.Left + myChart.parent.Width + 5 + offsetPosition
                Case "LEFT": newChart.parent.Left = IIf(myChart.parent.Left - myChart.parent.Width - 5 - offsetPosition > 0, myChart.parent.Left - myChart.parent.Width - 5 - offsetPosition, 0)
                Case "DOWN": newChart.parent.Top = myChart.parent.Top + myChart.parent.Height + 5 + offsetPosition
                Case "UP": newChart.parent.Top = IIf(myChart.parent.Top - myChart.parent.Height - 5 - offsetPosition > 0, myChart.parent.Top - myChart.parent.Height - 5 - offsetPosition, 0)
                Case Else: newChart.parent.Left = myChart.parent.Left + myChart.parent.Width + 5 + offsetPosition
            End Select
    End Select

    Set Copy_Chart = newChart
    Exit Function
ifError:
    Copy_Chart = CVErr(2001)
End Function

Private Function delete_UnMatchingSeries(myChart As Chart, sMatch As String, Optional compareAbsolute As Boolean = False) As Boolean
    Dim i       As Integer

    With myChart
        For i = 1 To .FullSeriesCollection.count
            If compareAbsolute Then
                If .FullSeriesCollection(i).name <> sMatch Then .FullSeriesCollection(i).Delete: i = i - 1
            Else
                If InStr(.FullSeriesCollection(i).name, sMatch) = 0 Then .FullSeriesCollection(i).Delete: i = i - 1
            End If
            If i + 1 > .FullSeriesCollection.count Then Exit For
        Next i

    End With

    delete_UnMatchingSeries = True

End Function

Private Function SmithChart_Formatting(myChart As Chart, Optional AddBackGround As Boolean = True, Optional FormattingOptions As String = "") As Boolean
    Dim cSerie1         As Variant
    Dim cSerie2         As Variant
    Dim picturePath     As String
    Dim isSmithData     As Boolean
    Dim fmtOptions      As String

    With myChart
        For Each cSerie1 In .FullSeriesCollection
            If cSerie1.name Like "*I:*" And Not cSerie1.IsFiltered Then
                For Each cSerie2 In .FullSeriesCollection
                    If Replace(cSerie1.name, "I:", "") = Replace(cSerie2.name, "R:", "") Then
                        isSmithData = True
                        Exit For
                    End If
                Next
            End If
        Next
        If Not isSmithData Then SmithChart_Formatting = False: Exit Function

        fmtOptions = "Y1Min=-1;Y1Max=1;X1Min=-1;X1Max=1;HasTitle=false;HasLegend=false;SquarePlot=1;X1HasTitle=false;Y1HasTitle=false;X1ShowGridLines=false;Y1ShowGridLines=false;X1ShowAxis=false;Y1ShowAxis=false" & _
                        ";X1ScaleType=" & xlLinear & ";Y1ScaleType=" & xlLinear
        ChartFormatting myChart, fmtOptions & ";" & FormattingOptions

        For Each cSerie1 In .FullSeriesCollection
            If cSerie1.name Like "*R:*" And Not cSerie1.IsFiltered Then
                For Each cSerie2 In .FullSeriesCollection
                    If Replace(cSerie1.name, "R:", "") = Replace(cSerie2.name, "I:", "") Then
                        cSerie2.XValues = cSerie1.values
                        cSerie1.Delete
                        Exit For
                    End If
                Next
            End If
        Next
        If AddBackGround Then
            picturePath = "SmithChart_BackGround.png" ' This needs to be a valid path
            If FalFile.FileExist(picturePath) Then
                .PlotArea.format.Fill.Visible = msoTrue
                .PlotArea.format.Fill.UserPicture picturePath
                .PlotArea.format.Fill.TextureTile = msoFalse
                .PlotArea.format.Fill.Transparency = 0.6
            End If
        Else
            .PlotArea.format.Fill.Visible = msoFalse
        End If
    End With

    SmithChart_Formatting = True
End Function

Private Function Delink_ChartData(myChart As Chart) As Boolean
    On Error GoTo ifError
    Dim cSerie          As Variant
    For Each cSerie In myChart.FullSeriesCollection
        cSerie.XValues = cSerie.XValues
        cSerie.values = cSerie.values
        cSerie.name = cSerie.name
    Next
    Delink_ChartData = True
    Exit Function
ifError:
    Delink_ChartData = False
End Function

Private Function Chart_YLog(myChart As Chart) As Boolean
    Chart_YAbsolute myChart
    myChart.Axes(xlValue, xlPrimary).ScaleType = xlLogarithmic
    Chart_YLog = True
End Function

Private Function Chart_YAbsolute(myChart As Chart) As Boolean
    Dim cSerie          As Variant
    For Each cSerie In myChart.FullSeriesCollection
        cSerie.values = FalArray.a1D_math_Abs(cSerie.values)
    Next
    Chart_YAbsolute = True
End Function

Private Function Chart_Derivate(myChart As Chart, Optional XSample As Long = 3) As Boolean
    Dim cSerie          As Variant
    Dim a2D_X           As Variant
    Dim a2D_Y           As Variant
    Dim a2D_XY          As Variant
    Dim a1D_dY          As Variant

    With myChart
        For Each cSerie In .FullSeriesCollection
            a2D_X = FalArray.a1D_To_Columna2D(cSerie.XValues)
            a2D_Y = FalArray.a1D_To_Columna2D(cSerie.values)
            a2D_XY = FalArray.a2D_Merge_ByColumn(a2D_X, a2D_Y)
            a2D_dY = FalArray.a2D_math_Derivate(a2D_XY, XSample, 2)
            a1D_dY = FalArray.a2D_to_a1D_Column(a2D_dY, 2)
            cSerie.values = a1D_dY
            cSerie.name = "d" & cSerie.name
        Next
    End With
    Chart_Derivate = True
End Function
