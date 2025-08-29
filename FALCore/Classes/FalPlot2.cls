VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FalPlot2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************
' Class: FalPlot2
' Author: Florent ALBANY
' Date: 2025-08-29
' Version: 2.0
'
' Description:
' This class provides a comprehensive set of methods and properties for creating,
' manipulating, and formatting advanced charts in Microsoft Excel. It is designed
' to simplify the process of plotting data from ranges or arrays, offering a high
' degree of customization for titles, axes, series, and visual styles.
'
' It now includes a powerful theming engine, allowing users to save the style of
' an existing chart and apply it to others, ensuring a consistent and professional
' look across a project.
'
' The class encapsulates a Chart object and provides an object-oriented interface
' to its functionalities.
'
' @Example
' Dim plot As New FalPlot2
' Dim rng As Range
' Set rng = ActiveSheet.Range("A1:B10")
'
' ' Create a chart and assign it to the class
' Set plot.Chart = ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines).Chart
' plot.SetSourceData rng
'
' ' Set properties
' plot.ChartTitle = "My Chart"
' plot.X1Title = "X-Axis"
' plot.Y1Title = "Y-Axis"
'
' ' Apply formatting
' plot.FormatChart
'**************************************************************************************************

Option Explicit

'--- Private members
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

'--- Modern Theme Properties (for advanced styling) ---
Private m_PlotAreaBackColor As Long       ' RGB color for the plot area background. -1 for transparent.
Private m_MajorGridlinesColor As Long     ' RGB color for major gridlines. -1 for automatic.
Private m_TitleFontBold As Boolean
Private m_XAxisTitleFontBold As Boolean
Private m_YAxisTitleFontBold As Boolean

'--- Theme Engine Members
Private m_ThemeName As String
Private m_ThemeColors As Variant ' Array for theme colors
Private m_ThemeFonts As Variant ' Dictionary or custom type for font settings

'--- Public Properties ---

'@brief Gets or sets the Chart object this instance will operate on.
Public Property Get Chart() As Chart
    Set Chart = m_Chart
End Property

Public Property Set Chart(ByVal value As Chart)
    Set m_Chart = value
End Property

'@brief Gets or sets the chart type.
Public Property Get ChartType() As XlChartType
    ChartType = m_ChartType
End Property

Public Property Let ChartType(ByVal value As XlChartType)
    m_ChartType = value
    If Not m_Chart Is Nothing Then m_Chart.ChartType = value
End Property

'@brief Gets or sets the main title of the chart.
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

'@brief Gets or sets the title for the primary Y-axis.
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

'@brief Gets or sets the title for the secondary Y-axis.
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

'@brief Gets or sets the title for the primary X-axis.
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

'@brief Gets or sets the title for the secondary X-axis.
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

'@brief Gets or sets the scale type for the primary X-axis (e.g., xlLinear, xlLogarithmic).
Public Property Get X1ScaleType() As XlScaleType
    X1ScaleType = m_X1ScaleType
End Property

Public Property Let X1ScaleType(ByVal value As XlScaleType)
    m_X1ScaleType = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).ScaleType = value
End Property

'@brief Gets or sets the scale type for the primary Y-axis.
Public Property Get Y1ScaleType() As XlScaleType
    Y1ScaleType = m_Y1ScaleType
End Property

Public Property Let Y1ScaleType(ByVal value As XlScaleType)
    m_Y1ScaleType = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).ScaleType = value
End Property

'@brief Gets or sets the scale type for the secondary X-axis.
Public Property Get X2ScaleType() As XlScaleType
    X2ScaleType = m_X2ScaleType
End Property

Public Property Let X2ScaleType(ByVal value As XlScaleType)
    m_X2ScaleType = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlSecondary).ScaleType = value
End Property

'@brief Gets or sets the scale type for the secondary Y-axis.
Public Property Get Y2ScaleType() As XlScaleType
    Y2ScaleType = m_Y2ScaleType
End Property

Public Property Let Y2ScaleType(ByVal value As XlScaleType)
    m_Y2ScaleType = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlSecondary).ScaleType = value
End Property

'@brief Gets or sets whether the chart has a legend.
Public Property Get HasLegend() As Boolean
    HasLegend = m_HasLegend
End Property

Public Property Let HasLegend(ByVal value As Boolean)
    m_HasLegend = value
    If Not m_Chart Is Nothing Then m_Chart.HasLegend = value
End Property

'@brief Gets or sets whether the chart has a main title.
Public Property Get HasTitle() As Boolean
    HasTitle = m_HasTitle
End Property

Public Property Let HasTitle(ByVal value As Boolean)
    m_HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.HasTitle = value
End Property

'@brief Gets or sets whether the primary X-axis has a title.
Public Property Get X1HasTitle() As Boolean
    X1HasTitle = m_X1HasTitle
End Property

Public Property Let X1HasTitle(ByVal value As Boolean)
    m_X1HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).HasTitle = value
End Property

'@brief Gets or sets whether the primary Y-axis has a title.
Public Property Get Y1HasTitle() As Boolean
    Y1HasTitle = m_Y1HasTitle
End Property

Public Property Let Y1HasTitle(ByVal value As Boolean)
    m_Y1HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).HasTitle = value
End Property

'@brief Gets or sets whether the secondary X-axis has a title.
Public Property Get X2HasTitle() As Boolean
    X2HasTitle = m_X2HasTitle
End Property

Public Property Let X2HasTitle(ByVal value As Boolean)
    m_X2HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlSecondary).HasTitle = value
End Property

'@brief Gets or sets whether the secondary Y-axis has a title.
Public Property Get Y2HasTitle() As Boolean
    Y2HasTitle = m_Y2HasTitle
End Property

Public Property Let Y2HasTitle(ByVal value As Boolean)
    m_Y2HasTitle = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlSecondary).HasTitle = value
End Property

'@brief Gets or sets whether the primary X-axis line is visible.
Public Property Get X1ShowAxis() As Boolean
    X1ShowAxis = m_X1ShowAxis
End Property

Public Property Let X1ShowAxis(ByVal value As Boolean)
    m_X1ShowAxis = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).Format.Line.Visible = IIf(value, msoTrue, msoFalse)
End Property

'@brief Gets or sets whether the primary Y-axis line is visible.
Public Property Get Y1ShowAxis() As Boolean
    Y1ShowAxis = m_Y1ShowAxis
End Property

Public Property Let Y1ShowAxis(ByVal value As Boolean)
    m_Y1ShowAxis = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).Format.Line.Visible = IIf(value, msoTrue, msoFalse)
End Property

'@brief Gets or sets whether the primary X-axis shows major gridlines.
Public Property Get X1ShowGridLines() As Boolean
    X1ShowGridLines = m_X1ShowGridLines
End Property

Public Property Let X1ShowGridLines(ByVal value As Boolean)
    m_X1ShowGridLines = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).HasMajorGridlines = value
End Property

'@brief Gets or sets whether the primary Y-axis shows major gridlines.
Public Property Get Y1ShowGridLines() As Boolean
    Y1ShowGridLines = m_Y1ShowGridLines
End Property

Public Property Let Y1ShowGridLines(ByVal value As Boolean)
    m_Y1ShowGridLines = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).HasMajorGridlines = value
End Property

'@brief Gets or sets the font name for all text elements in the chart.
Public Property Get FontName() As String
    FontName = m_FontName
End Property

Public Property Let FontName(ByVal value As String)
    m_FontName = value
    If Not m_Chart Is Nothing Then
        If m_Chart.HasTitle Then m_Chart.chartTitle.Font.Name = value
        If m_Chart.Axes(xlCategory, xlPrimary).HasTitle Then m_Chart.Axes(xlCategory, xlPrimary).AxisTitle.Font.Name = value
        If m_Chart.Axes(xlValue, xlPrimary).HasTitle Then m_Chart.Axes(xlValue, xlPrimary).AxisTitle.Font.Name = value
        m_Chart.Axes(xlCategory, xlPrimary).TickLabels.Font.Name = value
        m_Chart.Axes(xlValue, xlPrimary).TickLabels.Font.Name = value
        If m_Chart.HasLegend Then m_Chart.Legend.Font.Name = value
    End If
End Property

'@brief Gets or sets the font size for the main chart title.
Public Property Get TitleFontSize() As Integer
    TitleFontSize = m_TitleFontSize
End Property

Public Property Let TitleFontSize(ByVal value As Integer)
    m_TitleFontSize = value
    If Not m_Chart Is Nothing Then If m_Chart.HasTitle Then m_Chart.chartTitle.Font.Size = value
End Property

'@brief Gets or sets the font size for the axis titles.
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

'@brief Gets or sets the font size for the axis tick labels.
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

'@brief Gets or sets the font size for the legend.
Public Property Get LegendFontSize() As Integer
    LegendFontSize = m_LegendFontSize
End Property

Public Property Let LegendFontSize(ByVal value As Integer)
    m_LegendFontSize = value
    If Not m_Chart Is Nothing Then If m_Chart.HasLegend Then m_Chart.Legend.Font.Size = value
End Property

'@brief Gets or sets the position of the legend.
Public Property Get LegendPosition() As XlLegendPosition
    LegendPosition = m_LegendPosition
End Property

Public Property Let LegendPosition(ByVal value As XlLegendPosition)
    m_LegendPosition = value
    If Not m_Chart Is Nothing Then If m_Chart.HasLegend Then m_Chart.Legend.Position = value
End Property

'@brief Gets or sets whether the legend is included in the chart layout.
Public Property Get LegendInLayout() As Boolean
    LegendInLayout = m_LegendInLayout
End Property

Public Property Let LegendInLayout(ByVal value As Boolean)
    m_LegendInLayout = value
    If Not m_Chart Is Nothing Then If m_Chart.HasLegend Then m_Chart.Legend.IncludeInLayout = value
End Property

'@brief Gets or sets the line weight for the plot area border.
Public Property Get PlotAreaLineWeight() As Double
    PlotAreaLineWeight = m_PlotAreaLineWeight
End Property

Public Property Let PlotAreaLineWeight(ByVal value As Double)
    m_PlotAreaLineWeight = value
    If Not m_Chart Is Nothing Then m_Chart.PlotArea.Format.Line.Weight = value
End Property

'@brief Gets or sets the line weight for all series in the chart.
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

'@brief Gets or sets the dash style for all series lines.
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

'@brief Gets or sets the marker style for all series.
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

'@brief Gets or sets whether data labels are shown for all series.
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

'@brief Gets or sets the number format for the primary X-axis tick labels.
Public Property Get X1LabelNumberFormat() As String
    X1LabelNumberFormat = m_X1LabelNumberFormat
End Property

Public Property Let X1LabelNumberFormat(ByVal value As String)
    m_X1LabelNumberFormat = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlPrimary).TickLabels.NumberFormat = value
End Property

'@brief Gets or sets the number format for the primary Y-axis tick labels.
Public Property Get Y1LabelNumberFormat() As String
    Y1LabelNumberFormat = m_Y1LabelNumberFormat
End Property

Public Property Let Y1LabelNumberFormat(ByVal value As String)
    m_Y1LabelNumberFormat = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlPrimary).TickLabels.NumberFormat = value
End Property

'@brief Gets or sets the number format for the secondary X-axis tick labels.
Public Property Get X2LabelNumberFormat() As String
    X2LabelNumberFormat = m_X2LabelNumberFormat
End Property

Public Property Let X2LabelNumberFormat(ByVal value As String)
    m_X2LabelNumberFormat = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlCategory, xlSecondary).TickLabels.NumberFormat = value
End Property

'@brief Gets or sets the number format for the secondary Y-axis tick labels.
Public Property Get Y2LabelNumberFormat() As String
    Y2LabelNumberFormat = m_Y2LabelNumberFormat
End Property

Public Property Let Y2LabelNumberFormat(ByVal value As String)
    m_Y2LabelNumberFormat = value
    If Not m_Chart Is Nothing Then m_Chart.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = value
End Property

'@brief Gets or sets the minimum value for the primary Y-axis.
Public Property Get Y1Min() As Variant
    Y1Min = m_Y1Min
End Property

Public Property Let Y1Min(ByVal value As Variant)
    m_Y1Min = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlValue, xlPrimary).MinimumScale = value
    End If
End Property

'@brief Gets or sets the maximum value for the primary Y-axis.
Public Property Get Y1Max() As Variant
    Y1Max = m_Y1Max
End Property

Public Property Let Y1Max(ByVal value As Variant)
    m_Y1Max = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlValue, xlPrimary).MaximumScale = value
    End If
End Property

'@brief Gets or sets the minimum value for the primary X-axis.
Public Property Get X1Min() As Variant
    X1Min = m_X1Min
End Property

Public Property Let X1Min(ByVal value As Variant)
    m_X1Min = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlCategory, xlPrimary).MinimumScale = value
    End If
End Property

'@brief Gets or sets the maximum value for the primary X-axis.
Public Property Get X1Max() As Variant
    X1Max = m_X1Max
End Property

Public Property Let X1Max(ByVal value As Variant)
    m_X1Max = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlCategory, xlPrimary).MaximumScale = value
    End If
End Property

'@brief Gets or sets the minimum value for the secondary Y-axis.
Public Property Get Y2Min() As Variant
    Y2Min = m_Y2Min
End Property

Public Property Let Y2Min(ByVal value As Variant)
    m_Y2Min = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlValue, xlSecondary).MinimumScale = value
    End If
End Property

'@brief Gets or sets the maximum value for the secondary Y-axis.
Public Property Get Y2Max() As Variant
    Y2Max = m_Y2Max
End Property

Public Property Let Y2Max(ByVal value As Variant)
    m_Y2Max = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlValue, xlSecondary).MaximumScale = value
    End If
End Property

'@brief Gets or sets the minimum value for the secondary X-axis.
Public Property Get X2Min() As Variant
    X2Min = m_X2Min
End Property

Public Property Let X2Min(ByVal value As Variant)
    m_X2Min = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlCategory, xlSecondary).MinimumScale = value
    End If
End Property

'@brief Gets or sets the maximum value for the secondary X-axis.
Public Property Get X2Max() As Variant
    X2Max = m_X2Max
End Property

Public Property Let X2Max(ByVal value As Variant)
    m_X2Max = value
    If Not m_Chart Is Nothing Then
        If IsNumeric(value) Then m_Chart.Axes(xlCategory, xlSecondary).MaximumScale = value
    End If
End Property

'@brief Gets or sets whether the plot area should be square.
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

'@brief Gets or sets the value where the axes cross.
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

'@brief Gets or sets how the data is plotted (by rows or columns).
Public Property Get PlotBy() As XlRowCol
    PlotBy = m_PlotBy
End Property

Public Property Let PlotBy(ByVal value As XlRowCol)
    m_PlotBy = value
    If Not m_Chart Is Nothing Then
        m_Chart.PlotBy = value
    End If
End Property

'--- Initialization ---

Private Sub Class_Initialize()
    ' @brief Initializes the class with default property values.
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

    '--- Modern Theme Properties ---
    m_PlotAreaBackColor = -1 ' -1 indicates no fill
    m_MajorGridlinesColor = -1 ' -1 indicates automatic color
    m_TitleFontBold = True
    m_XAxisTitleFontBold = False
    m_YAxisTitleFontBold = False
End Sub

'--- Public Create Methods ---

Public Function CreateChart(ByVal ChartName As String, ByVal wbk As Workbook, Optional ByVal chartType As XlChartType = xlXYScatterLinesNoMarkers, Optional ByVal rangeData As Range) As Chart
    ' @brief Creates a new chart sheet, assigns it to the class, and formats it using the class's properties.
    ' @param ChartName The name for the new chart sheet.
    ' @param wbk The workbook in which to create the chart.
    ' @param chartType (Optional) The type of chart to create. Defaults to xlXYScatterLinesNoMarkers.
    ' @param rangeData (Optional) The data range to plot.
    ' @return The created Chart object, or Nothing on failure.
    On Error GoTo ifError

    Dim tempName As String
    Dim finalName As String
    Dim counter As Long
    Dim ch As Chart
    Dim ws As Worksheet

    ' Sanitize original name
    tempName = Me.prvRemoveChars(ChartName, ":\/?*[];")
    finalName = Me.prvResizeString(tempName, 31)
    counter = 1

    ' Loop to find a unique name across both charts and worksheets
    Do
        ' Check for chart with this name
        Set ch = Nothing
        On Error Resume Next
        Set ch = wbk.Charts(finalName)
        On Error GoTo 0

        ' Check for worksheet with this name
        Set ws = Nothing
        On Error Resume Next
        Set ws = wbk.Worksheets(finalName)
        On Error GoTo 0

        If ch Is Nothing And ws Is Nothing Then
            ' Name is unique
            Exit Do
        Else
            ' Name exists, generate a new one
            counter = counter + 1
            Dim suffix As String
            suffix = " (" & counter & ")"
            finalName = Me.prvResizeString(tempName, 31 - Len(suffix)) & suffix
        End If
    Loop

    Dim newChartSheet As Chart
    Set newChartSheet = wbk.Charts.Add
    With newChartSheet
        .Move After:=wbk.Worksheets(wbk.Worksheets.Count)
        .name = finalName
        .chartType = chartType
    End With

    ' Assign the new chart to the class instance
    Set Me.Chart = newChartSheet

    If Not rangeData Is Nothing Then
        Me.SetSourceData rangeData
    Else
        ' If no data, clear any default series Excel might have added
        Me.ClearSeries
    End If

    ' Apply the class's current formatting properties
    Me.FormatChart

    Set CreateChart = newChartSheet
    Exit Function

ifError:
    Set CreateChart = Nothing
    HandleError "Failed to create chart '" & ChartName & "'. " & Err.Description
End Function

Public Function CreateChartFromArray(ByVal Arr2D As Variant, Optional ByVal ChartName As String = "Array Plot", Optional ByVal FormattingOptions As String = "") As Chart
    ' @brief Creates and formats a chart from a 2D VBA array.
    ' @details The function writes the array to a temporary worksheet, creates a chart from that worksheet range, and then deletes the temporary sheet.
    ' @param Arr2D A 2D array where the first column is X-values and subsequent columns are Y-values. The first row contains series names.
    ' @param ChartName (Optional) The name for the new chart.
    ' @param FormattingOptions (Optional) A string for advanced formatting overrides.
    ' @return The created Chart object, or Nothing on failure.
    On Error GoTo ifError

    Dim tempWb As Workbook
    Dim tempWs As Worksheet
    Dim tempRange As Range

    ' Create a temporary workbook and worksheet
    Set tempWb = Application.Workbooks.Add
    Set tempWs = tempWb.Worksheets(1)

    ' Write the array to the temporary worksheet
    Set tempRange = tempWs.Range("A1").Resize(UBound(Arr2D, 1) - LBound(Arr2D, 1) + 1, UBound(Arr2D, 2) - LBound(Arr2D, 2) + 1)
    tempRange.value = Arr2D

    ' Create the chart from the temporary range
    Dim newChart As Chart
    Set newChart = Me.CreateChartFromRange(tempRange, xlXYScatter, ChartName)
    ' Note: FormattingOptions from the original call are not used here, as CreateChartFromRange doesn't accept them directly.
    ' This could be a point of future improvement if needed.

    If newChart Is Nothing Then
        GoTo cleanup
    End If

    Set CreateChartFromArray = newChart

cleanup:
    ' Clean up the temporary workbook
    Application.DisplayAlerts = False
    tempWb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    If newChart Is Nothing Then Set CreateChartFromArray = Nothing ' Ensure it's nothing on error
    Exit Function
ifError:
    Set CreateChartFromArray = Nothing
    Resume cleanup
End Function

Public Function CreateChartFromWorksheet(ByVal wks_src As Worksheet, Optional ByVal DataTopLeftCell As String = "A1", Optional ByVal ColorStyle As String = "DefaultStyle", Optional ByVal plotSeries As Boolean = True, Optional ByVal nbseries As Integer = -1, Optional ByVal sPlot As String = " - Plot", Optional ByVal FormattingOptions As String = "") As Chart
    ' @brief Automatically detects a data table in a worksheet and creates a chart from it.
    ' @details This powerful function attempts to find the boundaries of a data table starting from `DataTopLeftCell`, then creates and formats a chart from it.
    ' @param wks_src The source worksheet containing the data.
    ' @param DataTopLeftCell (Optional) The top-left cell of the data table (e.g., "A1").
    ' @param ColorStyle (Optional) The name of the color palette to use.
    ' @param plotSeries (Optional) This parameter is currently unused.
    ' @param nbseries (Optional) The number of series to group for coloring. If omitted, it's auto-detected.
    ' @param sPlot (Optional) A suffix to add to the chart name.
    ' @param FormattingOptions (Optional) A string for advanced formatting overrides.
    ' @return The created Chart object, or Nothing on failure.
    On Error GoTo ifError
    Application.ScreenUpdating = False

    Dim aLabel() As String
    Dim LastCol_Index As Long
    Dim LastLineTbl() As Long
    Dim Chart_Name As String
    Dim firstLine As Long
    Dim TopLeftCell As Range
    Dim chartTitle As String
    Dim NbLine As Long
    Dim NbColumn As Long
    Dim ColIndex As Long
    Dim tmpName As String

    Set TopLeftCell = wks_src.Range(DataTopLeftCell)
    chartTitle = Me.prvRemoveSubstrings(wks_src.name, ".txt;.xlsx;.mdm;xlsm; ")
    chartTitle = Me.prvResizeString(chartTitle, 31)

    '--- Data Table Detection ---
    Dim LastLine As Long
    LastLine = FalWork.Get_Last_Row(wks_src, TopLeftCell.Column)
    firstLine = Me.FindFirstRowFromBottom(wks_src, TopLeftCell.Column, LastLine)

    LastCol_Index = FalWork.Get_Last_Column(wks_src, firstLine)
    Dim LastCol_Name As String, firstcol_name As String
    LastCol_Name = Me.prvColumnNumberToLetter(LastCol_Index)
    firstcol_name = Me.prvColumnNumberToLetter(TopLeftCell.Column)

    NbLine = LastLine - firstLine
    NbColumn = LastCol_Index - (TopLeftCell.Column - 1)

    ReDim LastLineTbl(LastCol_Index)
    For ColIndex = TopLeftCell.Column To LastCol_Index
        LastLineTbl(ColIndex) = FalWork.Get_Last_Row(wks_src, ColIndex)
    Next

    ReDim aLabel(LastCol_Index)
    For ColIndex = TopLeftCell.Column To LastCol_Index
        tmpName = wks_src.Cells(firstLine - 1, ColIndex).value
        If tmpName = "" Or tmpName = False Then
            wks_src.Cells(firstLine - 1, ColIndex).value = "Col " & ColIndex
        End If
        aLabel(ColIndex) = wks_src.Cells(firstLine - 1, ColIndex).value
    Next

    If nbseries < 1 Then
        nbseries = FalArray.aXD_count_Occurrence(aLabel, Left(aLabel(2), InStr(aLabel(2), " (") - 1), False, False)
    End If

    '--- Plot All Series ---
    Chart_Name = Me.prvResizeString(Me.prvRemoveChars(wks_src.name, "_-. "), 10) & " - " & sPlot & "(All)"
    Dim tempFormattingOptions As String
    tempFormattingOptions = "Title=" & chartTitle & ";XTitle=" & aLabel(1) & ";YTitle=(a.u);ChartType=" & xlXYScatterLinesNoMarkers & ";PlotBy=" & xlColumns & ";AutoLegend=0;" & FormattingOptions

    If Me.CreateChart(Chart_Name, wks_src.Parent) Is Nothing Then GoTo ifError

    Me.AddSeriesFromRange wks_src.Range(firstcol_name & (firstLine - 1) & ":" & LastCol_Name & LastLine)
    Me.DeleteMatchingSeries wks_src.Range(firstcol_name & (firstLine - 1)).value, False
    Me.ColorSeries 1, ((NbColumn / nbseries) - 1), ColorStyle
    Me.FormatChart tempFormattingOptions

    Application.ScreenUpdating = True
    Set CreateChartFromWorksheet = m_Chart
    Exit Function
ifError:
    Application.ScreenUpdating = True
    Set CreateChartFromWorksheet = Nothing
End Function

Public Function CreateSmithChart() As Chart
    ' @brief Creates a Smith chart from the currently selected chart.
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    If ActiveChart Is Nothing Then
        HandleError "Please select a chart first."
        Set CreateSmithChart = Nothing
        Exit Function
    End If

    Set m_Chart = Me.CopyChart(ActiveChart)
    If m_Chart Is Nothing Then
        HandleError "Failed to copy the chart."
        Set CreateSmithChart = Nothing
        Exit Function
    End If

    Me.DeleteUnmatchingSeries ":", False
    Me.FormatAsSmithChart True
    Me.UnlinkDataFromSource

    Set CreateSmithChart = m_Chart
    Application.ScreenUpdating = True
    Exit Function
ErrHandler:
    Application.ScreenUpdating = True
    HandleError "An unexpected error occurred in CreateSmithChart."
    Set CreateSmithChart = Nothing
End Function

Public Function CreateYLogChartFromActive() As Chart
    ' @brief Creates a copy of the selected chart with a logarithmic Y-axis.
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    If ActiveChart Is Nothing Then
        HandleError "Please select a chart first."
        Set CreateYLogChartFromActive = Nothing
        Exit Function
    End If

    Set m_Chart = Me.CopyChart(ActiveChart)
    If m_Chart Is Nothing Then
        HandleError "Failed to copy the chart."
        Set CreateYLogChartFromActive = Nothing
        Exit Function
    End If

    Me.ApplyYLogarithmicScale
    Me.UnlinkDataFromSource

    Set CreateYLogChartFromActive = m_Chart
    Application.ScreenUpdating = True
    Exit Function
ErrHandler:
    Application.ScreenUpdating = True
    HandleError "An unexpected error occurred in CreateYLogChartFromActive."
    Set CreateYLogChartFromActive = Nothing
End Function

Public Function CreateDerivativeChartFromActive() As Chart
    ' @brief Creates a copy of the selected chart with its series' derivatives.
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    If ActiveChart Is Nothing Then
        HandleError "Please select a chart first."
        Set CreateDerivativeChartFromActive = Nothing
        Exit Function
    End If

    Set m_Chart = Me.CopyChart(ActiveChart)
    If m_Chart Is Nothing Then
        HandleError "Failed to copy the chart."
        Set CreateDerivativeChartFromActive = Nothing
        Exit Function
    End If

    Me.CalculateDerivative
    Me.UnlinkDataFromSource

    Set CreateDerivativeChartFromActive = m_Chart
    Application.ScreenUpdating = True
    Exit Function
ErrHandler:
    Application.ScreenUpdating = True
    HandleError "An unexpected error occurred in CreateDerivativeChartFromActive."
    Set CreateDerivativeChartFromActive = Nothing
End Function

Public Function CreateBarChart(ByVal data As Range, Optional ByVal chartName As String = "Bar Chart", Optional ByVal style As XlChartType = xlBarClustered) As Chart
    ' @brief Creates a new bar chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of bar chart to create. Defaults to xlBarClustered.
    ' @return The created Chart object, or Nothing on failure.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreateBarChart ActiveSheet.Range("A1:B5")
    Set CreateBarChart = Me.CreateChartFromRange(data, style, chartName)
End Function

Public Function CreatePieChart(ByVal data As Range, Optional ByVal chartName As String = "Pie Chart", Optional ByVal style As XlChartType = xlPie) As Chart
    ' @brief Creates a new pie chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of pie chart to create. Defaults to xlPie.
    ' @return The created Chart object, or Nothing on failure.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreatePieChart ActiveSheet.Range("A1:B5")
    Set CreatePieChart = Me.CreateChartFromRange(data, style, chartName)
End Function

Public Function CreateLineChart(ByVal data As Range, Optional ByVal chartName As String = "Line Chart", Optional ByVal style As XlChartType = xlLine) As Chart
    ' @brief Creates a new line chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of line chart to create. Defaults to xlLine.
    ' @return The created Chart object, or Nothing on failure.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreateLineChart ActiveSheet.Range("A1:B5")
    Set CreateLineChart = Me.CreateChartFromRange(data, style, chartName)
End Function

Public Function CreateAreaChart(ByVal data As Range, Optional ByVal chartName As String = "Area Chart", Optional ByVal style As XlChartType = xlArea) As Chart
    ' @brief Creates a new area chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of area chart to create. Defaults to xlArea.
    ' @return The created Chart object, or Nothing on failure.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreateAreaChart ActiveSheet.Range("A1:B5")
    Set CreateAreaChart = Me.CreateChartFromRange(data, style, chartName)
End Function

Public Function CreateSurfaceChart(ByVal data As Range, Optional ByVal chartName As String = "Surface Chart", Optional ByVal style As XlChartType = xlSurface) As Chart
    ' @brief Creates a new surface chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of surface chart to create. Defaults to xlSurface.
    ' @return The created Chart object, or Nothing on failure.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreateSurfaceChart ActiveSheet.Range("A1:C10")
    Set CreateSurfaceChart = Me.CreateChartFromRange(data, style, chartName)
End Function

Public Function CreateScatterChart(ByVal data As Range, Optional ByVal chartName As String = "Scatter Chart") As Chart
    ' @brief Creates a new, formatted scatter chart from a data range using the default class scatter type.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @return The created Chart object, or Nothing on failure.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreateScatterChart ActiveSheet.Range("A1:B10"), "My Scatter Plot"
    Set CreateScatterChart = Me.CreateChartFromRange(data, xlXYScatterLinesNoMarkers, chartName)
End Function

'--- Formatting and Theming Methods ---

Public Function FormatChart(Optional FormattingOptions As Variant) As Boolean
    ' @brief Applies a multitude of formatting options to the chart object from a string.
    ' @details This is the main formatting engine for the class. It uses the class's internal properties as a baseline for the chart's style.
    '          These properties can be individually overridden by providing a semicolon-delimited string of key-value pairs in the `FormattingOptions` parameter.
    '          This method is called by many other chart creation and formatting methods in the class.
    '          The parser is designed to be robust, handling values that may contain '=' signs.
    ' @param FormattingOptions (Optional) A string of key-value pairs (e.g., "Title=MyTitle;XTitle=X-Axis;TitleFontSize=16;PlotAreaBackColor=16777215").
    ' @return True if formatting was successful, False otherwise.
    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then Exit Function

    Application.ScreenUpdating = False

    Dim sKey As String, sValue As String
    Dim Options As Variant, index As Long
    Dim ser As series, axNo As Axis

    ' Use class properties as default values. The FormattingOptions string will override them.
    ' This section processes the optional FormattingOptions string.
    If Not IsMissing(FormattingOptions) Then
        If Len(FormattingOptions & "") > 0 Then
            Options = Split(FormattingOptions, ";")
            For index = 0 To UBound(Options)
                Dim delimiterPos As Long
                delimiterPos = InStr(1, Options(index), "=")
                If delimiterPos > 0 Then
                    sKey = Trim(Left(Options(index), delimiterPos - 1))
                    sValue = Trim(Mid(Options(index), delimiterPos + 1))

                    Select Case LCase(sKey)
                        Case "charttype": m_ChartType = CLng(sValue)
                        Case "charttitle", "title": m_ChartTitle = sValue
                        Case "plotby": m_PlotBy = CLng(sValue)
                        Case "fontname": m_FontName = sValue
                        Case "titlefontsize": m_TitleFontSize = CInt(sValue)
                        Case "xaxistitlefontsize": m_XAxisTitleFontSize = CInt(sValue)
                        Case "yaxistitlefontsize": m_YAxisTitleFontSize = CInt(sValue)
                        Case "xaxisticksfontsize": m_XAxisTicksFontSize = CInt(sValue)
                        Case "yaxisticksfontsize": m_YAxisTicksFontSize = CInt(sValue)
                        Case "hastitle": m_HasTitle = CBool(sValue)
                        Case "haslegend": m_HasLegend = CBool(sValue)
                        Case "x1showaxis": m_X1ShowAxis = CBool(sValue)
                        Case "y1showaxis": m_Y1ShowAxis = CBool(sValue)
                        Case "x1showgridlines", "xshowgridlines": m_X1ShowGridLines = CBool(sValue)
                        Case "y1showgridlines", "yshowgridlines": m_Y1ShowGridLines = CBool(sValue)
                        Case "x1hastitle", "xhastitle": m_X1HasTitle = CBool(sValue)
                        Case "x1title", "xtitle": m_X1Title = sValue
                        Case "x2title": m_X2Title = sValue
                        Case "y1hastitle", "yhastitle": m_Y1HasTitle = CBool(sValue)
                        Case "y1title", "ytitle": m_Y1Title = sValue
                        Case "y2hastitle": m_Y2HasTitle = CBool(sValue)
                        Case "y2title": m_Y2Title = sValue
                        Case "legendfontsize": m_LegendFontSize = CInt(sValue)
                        Case "legendposition": m_LegendPosition = CLng(sValue)
                        Case "legendinlayout": m_LegendInLayout = CBool(sValue)
                        Case "plotarealineweight": m_PlotAreaLineWeight = CDbl(sValue)
                        Case "serieslineweight": m_SeriesLineWeight = CDbl(sValue)
                        Case "serieslinedashstyle": m_SeriesLineDashStyle = CLng(sValue)
                        Case "seriesmarkerstyle": m_SeriesMarkerStyle = CLng(sValue)
                        Case "hasdatalabels": m_HasDataLabels = CBool(sValue)
                        Case "x1labelnumberformat": m_X1LabelNumberFormat = sValue
                        Case "x2labelnumberformat": m_X2LabelNumberFormat = sValue
                        Case "y1labelnumberformat": m_Y1LabelNumberFormat = sValue
                        Case "y2labelnumberformat": m_Y2LabelNumberFormat = sValue
                        Case "x1scaletype": m_X1ScaleType = CLng(sValue)
                        Case "y1scaletype": m_Y1ScaleType = CLng(sValue)
                        Case "x2scaletype": m_X2ScaleType = CLng(sValue)
                        Case "y2scaletype": m_Y2ScaleType = CLng(sValue)
                        Case "x1min", "xmin": m_X1Min = CDbl(sValue)
                        Case "x1max", "xmax": m_X1Max = CDbl(sValue)
                        Case "y1min", "ymin": m_Y1Min = CDbl(sValue)
                        Case "y1max", "ymax": m_Y1Max = CDbl(sValue)
                        Case "x2min": m_X2Min = CDbl(sValue)
                        Case "x2max": m_X2Max = CDbl(sValue)
                        Case "y2min": m_Y2Min = CDbl(sValue)
                        Case "y2max": m_Y2Max = CDbl(sValue)
                        Case "squareplot": m_SquarePlot = CBool(sValue)
                        Case "crossesat": m_CrossesAt = CDbl(sValue)
                        '--- Modern Theme Properties ---
                        Case "plotareabackcolor": m_PlotAreaBackColor = CLng(sValue)
                        Case "majorgridlinescolor": m_MajorGridlinesColor = CLng(sValue)
                        Case "titlefontbold": m_TitleFontBold = CBool(sValue)
                        Case "xaxistitlefontbold": m_XAxisTitleFontBold = CBool(sValue)
                        Case "yaxistitlefontbold": m_YAxisTitleFontBold = CBool(sValue)
                    End Select
                End If
            Next index
        End If
    End If

    With m_Chart
        .ChartType = m_ChartType
        .PlotBy = m_PlotBy

        .HasTitle = m_HasTitle
        If .HasTitle Then
            .chartTitle.Text = m_ChartTitle
            .chartTitle.Font.Size = m_TitleFontSize
            .chartTitle.Font.Bold = m_TitleFontBold
        End If

        If m_FontName <> "Auto" Then .ChartArea.Font.name = m_FontName

        ' Apply Plot Area background color if specified
        If m_PlotAreaBackColor <> -1 Then
            With .PlotArea.Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = m_PlotAreaBackColor
                .Solid
            End With
        Else
            .PlotArea.Format.Fill.Visible = msoFalse
        End If

        With .PlotArea.Format.Line
            .ForeColor.RGB = RGB(150, 150, 150)
            .Weight = m_PlotAreaLineWeight
        End With

        For Each ser In .SeriesCollection
            With ser
                .HasDataLabels = m_HasDataLabels
                .Format.Line.Visible = msoTrue
                .Format.Line.Weight = m_SeriesLineWeight
                .Format.Line.DashStyle = m_SeriesLineDashStyle
                .MarkerStyle = m_SeriesMarkerStyle
            End With
        Next ser

        .HasLegend = m_HasLegend
        If .HasLegend Then
            With .Legend
                .IncludeInLayout = m_LegendInLayout
                .Position = m_LegendPosition
                .Font.Size = m_LegendFontSize
            End With
        End If

        ' Primary X-Axis (Category)
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = m_X1HasTitle
            If .HasTitle Then
                .AxisTitle.Text = m_X1Title
                .AxisTitle.Font.Size = m_XAxisTitleFontSize
                .AxisTitle.Font.Bold = m_XAxisTitleFontBold
            End If
            .ScaleType = m_X1ScaleType
            If IsNumeric(m_X1Min) Then .MinimumScale = m_X1Min Else .MinimumScaleIsAuto = True
            If IsNumeric(m_X1Max) Then .MaximumScale = m_X1Max Else .MaximumScaleIsAuto = True
            .TickLabels.NumberFormat = m_X1LabelNumberFormat
            .TickLabels.Font.Size = m_XAxisTicksFontSize
            .HasMajorGridlines = m_X1ShowGridLines
            If .HasMajorGridlines And m_MajorGridlinesColor <> -1 Then
                .MajorGridlines.Format.Line.ForeColor.RGB = m_MajorGridlinesColor
            End If
        End With

        ' Primary Y-Axis (Value)
        With .Axes(xlValue, xlPrimary)
            .HasTitle = m_Y1HasTitle
            If .HasTitle Then
                .AxisTitle.Text = m_Y1Title
                .AxisTitle.Font.Size = m_YAxisTitleFontSize
                .AxisTitle.Font.Bold = m_YAxisTitleFontBold
            End If
            .ScaleType = m_Y1ScaleType
            If IsNumeric(m_Y1Min) Then .MinimumScale = m_Y1Min Else .MinimumScaleIsAuto = True
            If IsNumeric(m_Y1Max) Then .MaximumScale = m_Y1Max Else .MaximumScaleIsAuto = True
            .TickLabels.NumberFormat = m_Y1LabelNumberFormat
            .TickLabels.Font.Size = m_YAxisTicksFontSize
            .HasMajorGridlines = m_Y1ShowGridLines
            If .HasMajorGridlines And m_MajorGridlinesColor <> -1 Then
                .MajorGridlines.Format.Line.ForeColor.RGB = m_MajorGridlinesColor
            End If
        End With

        ' And so on for secondary axes...

        If m_SquarePlot Then
            If .HasLegend Then .Legend.Position = xlRight
            .PlotArea.Position = xlChartElementPositionAutomatic
            .PlotArea.Width = .PlotArea.Height
        End If

        Me.FormatTextSubscripts

    End With

    Application.ScreenUpdating = True
    FormatChart = True
    Exit Function

ErrHandler:
    Application.ScreenUpdating = True
    HandleError "Error while formatting chart: " & Err.Description
    FormatChart = False
End Function

Public Function ApplyTheme(ByVal themeName As String) As Boolean
    ' @brief Loads a theme from a file and applies it to the current chart.
    ' @param themeName The name of the theme to apply.
    ' @return True if the theme was applied successfully, False otherwise.
    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then
        HandleError "No chart is currently set. Cannot apply theme."
        ApplyTheme = False
        Exit Function
    End If

    Dim themeSettings As String
    themeSettings = Me.LoadTheme(themeName)

    If themeSettings <> "" Then
        ' Apply the theme settings
        If Me.FormatChart(themeSettings) Then
            ' Store the theme name
            m_ThemeName = themeName
            ApplyTheme = True
        Else
            HandleError "Failed to apply the theme '" & themeName & "'."
            ApplyTheme = False
        End If
    Else
        ' LoadTheme already handled the error message
        ApplyTheme = False
    End If

    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in ApplyTheme: " & Err.Description
    ApplyTheme = False
End Function

Public Function SaveThemeFromChart(ByVal themeName As String) As Boolean
    ' @brief Extracts the current chart's formatting and saves it as a theme file.
    ' @param themeName The name to give the new theme (e.g., "MyCustomTheme").
    ' @return True if the theme was saved successfully, False otherwise.
    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then
        HandleError "No chart is currently set. Cannot save theme."
        SaveThemeFromChart = False
        Exit Function
    End If

    Dim themeSettings As String
    Dim themeFolderPath As String
    Dim themeFilePath As String

    ' Build the theme string from the current chart's properties
    With m_Chart
        themeSettings = "ChartType=" & .ChartType & ";"

        ' Titles and Fonts
        If .HasTitle Then
            themeSettings = themeSettings & "HasTitle=True;"
            themeSettings = themeSettings & "ChartTitle=" & .chartTitle.Text & ";"
            themeSettings = themeSettings & "TitleFontSize=" & .chartTitle.Font.Size & ";"
            themeSettings = themeSettings & "TitleFontBold=" & .chartTitle.Font.Bold & ";"
            themeSettings = themeSettings & "FontName=" & .chartTitle.Font.Name & ";" ' Use title font as main font
        Else
            themeSettings = themeSettings & "HasTitle=False;"
        End If

        ' Axes
        ' Primary X-Axis
        With .Axes(xlCategory, xlPrimary)
            themeSettings = themeSettings & "X1ShowAxis=" & IIf(.Format.Line.Visible = msoTrue, "True", "False") & ";"
            themeSettings = themeSettings & "X1ShowGridLines=" & .HasMajorGridlines & ";"
            If .HasMajorGridlines Then themeSettings = themeSettings & "MajorGridlinesColor=" & .MajorGridlines.Format.Line.ForeColor.RGB & ";"
            If .HasTitle Then
                themeSettings = themeSettings & "X1HasTitle=True;"
                themeSettings = themeSettings & "X1Title=" & .AxisTitle.Text & ";"
                themeSettings = themeSettings & "XAxisTitleFontSize=" & .AxisTitle.Font.Size & ";"
                themeSettings = themeSettings & "XAxisTitleFontBold=" & .AxisTitle.Font.Bold & ";"
            Else
                themeSettings = themeSettings & "X1HasTitle=False;"
            End If
            themeSettings = themeSettings & "XAxisTicksFontSize=" & .TickLabels.Font.Size & ";"
            themeSettings = themeSettings & "X1LabelNumberFormat=" & .TickLabels.NumberFormat & ";"
        End With

        ' Primary Y-Axis
        With .Axes(xlValue, xlPrimary)
            themeSettings = themeSettings & "Y1ShowAxis=" & IIf(.Format.Line.Visible = msoTrue, "True", "False") & ";"
            themeSettings = themeSettings & "Y1ShowGridLines=" & .HasMajorGridlines & ";"
            If .HasTitle Then
                themeSettings = themeSettings & "Y1HasTitle=True;"
                themeSettings = themeSettings & "Y1Title=" & .AxisTitle.Text & ";"
                themeSettings = themeSettings & "YAxisTitleFontSize=" & .AxisTitle.Font.Size & ";"
                themeSettings = themeSettings & "YAxisTitleFontBold=" & .AxisTitle.Font.Bold & ";"
            Else
                themeSettings = themeSettings & "Y1HasTitle=False;"
            End If
            themeSettings = themeSettings & "YAxisTicksFontSize=" & .TickLabels.Font.Size & ";"
            themeSettings = themeSettings & "Y1LabelNumberFormat=" & .TickLabels.NumberFormat & ";"
        End With

        ' Legend
        If .HasLegend Then
            themeSettings = themeSettings & "HasLegend=True;"
            themeSettings = themeSettings & "LegendPosition=" & .Legend.Position & ";"
            themeSettings = themeSettings & "LegendFontSize=" & .Legend.Font.Size & ";"
            themeSettings = themeSettings & "LegendInLayout=" & .Legend.IncludeInLayout & ";"
        Else
            themeSettings = themeSettings & "HasLegend=False;"
        End If

        ' Plot Area
        If .PlotArea.Format.Fill.Visible = msoTrue Then
            themeSettings = themeSettings & "PlotAreaBackColor=" & .PlotArea.Format.Fill.ForeColor.RGB & ";"
        End If
        themeSettings = themeSettings & "PlotAreaLineWeight=" & .PlotArea.Format.Line.Weight & ";"

        ' Series formatting (taking the first series as representative)
        If .SeriesCollection.Count > 0 Then
            Dim ser As series
            Set ser = .SeriesCollection(1)
            themeSettings = themeSettings & "SeriesLineWeight=" & ser.Format.Line.Weight & ";"
            themeSettings = themeSettings & "SeriesLineDashStyle=" & ser.Format.Line.DashStyle & ";"
            themeSettings = themeSettings & "SeriesMarkerStyle=" & ser.MarkerStyle & ";"
        End If
    End With

    ' Define file path. Themes are stored relative to the workbook.
    If ThisWorkbook.Path <> "" Then
        themeFolderPath = Me.prvCombinePaths(ThisWorkbook.Path, "FALCore/Themes")
    Else
        ' Fallback if workbook is not saved (e.g., new file)
        HandleError "Please save the workbook before saving a theme."
        SaveThemeFromChart = False
        Exit Function
    End If

    themeFilePath = Me.prvCombinePaths(themeFolderPath, themeName & ".theme")

    ' Save the theme file
    If Me.prvSaveStringToFile(themeSettings, themeFilePath) Then
        SaveThemeFromChart = True
    Else
        HandleError "Failed to save the theme file to: " & themeFilePath
        SaveThemeFromChart = False
    End If

    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in SaveThemeFromChart: " & Err.Description
    SaveThemeFromChart = False
End Function

Public Function ColorSeries(Optional ByVal nbBeforeChange As Long = 1, Optional ByVal nbSeriesByGroup As Integer = -1, Optional ByVal Style As String = "DefaultStyle") As Boolean
    ' @brief Applies a color palette to the chart's series using RGB colors and intelligent application based on chart type.
    ' @param nbBeforeChange The number of series to color with the same color before moving to the next color in the palette.
    ' @param nbSeriesByGroup The number of series in a "group". The color palette will reset for each new group. -1 to treat all series as one group.
    ' @param Style The name of the color palette to use (defined in the `ColorPalette` helper function).
    If m_Chart Is Nothing Then Exit Function

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Dim ser As series
    Dim aColor As Variant
    Dim lbColor As Long, ubColor As Long, iColor As Long
    Dim seriesCounterInGroup As Long, seriesCounterBeforeChange As Long

    aColor = Me.ColorPalette(Style)
    If IsEmpty(aColor) Then
        HandleError "Color style '" & Style & "' not found or palette is empty."
        ColorSeries = False
        Exit Function
    End If

    lbColor = LBound(aColor)
    ubColor = UBound(aColor)
    iColor = lbColor

    Dim effectiveNbSeriesByGroup As Long
    If nbSeriesByGroup = -1 Then
        effectiveNbSeriesByGroup = m_Chart.FullSeriesCollection.Count
    ElseIf nbSeriesByGroup <= 0 Then
        HandleError "The parameter 'nbSeriesByGroup' must be -1 or a positive number."
        ColorSeries = False
        Exit Function
    Else
        effectiveNbSeriesByGroup = nbSeriesByGroup
    End If

    seriesCounterInGroup = 0
    seriesCounterBeforeChange = 0

    For Each ser In m_Chart.FullSeriesCollection

        ' Apply color intelligently based on the type of chart series.
        ' This ensures that the most prominent visual element (line, fill, marker) is colored.
        Select Case ser.ChartType
            Case xlLine, xlLineMarkers, xlLineMarkersStacked, xlLineMarkersStacked100, xlLineStacked, xlLineStacked100
                ser.Format.Line.ForeColor.RGB = aColor(iColor)

            Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
                ser.Format.Line.ForeColor.RGB = aColor(iColor)
                ser.MarkerForegroundColor = aColor(iColor)
                ser.MarkerBackgroundColor = aColor(iColor)

            Case xlArea, xlAreaStacked, xlAreaStacked100, _
                 xlBarClustered, xlBarOfPie, xlBarStacked, xlBarStacked100, _
                 xlColumnClustered, xlColumnStacked, xlColumnStacked100
                ser.Format.Fill.ForeColor.RGB = aColor(iColor)

            Case xlPie, xlPieOfPie
                ' Note: For Pie charts, this colors the entire series.
                ' For individual slice colors, a different approach is needed.
                ser.Format.Fill.ForeColor.RGB = aColor(iColor)

            Case Else
                ' Default to border color for unknown types
                ser.Border.Color = aColor(iColor)
        End Select

        ' Increment color index based on grouping logic
        seriesCounterInGroup = seriesCounterInGroup + 1
        seriesCounterBeforeChange = seriesCounterBeforeChange + 1

        If seriesCounterBeforeChange >= nbBeforeChange Then
            iColor = iColor + 1
            seriesCounterBeforeChange = 0
        End If

        If seriesCounterInGroup >= effectiveNbSeriesByGroup Then
            iColor = lbColor
            seriesCounterInGroup = 0
        ElseIf iColor > ubColor Then
            iColor = lbColor
        End If
    Next ser

    Application.ScreenUpdating = True
    ColorSeries = True
    Exit Function

ErrHandler:
    Application.ScreenUpdating = True
    HandleError "Error while coloring series: " & Err.Description
    ColorSeries = False
End Function

Public Function FormatTextSubscripts() As Boolean
    ' @brief Formats predefined terms in chart titles as subscripts.
    ' @details Searches for specific strings (e.g., "DS,sat", "max") in the main title and axis titles and applies subscript formatting. This is useful for scientific and engineering charts.
    ' @return True if the operation was successful.
    If m_Chart Is Nothing Then Exit Function

    Dim StrToFormat As String
    Dim toBaselineOffset As String
    Dim indiceTerms As Variant
    Dim index As Long, textLenght As Long, textDebut As Long
    Dim axNo As Axis

    indiceTerms = Array("DS,sat", "DSS", "D,max", "G,inv", "ON ", "OFF ", "th ", "th,lin", "DS ", "GS ", "m ", "m,max ", "B ", "GS,ref", "DS,ref", "sat", "max", "min", ".abs")
    With m_Chart
        For index = 0 To UBound(indiceTerms)
            textLenght = Len(indiceTerms(index))

            If .HasTitle Then
                textDebut = InStr(1, .chartTitle.Text, indiceTerms(index), vbTextCompare)
                If textDebut Then .chartTitle.Format.TextFrame2.TextRange.Characters(textDebut, textLenght).Font.BaselineOffset = -0.25
            End If

            For Each axNo In .Axes
                If axNo.HasTitle Then
                    textDebut = InStr(1, axNo.AxisTitle.Text, indiceTerms(index), vbTextCompare)
                    If textDebut Then axNo.AxisTitle.Format.TextFrame2.TextRange.Characters(textDebut, textLenght).Font.BaselineOffset = -0.25
                End If
            Next
        Next
    End With
    FormatTextSubscripts = True
End Function

Public Function ResizeFonts(Optional fontSize As Double = 14) As Boolean
    ' @brief Resizes all fonts in a chart to a specified size.
    On Error GoTo ErrHandler
    If m_Chart Is Nothing Then
        HandleError "No chart has been set to resize fonts."
        ResizeFonts = False
        Exit Function
    End If

    m_Chart.ChartArea.format.TextFrame2.TextRange.Font.Size = fontSize
    ResizeFonts = True
    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in ResizeFonts."
    ResizeFonts = False
End Function

'--- Other Public Methods ---

Public Function ExportAsImage(Optional ByVal imagePath As String = "", Optional ByVal imageFilterIndex As Long = 1) As Boolean
    ' @brief Exports the chart object to an image file.
    ' @param imagePath (Optional) The full path to save the image. If omitted, a "Save As" dialog is shown.
    ' @param imageFilterIndex (Optional) The default file type for the dialog (1=PNG, 2=JPG, etc.).
    ' @return True if the export was successful, False otherwise.
    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then
        HandleError "No chart has been set to export."
        Exit Function
    End If

    Dim defaultFileName As String
    Dim fileExtension As String
    Dim filters As String
    Dim selectedFile As Variant

    ExportAsImage = False ' Default to failure

    If imagePath = "" Then
        filters = "PNG File (*.png),*.png," & _
                  "JPEG File (*.jpg),*.jpg," & _
                  "GIF File (*.gif),*.gif," & _
                  "BMP File (*.bmp),*.bmp"

        defaultFileName = Environ("USERPROFILE") & "\Desktop\" & Me.SanitizeFileName(m_Chart.Name) & ".png"

        With Application.FileDialog(msoFileDialogSaveAs)
            .InitialFileName = defaultFileName
            .Title = "Save Chart As..."
            .filters.Clear
            .filters.Add "Images", filters, 1
            .FilterIndex = imageFilterIndex

            If .Show = -1 Then
                selectedFile = .SelectedItems(1)
            Else
                Exit Function ' User cancelled
            End If
        End With
        imagePath = selectedFile
    End If

    If InStr(imagePath, ".") = 0 Then
        Select Case imageFilterIndex
            Case 1: fileExtension = ".png"
            Case 2: fileExtension = ".jpg"
            Case 3: fileExtension = ".gif"
            Case 4: fileExtension = ".bmp"
            Case Else: fileExtension = ".png"
        End Select
        imagePath = imagePath & fileExtension
    End If

    Dim picType As XlChartPictureType
    Select Case LCase(Right(imagePath, 3))
        Case "png": picType = xlPNG
        Case "jpg": picType = xlJPEG
        Case "gif": picType = xlGIF
        Case "bmp": picType = xlBitmap
        Case Else: picType = xlPNG
    End Select

    m_Chart.Export Filename:=imagePath, FilterName:=picType

    ExportAsImage = True
    Exit Function

ErrHandler:
    HandleError "Error while exporting chart to image: " & Err.Description
End Function

Public Sub SetChartFromSelection()
    ' @brief Sets the class's internal chart object from the currently selected chart.
    ' @details This procedure allows the user to easily assign the active chart to the class instance for manipulation.
    On Error GoTo ErrHandler

    If Not ActiveChart Is Nothing Then
        Set Me.Chart = ActiveChart
    Else
        HandleError "No chart is currently selected. Please select a chart and try again."
    End If
    Exit Sub

ErrHandler:
    HandleError "An unexpected error occurred while setting the chart from selection."
End Sub

Public Function CopyChart(myChart As Chart, Optional cpyPosition As String = "Right", Optional offsetPosition As Double = 0) As Variant
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

    Set CopyChart = newChart
    Exit Function
ifError:
    CopyChart = CVErr(2001)
End Function

Public Function DeleteUnmatchingSeries(sMatch As String, Optional compareAbsolute As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim i As Integer

    If m_Chart Is Nothing Then
        HandleError "No chart has been set to delete series from."
        DeleteUnmatchingSeries = False
        Exit Function
    End If

    With m_Chart
        For i = 1 To .FullSeriesCollection.count
            If compareAbsolute Then
                If .FullSeriesCollection(i).name <> sMatch Then .FullSeriesCollection(i).Delete: i = i - 1
            Else
                If InStr(.FullSeriesCollection(i).name, sMatch) = 0 Then .FullSeriesCollection(i).Delete: i = i - 1
            End If
            If i + 1 > .FullSeriesCollection.count Then Exit For
        Next i
    End With

    DeleteUnmatchingSeries = True
    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in DeleteUnmatchingSeries."
    DeleteUnmatchingSeries = False
End Function

Public Function FormatAsSmithChart(Optional AddBackGround As Boolean = True, Optional FormattingOptions As String = "") As Boolean
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Dim cSerie1         As Variant
    Dim cSerie2         As Variant
    Dim picturePath     As String
    Dim isSmithData     As Boolean
    Dim fmtOptions      As String

    If m_Chart Is Nothing Then
        HandleError "No chart has been set to format as a Smith Chart."
        FormatAsSmithChart = False
        Exit Function
    End If

    With m_Chart
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
        If Not isSmithData Then
            HandleError "The chart does not appear to contain valid Smith Chart data (R:/I: series pairs)."
            FormatAsSmithChart = False
            Exit Function
        End If

        fmtOptions = "Y1Min=-1;Y1Max=1;X1Min=-1;X1Max=1;HasTitle=false;HasLegend=false;SquarePlot=1;X1HasTitle=false;Y1HasTitle=false;X1ShowGridLines=false;Y1ShowGridLines=false;X1ShowAxis=false;Y1ShowAxis=false" & _
                        ";X1ScaleType=" & xlLinear & ";Y1ScaleType=" & xlLinear
        Me.FormatChart fmtOptions & ";" & FormattingOptions

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
            If Me.prvFileExists(picturePath) Then
                .PlotArea.format.Fill.Visible = msoTrue
                .PlotArea.format.Fill.UserPicture picturePath
                .PlotArea.format.Fill.TextureTile = msoFalse
                .PlotArea.format.Fill.Transparency = 0.6
            End If
        Else
            .PlotArea.format.Fill.Visible = msoFalse
        End If
    End With

    Application.ScreenUpdating = True
    FormatAsSmithChart = True
    Exit Function

ErrHandler:
    Application.ScreenUpdating = True
    HandleError "An unexpected error occurred in FormatAsSmithChart."
    FormatAsSmithChart = False
End Function

Public Function UnlinkDataFromSource() As Boolean
    On Error GoTo ifError
    Dim cSerie          As Variant
    For Each cSerie In m_Chart.FullSeriesCollection
        cSerie.XValues = cSerie.XValues
        cSerie.values = cSerie.values
        cSerie.name = cSerie.name
    Next
    UnlinkDataFromSource = True
    Exit Function
ifError:
    UnlinkDataFromSource = False
End Function

Public Function ApplyYLogarithmicScale() As Boolean
    On Error GoTo ErrHandler
    If m_Chart Is Nothing Then
        HandleError "No chart has been set to apply a log scale."
        ApplyYLogarithmicScale = False
        Exit Function
    End If

    Me.ApplyYAxisAbsoluteValue
    m_Chart.Axes(xlValue, xlPrimary).ScaleType = xlLogarithmic
    ApplyYLogarithmicScale = True
    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in ApplyYLogarithmicScale."
    ApplyYLogarithmicScale = False
End Function

Public Function ApplyYAxisAbsoluteValue() As Boolean
    On Error GoTo ErrHandler
    Dim cSerie As Variant

    If m_Chart Is Nothing Then
        HandleError "No chart has been set to apply absolute values."
        ApplyYAxisAbsoluteValue = False
        Exit Function
    End If

    For Each cSerie In m_Chart.FullSeriesCollection
        cSerie.values = FalArray.a1D_math_Abs(cSerie.values)
    Next
    ApplyYAxisAbsoluteValue = True
    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in ApplyYAxisAbsoluteValue."
    ApplyYAxisAbsoluteValue = False
End Function

Public Function CalculateDerivative(Optional XSample As Long = 3) As Boolean
    On Error GoTo ErrHandler
    Dim cSerie          As Variant
    Dim a2D_X           As Variant
    Dim a2D_Y           As Variant
    Dim a2D_XY          As Variant
    Dim a1D_dY          As Variant

    If m_Chart Is Nothing Then
        HandleError "No chart has been set to calculate derivatives."
        CalculateDerivative = False
        Exit Function
    End If

    With m_Chart
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
    CalculateDerivative = True
    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in CalculateDerivative."
    CalculateDerivative = False
End Function

Public Function FilterSeries(sMatch As String, Optional compareAbsolute As Boolean = False) As Boolean
    ' @brief Filters series in a chart based on a match condition.
    On Error GoTo ErrHandler
    Dim i As Integer

    If m_Chart Is Nothing Then
        HandleError "No chart has been set to filter series."
        FilterSeries = False
        Exit Function
    End If

    With m_Chart
        For i = 1 To .FullSeriesCollection.count
            If compareAbsolute Then
                .FullSeriesCollection(i).IsFiltered = IIf(.FullSeriesCollection(i).name = sMatch, False, True)
            Else
                .FullSeriesCollection(i).IsFiltered = IIf(InStr(.FullSeriesCollection(i).name, sMatch) > 0, False, True)
            End If
        Next i
    End With
    FilterSeries = True
    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in FilterSeries."
    FilterSeries = False
End Function

Public Function DeleteMatchingSeries(sMatch As String, Optional compareAbsolute As Boolean = False) As Boolean
    ' @brief Deletes series in a chart based on a match condition.
    On Error GoTo ErrHandler
    Dim i As Integer

    If m_Chart Is Nothing Then
        HandleError "No chart has been set to delete series from."
        DeleteMatchingSeries = False
        Exit Function
    End If

    With m_Chart
        For i = 1 To .FullSeriesCollection.count
            If compareAbsolute Then
                If .FullSeriesCollection(i).name = sMatch Then .FullSeriesCollection(i).Delete: i = i - 1
            Else
                If InStr(.FullSeriesCollection(i).name, sMatch) > 0 Then .FullSeriesCollection(i).Delete: i = i - 1
            End If
            If i + 1 > .FullSeriesCollection.count Then Exit For
        Next i
    End With

    DeleteMatchingSeries = True
    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in DeleteMatchingSeries."
    DeleteMatchingSeries = False
End Function

Public Function ClearSeries() As Boolean
    ' @brief Removes all series from the chart.
    If m_Chart Is Nothing Then Exit Function

    On Error GoTo ifError
    Application.ScreenUpdating = False
    ' Ensure the chart has existing series
    If m_Chart.SeriesCollection.Count > 0 Then
        ' Delete existing series
        Do While m_Chart.SeriesCollection.Count > 0
            m_Chart.SeriesCollection(1).Delete
        Loop
    End If
    If m_Chart.SeriesCollection.Count = 0 Then ClearSeries = True Else ClearSeries = False
    Application.ScreenUpdating = True
    Exit Function
ifError:
    Application.ScreenUpdating = True
    ClearSeries = False
End Function

Public Function SetSourceData(ByVal rangeData As Range, Optional ByVal PlotBy As XlRowCol = xlColumns) As Boolean
    ' @brief Sets the source data for the chart from a range.
    If m_Chart Is Nothing Then Exit Function
    If TypeName(rangeData) <> "Range" Then SetSourceData = False: Exit Function

    On Error GoTo ifError
    m_Chart.SetSourceData Source:=rangeData
    m_Chart.PlotBy = PlotBy
    SetSourceData = True
    Exit Function
ifError:
    SetSourceData = False
End Function

Public Function AddSeriesFromRange(ByVal data_src As Range) As Boolean
    ' @brief Adds a new series to the chart from a data range.
    If m_Chart Is Nothing Then Exit Function

    On Error GoTo ifError
    Application.ScreenUpdating = False
    Dim seriesRange As Range
    Dim newSeries As series
    Dim i As Integer

    ' Add a series for each column in the data source
    For i = 2 To data_src.Columns.Count ' Start from column 2 to skip the first column (XValues)
        If m_Chart.FullSeriesCollection.Count >= 255 Then GoTo ifError  ' Max Series exceeded

        If data_src.Cells(1, i).value <> data_src.Cells(1, 1).value Then
            Set newSeries = m_Chart.SeriesCollection.newSeries

            ' Set the series values
            Set seriesRange = data_src.Columns(i).Offset(1, 0).Resize(data_src.Rows.Count - 1, 1)
            newSeries.values = seriesRange

            ' Set the X-axis values
            Set seriesRange = data_src.Columns(1).Offset(1, 0).Resize(data_src.Rows.Count - 1, 1)
            newSeries.XValues = seriesRange

            ' Set series names
            newSeries.name = "='" & data_src.Parent.name & "'!" & data_src.Cells(1, i).Address
        End If
    Next i
    Application.ScreenUpdating = True
    AddSeriesFromRange = True
    Exit Function
ifError:
    Application.ScreenUpdating = True
    AddSeriesFromRange = False
End Function

Public Function AddSeriesFromArray(ByVal Arr2D As Variant) As Boolean
    ' @brief Adds a new series to the chart from a 2D array.
    If m_Chart Is Nothing Then Exit Function

    On Error GoTo ifError
    Dim newSeries As series
    Dim i As Integer

    ' Add a series for each element in the array
    For i = LBound(Arr2D, 2) + 1 To UBound(Arr2D, 2)
        If m_Chart.FullSeriesCollection.Count >= 255 Then Exit Function  ' Max Series exceeded

        Set newSeries = m_Chart.SeriesCollection.newSeries

        ' Set the series values
        newSeries.values = Application.Index(Arr2D, 0, i)

        ' Set the X-axis values (first column of the array)
        newSeries.XValues = Application.Index(Arr2D, 0, 1)

        ' Set series names
        newSeries.name = Arr2D(LBound(Arr2D, 1), i)
    Next i
    AddSeriesFromArray = True
    Exit Function
ifError:
    AddSeriesFromArray = False
End Function

Public Function AddTrendline(ByVal seriesIndex As Integer, Optional ByVal trendlineType As XlTrendlineType = xlLinear, Optional ByVal order As Integer = 2, Optional ByVal forecastForward As Double = 0, Optional ByVal forecastBackward As Double = 0, Optional ByVal displayEquation As Boolean = False, Optional ByVal displayRSquared As Boolean = False) As Boolean
    ' @brief Adds a trendline to a specified series on the chart.
    ' @param seriesIndex The index number of the series to add the trendline to.
    ' @param trendlineType (Optional) The type of trendline to add (e.g., xlLinear, xlPolynomial).Defaults to xlLinear.
    ' @param order (Optional) The polynomial order if the trendlineType is xlPolynomial. Defaults to 2.
    ' @param forecastForward (Optional) The number of periods to forecast forward. Defaults to 0.
    ' @param forecastBackward (Optional) The number of periods to forecast backward. Defaults to 0.
    ' @param displayEquation (Optional) True to display the equation on the chart. Defaults to False.
    ' @param displayRSquared (Optional) True to display the R-squared value on the chart. Defaults to False.
    ' @return True if the trendline was added successfully, False otherwise.
    ' @example
    ' Dim plot As New FalPlot
    ' Set plot.Chart = ActiveChart
    ' plot.AddTrendline seriesIndex:=1, trendlineType:=xlPolynomial, order:=3, displayEquation:=True

    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then GoTo ErrHandler
    If seriesIndex < 1 Or seriesIndex > m_Chart.SeriesCollection.Count Then GoTo ErrHandler

    Dim tl As Trendline
    Set tl = m_Chart.SeriesCollection(seriesIndex).Trendlines.Add(Type:=trendlineType, Order:=order)

    With tl
        .Forward = forecastForward
        .Backward = forecastBackward
        .DisplayEquation = displayEquation
        .DisplayRSquared = displayRSquared
    End With

    AddTrendline = True
    Exit Function

ErrHandler:
    AddTrendline = False
End Function

Public Function AddErrorBars(ByVal seriesIndex As Integer, Optional ByVal direction As XlErrorBarDirection = xlY, Optional ByVal include As XlErrorBarInclude = xlErrorBarIncludeBoth, Optional ByVal type As XlErrorBarType = xlErrorBarTypeStDev, Optional ByVal amount As Double = 1) As Boolean
    ' @brief Adds error bars to a specified series on the chart.
    ' @param seriesIndex The index number of the series to add the error bars to.
    ' @param direction (Optional) The direction of the error bars (xlX or xlY). Defaults to xlY.
    ' @param include (Optional) The parts of the error bar to include (e.g., xlErrorBarIncludeBoth). Defaults to xlErrorBarIncludeBoth.
    ' @param type (Optional) The type of error bars (e.g., xlErrorBarTypeStDev, xlErrorBarTypeFixedValue). Defaults to xlErrorBarTypeStDev.
    ' @param amount (Optional) The amount for the error bars, used for fixed value, percentage, and standard deviation types. Defaults to 1.
    ' @return True if the error bars were added successfully, False otherwise.
    ' @example
    ' Dim plot As New FalPlot
    ' Set plot.Chart = ActiveChart
    ' plot.AddErrorBars seriesIndex:=1, type:=xlErrorBarTypeFixedValue, amount:=5

    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then GoTo ErrHandler
    If seriesIndex < 1 Or seriesIndex > m_Chart.SeriesCollection.Count Then GoTo ErrHandler

    Dim ser As series
    Set ser = m_Chart.SeriesCollection(seriesIndex)

    '--- This is the fix: Check if the series chart type supports error bars ---
    Select Case ser.ChartType
        Case xlArea, xlAreaStacked, xlAreaStacked100, _
             xlBarClustered, xlBarOfPie, xlBarStacked, xlBarStacked100, _
             xlColumnClustered, xlColumnStacked, xlColumnStacked100, _
             xlLine, xlLineMarkers, xlLineMarkersStacked, xlLineMarkersStacked100, xlLineStacked, xlLineStacked100, _
             xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            ' This chart type supports error bars, proceed.
        Case Else
            ' This chart type does not support error bars.
            HandleError "The chart type for the selected series does not support error bars."
            AddErrorBars = False
            Exit Function
    End Select

    ser.ErrorBar Direction:=direction, Include:=include, Type:=type, Amount:=amount

    AddErrorBars = True
    Exit Function

ErrHandler:
    AddErrorBars = False
    HandleError "Failed to add error bars. " & Err.Description
End Function

'--- Private Helper Functions ---

Private Function CreateChartFromRange(ByVal data As Range, Optional ByVal chartType As XlChartType = xlXYScatter, Optional ByVal chartName As String = "") As Chart
    ' @brief Helper to create and format a chart from a range.
    ' @details This now acts as a wrapper for the public CreateChart method.
    On Error GoTo ifError

    Dim tempName As String
    Dim newChart As Chart

    If chartName = "" Then
        tempName = "Chart_" & data.Parent.Name
    Else
        tempName = chartName
    End If

    ' Call the new, powerful CreateChart method, which also handles formatting
    Set newChart = Me.CreateChart(tempName, data.Parent.Parent, chartType, data)
    Set CreateChartFromRange = newChart

    Exit Function
ifError:
    Set CreateChartFromRange = Nothing
End Function

Private Function LoadTheme(ByVal themeName As String) As String
    ' @brief Loads theme settings from a .theme file.
    ' @param themeName The name of the theme to load.
    ' @return A string containing the theme settings, or an empty string on failure.
    On Error GoTo ErrHandler

    Dim themeFolderPath As String
    Dim themeFilePath As String
    Dim themeSettings As String

    ' Define file path
    If ThisWorkbook.Path <> "" Then
        themeFolderPath = Me.prvCombinePaths(ThisWorkbook.Path, "FALCore/Themes")
    Else
        HandleError "Please save the workbook before loading a theme."
        LoadTheme = ""
        Exit Function
    End If

    themeFilePath = Me.prvCombinePaths(themeFolderPath, themeName & ".theme")

    If Me.prvFileExists(themeFilePath) Then
        themeSettings = Me.prvReadFile(themeFilePath)
        LoadTheme = themeSettings
    Else
        HandleError "Theme file not found: " & themeFilePath
        LoadTheme = ""
    End If

    Exit Function

ErrHandler:
    HandleError "An unexpected error occurred in LoadTheme: " & Err.Description
    LoadTheme = ""
End Function

Private Sub HandleError(ByVal ErrMsg As String)
    MsgBox "An error occurred:" & vbCrLf & ErrMsg, vbOKOnly + vbCritical, "Operation impossible"
End Sub

Private Function prvRemoveChars(ByVal text As String, ByVal charsToRemove As String) As String
    ' @brief Removes specified characters from a string.
    Dim i As Long
    prvRemoveChars = text
    For i = 1 To Len(charsToRemove)
        prvRemoveChars = Replace(prvRemoveChars, Mid(charsToRemove, i, 1), "")
    Next i
End Function

Private Function prvResizeString(ByVal text As String, ByVal maxLength As Long) As String
    ' @brief Truncates a string to a specified maximum length.
    If Len(text) > maxLength Then
        prvResizeString = Left(text, maxLength)
    Else
        prvResizeString = text
    End If
End Function

Private Function prvRemoveSubstrings(ByVal text As String, ByVal substringsToRemove As String) As String
    ' @brief Removes specified substrings from a string. Substrings are semicolon-delimited.
    Dim i As Long
    Dim subs As Variant
    prvRemoveSubstrings = text
    subs = Split(substringsToRemove, ";")
    For i = LBound(subs) To UBound(subs)
        prvRemoveSubstrings = Replace(prvRemoveSubstrings, subs(i), "")
    Next i
End Function

Private Function prvColumnNumberToLetter(ByVal colIndex As Long) As String
    ' @brief Converts a column number to its Excel letter representation.
    Dim s As String
    Do While colIndex > 0
        s = Chr$(((colIndex - 1) Mod 26) + 65) & s
        colIndex = (colIndex - 26) \ 26
    Loop
    prvColumnNumberToLetter = s
End Function

' --- File Operations ---
' These functions replicate the functionality of the FalFile module,
' removing the external dependency and making the class more self-contained.
' They use the Scripting.FileSystemObject for file system interactions.

Private Function prvCombinePaths(ByVal p1 As String, ByVal p2 As String) As String
    ' @brief Combines two path parts, handling the path separator.
    If Right(p1, 1) = "\" Then
        prvCombinePaths = p1 & p2
    Else
        prvCombinePaths = p1 & "\" & p2
    End If
End Function

Private Function prvFileExists(ByVal filePath As String) As Boolean
    ' @brief Checks if a file exists.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    prvFileExists = fso.FileExists(filePath)
    Set fso = Nothing
End Function

Private Function prvReadFile(ByVal filePath As String) As String
    ' @brief Reads the entire content of a text file.
    Dim fso As Object
    Dim fileStream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileStream = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    prvReadFile = fileStream.ReadAll
    fileStream.Close
    Set fso = Nothing
End Function

Private Function prvSaveStringToFile(ByVal text As String, ByVal filePath As String) As Boolean
    ' @brief Saves a string to a text file, overwriting if it exists.
    On Error GoTo SaveErr
    Dim fso As Object
    Dim fileStream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ensure the directory exists
    Dim parentFolder As String
    parentFolder = fso.GetParentFolderName(filePath)
    If Not fso.FolderExists(parentFolder) Then
        fso.CreateFolder parentFolder
    End If

    Set fileStream = fso.CreateTextFile(filePath, True) ' True = Overwrite
    fileStream.Write text
    fileStream.Close
    Set fso = Nothing
    prvSaveStringToFile = True
    Exit Function
SaveErr:
    prvSaveStringToFile = False
End Function

Private Function SanitizeFileName(ByVal fileName As String) As String
    ' @brief Replaces characters that are invalid in file names with underscores.
    Dim invalidChars As Variant
    Dim i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SanitizeFileName = fileName
    For i = LBound(invalidChars) To UBound(invalidChars)
        SanitizeFileName = Replace(SanitizeFileName, invalidChars(i), "_")
    Next i
    SanitizeFileName = Replace(SanitizeFileName, " ", "_")
End Function

Private Function ColorPalette(Optional ByVal Style As String = "DefaultStyle") As Variant
    ' @brief Returns a predefined array of RGB color values for chart series.
    ' @param Style The name of the color palette to return.
    Select Case Style
        Case "Vibrant"
            ColorPalette = Array(RGB(27, 158, 119), RGB(217, 95, 2), RGB(117, 112, 179), RGB(231, 41, 138), RGB(102, 166, 30), RGB(230, 171, 2), RGB(166, 118, 29), RGB(102, 102, 102), _
                                 RGB(23, 190, 207), RGB(188, 189, 34), RGB(255, 127, 14), RGB(148, 103, 189), RGB(140, 86, 75), RGB(227, 119, 194), RGB(127, 127, 127), RGB(44, 160, 44))
        Case "Rainbow"
            ColorPalette = Array(RGB(228, 26, 28), RGB(247, 129, 48), RGB(253, 191, 111), RGB(255, 255, 179), RGB(166, 216, 84), RGB(51, 160, 44), RGB(31, 120, 180), RGB(178, 223, 138), _
                                 RGB(152, 78, 163), RGB(251, 154, 153), RGB(106, 61, 154), RGB(255, 127, 0), RGB(177, 89, 40), RGB(202, 178, 214), RGB(102, 194, 165), RGB(252, 141, 98))
        Case "Grayscale"
            ColorPalette = Array(RGB(0, 0, 0), RGB(32, 32, 32), RGB(64, 64, 64), RGB(96, 96, 96), RGB(128, 128, 128), RGB(160, 160, 160), RGB(192, 192, 192), RGB(224, 224, 224))
        Case Else ' DefaultStyle / Tableau
            ColorPalette = Array(RGB(31, 119, 180), RGB(255, 127, 14), RGB(44, 160, 44), RGB(214, 39, 40), RGB(148, 103, 189), RGB(140, 86, 75), RGB(227, 119, 194), RGB(127, 127, 127), _
                                 RGB(188, 189, 34), RGB(23, 190, 207), RGB(174, 199, 232), RGB(255, 187, 120), RGB(152, 223, 138), RGB(255, 152, 150), RGB(197, 176, 213), RGB(196, 156, 148))
    End Select
End Function

Private Function FindFirstRowFromBottom(ByVal ws As Worksheet, ByVal col As Long, ByVal startRow As Long) As Long
    ' @brief Finds the first non-empty row in a column, searching upwards from a starting row.
    On Error GoTo ifError

    Dim r As Long
    For r = startRow To 1 Step -1
        If Not IsEmpty(ws.Cells(r, col).value) Then
            FindFirstRowFromBottom = r
            Exit Function
        End If
    Next r

    FindFirstRowFromBottom = 1 ' Fallback
    Exit Function
ifError:
    FindFirstRowFromBottom = 1 ' Fallback on error
End Function
