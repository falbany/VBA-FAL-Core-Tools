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
' Dim plot As New FalPlot
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

'--- Private Methods ---

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
        themeFolderPath = FalFile.Combine_Paths(ThisWorkbook.Path, "FALCore/Themes")
    Else
        Me.HandleError "Please save the workbook before loading a theme."
        LoadTheme = ""
        Exit Function
    End If

    themeFilePath = FalFile.Combine_Paths(themeFolderPath, themeName & ".theme")

    If FalFile.FileExist(themeFilePath) Then
        themeSettings = FalFile.ReadFile(themeFilePath)
        LoadTheme = themeSettings
    Else
        Me.HandleError "Theme file not found: " & themeFilePath
        LoadTheme = ""
    End If

    Exit Function

ErrHandler:
    Me.HandleError "An unexpected error occurred in LoadTheme: " & Err.Description
    LoadTheme = ""
End Function

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
End Sub

'--- Public Methods ---

Public Sub FormatChart()
    ' @brief Applies the current formatting properties to the chart.
    ' @details This method is a simple wrapper around the more powerful ApplyFormattingFromString method.
    '          It uses the current state of the class properties to format the chart.
    Me.ApplyFormattingFromString
End Sub

Public Sub CreateSmithChart()
    ' @brief Creates a Smith chart from the currently selected chart.
    On Error GoTo ErrHandler

    If ActiveChart Is Nothing Then
        Me.HandleError "Please select a chart first."
        Exit Sub
    End If

    Set m_Chart = Me.CopyChart(ActiveChart)
    If m_Chart Is Nothing Then
        Me.HandleError "Failed to copy the chart."
        Exit Sub
    End If

    Me.DeleteUnmatchingSeries ":", False
    Me.FormatAsSmithChart True
    Me.DelinkData

    Exit Sub
ErrHandler:
    Me.HandleError "An unexpected error occurred in CreateSmithChart."
End Sub

Public Sub CreateYLog()
    ' @brief Creates a copy of the selected chart with a logarithmic Y-axis.
    On Error GoTo ErrHandler

    If ActiveChart Is Nothing Then
        Me.HandleError "Please select a chart first."
        Exit Sub
    End If

    Set m_Chart = Me.CopyChart(ActiveChart)
    If m_Chart Is Nothing Then
        Me.HandleError "Failed to copy the chart."
        Exit Sub
    End If

    Me.ApplyYLogScale
    Me.DelinkData

    Exit Sub
ErrHandler:
    Me.HandleError "An unexpected error occurred in CreateYLog."
End Sub

Public Sub CreateDerivative()
    ' @brief Creates a copy of the selected chart with its series' derivatives.
    On Error GoTo ErrHandler

    If ActiveChart Is Nothing Then
        Me.HandleError "Please select a chart first."
        Exit Sub
    End If

    Set m_Chart = Me.CopyChart(ActiveChart)
    If m_Chart Is Nothing Then
        Me.HandleError "Failed to copy the chart."
        Exit Sub
    End If

    Me.CalculateDerivative
    Me.DelinkData

    Exit Sub
ErrHandler:
    Me.HandleError "An unexpected error occurred in CreateDerivative."
End Sub

Public Function ExportAsImage(Optional ByVal imagePath As String = "", Optional ByVal imageFilterIndex As Long = 1) As Boolean
    ' @brief Exports the chart object to an image file.
    ' @param imagePath (Optional) The full path to save the image. If omitted, a "Save As" dialog is shown.
    ' @param imageFilterIndex (Optional) The default file type for the dialog (1=PNG, 2=JPG, etc.).
    ' @return True if the export was successful, False otherwise.
    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then
        Me.HandleError "No chart has been set to export."
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
    Me.HandleError "Error while exporting chart to image: " & Err.Description
End Function

Private Function HandleError(ByVal ErrMsg As String)
    MsgBox "An error occurred:" & vbCrLf & ErrMsg, vbOKOnly + vbCritical, "Operation impossible"
End Function

Public Function GetSelectedChart() As Chart
    ' @brief Returns the currently selected chart in the application.
    ' @details This is a helper function to get a reference to the active chart.
    '          It is not tied to the class instance's m_Chart property.
    ' @return A Chart object if a chart is selected, otherwise Nothing.
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
    Dim i       As Integer

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

End Function

Public Function FormatAsSmithChart(Optional AddBackGround As Boolean = True, Optional FormattingOptions As String = "") As Boolean
    Dim cSerie1         As Variant
    Dim cSerie2         As Variant
    Dim picturePath     As String
    Dim isSmithData     As Boolean
    Dim fmtOptions      As String

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
        If Not isSmithData Then FormatAsSmithChart = False: Exit Function

        fmtOptions = "Y1Min=-1;Y1Max=1;X1Min=-1;X1Max=1;HasTitle=false;HasLegend=false;SquarePlot=1;X1HasTitle=false;Y1HasTitle=false;X1ShowGridLines=false;Y1ShowGridLines=false;X1ShowAxis=false;Y1ShowAxis=false" & _
                        ";X1ScaleType=" & xlLinear & ";Y1ScaleType=" & xlLinear
        Me.ApplyFormattingFromString fmtOptions & ";" & FormattingOptions

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

    FormatAsSmithChart = True
End Function

Public Function DelinkData() As Boolean
    On Error GoTo ifError
    Dim cSerie          As Variant
    For Each cSerie In m_Chart.FullSeriesCollection
        cSerie.XValues = cSerie.XValues
        cSerie.values = cSerie.values
        cSerie.name = cSerie.name
    Next
    DelinkData = True
    Exit Function
ifError:
    DelinkData = False
End Function

Public Function ApplyYLogScale() As Boolean
    Me.ApplyYAbsoluteValue
    m_Chart.Axes(xlValue, xlPrimary).ScaleType = xlLogarithmic
    ApplyYLogScale = True
End Function

Public Function ApplyYAbsoluteValue() As Boolean
    Dim cSerie          As Variant
    For Each cSerie In m_Chart.FullSeriesCollection
        cSerie.values = FalArray.a1D_math_Abs(cSerie.values)
    Next
    ApplyYAbsoluteValue = True
End Function

Public Function CalculateDerivative(Optional XSample As Long = 3) As Boolean
    Dim cSerie          As Variant
    Dim a2D_X           As Variant
    Dim a2D_Y           As Variant
    Dim a2D_XY          As Variant
    Dim a1D_dY          As Variant

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
End Function


Public Function FilterSeries(sMatch As String, Optional compareAbsolute As Boolean = False) As Boolean
    ' @brief Filters series in a chart based on a match condition.
    Dim i       As Integer

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
End Function

Public Function DeleteMatchingSeries(sMatch As String, Optional compareAbsolute As Boolean = False) As Boolean
    ' @brief Deletes series in a chart based on a match condition.
    Dim i       As Integer

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
End Function

Public Function ResizeFonts(Optional fontSize As Double = 14) As Boolean
    ' @brief Resizes all fonts in a chart to a specified size.
    m_Chart.ChartArea.format.TextFrame2.TextRange.Font.Size = fontSize
    ResizeFonts = True
End Function

Public Function ClearSeries() As Boolean
    ' @brief Removes all series from the chart.
    If m_Chart Is Nothing Then Exit Function

    On Error GoTo ifError
    ' Ensure the chart has existing series
    If m_Chart.SeriesCollection.Count > 0 Then
        ' Delete existing series
        Do While m_Chart.SeriesCollection.Count > 0
            m_Chart.SeriesCollection(1).Delete
        Loop
    End If
    If m_Chart.SeriesCollection.Count = 0 Then ClearSeries = True Else ClearSeries = False
    Exit Function
ifError:
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
    AddSeriesFromRange = True
    Exit Function
ifError:
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

Public Function ColorSeries(Optional ByVal nbBeforeChange As Long = 1, Optional ByVal nbSeriesByGroup As Integer = -1, Optional ByVal Style As String = "DefaultStyle") As Boolean
    ' @brief Applies a color palette to the chart's series.
    If m_Chart Is Nothing Then Exit Function

    On Error GoTo ErrHandler

    Dim currentSeries As series
    Dim aColor As Variant
    Dim lbColor As Long
    Dim ubColor As Long
    Dim iColor As Long
    Dim seriesCounterInGroup As Long
    Dim seriesCounterBeforeChange As Long

    aColor = Me.ColorPalette(Style)
    If IsEmpty(aColor) Then
        Me.HandleError "Color style '" & Style & "' not found or palette is empty."
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
        Me.HandleError "The parameter 'nbSeriesByGroup' must be -1 or a positive number."
        ColorSeries = False
        Exit Function
    Else
        effectiveNbSeriesByGroup = nbSeriesByGroup
    End If

    seriesCounterInGroup = 0
    seriesCounterBeforeChange = 0

    For Each currentSeries In m_Chart.FullSeriesCollection
        currentSeries.Border.ColorIndex = aColor(iColor)

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
    Next currentSeries

    ColorSeries = True
    Exit Function

ErrHandler:
    Me.HandleError "Error while coloring series: " & Err.Description
    ColorSeries = False
End Function

Public Function SaveThemeFromChart(ByVal themeName As String) As Boolean
    ' @brief Extracts the current chart's formatting and saves it as a theme file.
    ' @param themeName The name to give the new theme (e.g., "MyCustomTheme").
    ' @return True if the theme was saved successfully, False otherwise.
    ' @dependencies FalFile.bas
    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then
        Me.HandleError "No chart is currently set. Cannot save theme."
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
            themeSettings = themeSettings & "FontName=" & .chartTitle.Font.Name & ";" ' Use title font as main font
        Else
            themeSettings = themeSettings & "HasTitle=False;"
        End If

        ' Axes
        ' Primary X-Axis
        With .Axes(xlCategory, xlPrimary)
            themeSettings = themeSettings & "X1ShowAxis=" & IIf(.Format.Line.Visible = msoTrue, "True", "False") & ";"
            themeSettings = themeSettings & "X1ShowGridLines=" & .HasMajorGridlines & ";"
            If .HasTitle Then
                themeSettings = themeSettings & "X1HasTitle=True;"
                themeSettings = themeSettings & "X1Title=" & .AxisTitle.Text & ";"
                themeSettings = themeSettings & "XAxisTitleFontSize=" & .AxisTitle.Font.Size & ";"
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
        themeFolderPath = FalFile.Combine_Paths(ThisWorkbook.Path, "FALCore/Themes")
    Else
        ' Fallback if workbook is not saved (e.g., new file)
        Me.HandleError "Please save the workbook before saving a theme."
        SaveThemeFromChart = False
        Exit Function
    End If

    themeFilePath = FalFile.Combine_Paths(themeFolderPath, themeName & ".theme")

    ' Save the theme file
    If FalFile.SaveStringToFile(themeSettings, themeFilePath) Then
        SaveThemeFromChart = True
    Else
        Me.HandleError "Failed to save the theme file to: " & themeFilePath
        SaveThemeFromChart = False
    End If

    Exit Function

ErrHandler:
    Me.HandleError "An unexpected error occurred in SaveThemeFromChart: " & Err.Description
    SaveThemeFromChart = False
End Function

Public Function ApplyTheme(ByVal themeName As String) As Boolean
    ' @brief Loads a theme from a file and applies it to the current chart.
    ' @param themeName The name of the theme to apply.
    ' @return True if the theme was applied successfully, False otherwise.
    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then
        Me.HandleError "No chart is currently set. Cannot apply theme."
        ApplyTheme = False
        Exit Function
    End If

    Dim themeSettings As String
    themeSettings = Me.LoadTheme(themeName)

    If themeSettings <> "" Then
        ' Apply the theme settings
        If Me.ApplyFormattingFromString(themeSettings) Then
            ' Store the theme name
            m_ThemeName = themeName
            ApplyTheme = True
        Else
            Me.HandleError "Failed to apply the theme '" & themeName & "'."
            ApplyTheme = False
        End If
    Else
        ' LoadTheme already handled the error message
        ApplyTheme = False
    End If

    Exit Function

ErrHandler:
    Me.HandleError "An unexpected error occurred in ApplyTheme: " & Err.Description
    ApplyTheme = False
End Function

Public Function ApplyFormattingFromString(Optional FormattingOptions As Variant) As Boolean
    ' @brief Applies a multitude of formatting options to the chart object.
    ' @details This is the main formatting engine. It uses the class properties as defaults,
    '          but allows them to be overridden by a semicolon-delimited string of options.
    ' @param FormattingOptions (Optional) A string of key-value pairs (e.g., "Title=MyTitle;XTitle=X-Axis").
    ' @return True if formatting was successful, False otherwise.
    On Error GoTo ErrHandler

    If m_Chart Is Nothing Then Exit Function

    Dim sKey As String, sValue As String
    Dim Options As Variant, index As Long
    Dim ser As series, axNo As Axis

    ' Use class properties as default values. The FormattingOptions string will override them.
    ' This section processes the optional FormattingOptions string.
    If Not IsMissing(FormattingOptions) Then
        If Len(FormattingOptions & "") > 0 Then
            Options = Split(FormattingOptions, ";")
            For index = 0 To UBound(Options)
                If InStr(Options(index), "=") > 0 Then
                    sKey = Trim(Split(Options(index), "=")(0))
                    sValue = Trim(Split(Options(index), "=")(1))

                    Select Case LCase(sKey)
                        Case "charttype": m_ChartType = CLng(sValue)
                        Case "charttitle", "title": m_ChartTitle = sValue
                        Case "plotby": m_PlotBy = CLng(sValue)
                        Case "fontpolice", "fontname": m_FontName = sValue
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
        End If

        If m_FontName <> "Auto" Then .ChartArea.Font.name = m_FontName

        With .PlotArea.Format.Line
            .ForeColor.RGB = RGB(150, 150, 150)
            .Weight = m_PlotAreaLineWeight
        End With
        .PlotArea.Interior.ColorIndex = xlNone

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
            If .HasTitle Then .AxisTitle.Text = m_X1Title
            .ScaleType = m_X1ScaleType
            If IsNumeric(m_X1Min) Then .MinimumScale = m_X1Min Else .MinimumScaleIsAuto = True
            If IsNumeric(m_X1Max) Then .MaximumScale = m_X1Max Else .MaximumScaleIsAuto = True
            .TickLabels.NumberFormat = m_X1LabelNumberFormat
            .TickLabels.Font.Size = m_XAxisTicksFontSize
            If .HasTitle Then .AxisTitle.Font.Size = m_XAxisTitleFontSize
            .HasMajorGridlines = m_X1ShowGridLines
        End With

        ' Primary Y-Axis (Value)
        With .Axes(xlValue, xlPrimary)
            .HasTitle = m_Y1HasTitle
            If .HasTitle Then .AxisTitle.Text = m_Y1Title
            .ScaleType = m_Y1ScaleType
            If IsNumeric(m_Y1Min) Then .MinimumScale = m_Y1Min Else .MinimumScaleIsAuto = True
            If IsNumeric(m_Y1Max) Then .MaximumScale = m_Y1Max Else .MaximumScaleIsAuto = True
            .TickLabels.NumberFormat = m_Y1LabelNumberFormat
            .TickLabels.Font.Size = m_YAxisTicksFontSize
            If .HasTitle Then .AxisTitle.Font.Size = m_YAxisTitleFontSize
            .HasMajorGridlines = m_Y1ShowGridLines
        End With

        ' And so on for secondary axes...

        If m_SquarePlot Then
            If .HasLegend Then .Legend.Position = xlRight
            .PlotArea.Position = xlChartElementPositionAutomatic
            .PlotArea.Width = .PlotArea.Height
        End If

        Me.FormatTextSubscripts

    End With

    ApplyFormattingFromString = True
    Exit Function

ErrHandler:
    Me.HandleError "Error while formatting chart: " & Err.Description
    ApplyFormattingFromString = False
End Function

Public Function FormatTextSubscripts() As Boolean
    ' @brief Formats subscripts and superscripts in chart titles and axis titles.
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

Private Function CreateChart(ByVal ChartName As String, ByVal wbk As Workbook, Optional ByVal rangeData As Range) As Chart
    On Error GoTo ifError

    Dim cleanName As String
    cleanName = FalLang.RemoveChars(ChartName, ":\/?*[];")
    cleanName = FalLang.ResizeString(cleanName, 31)

    On Error Resume Next
    wbk.Charts(cleanName).Delete
    On Error GoTo ifError

    Dim newChartSheet As Chart
    Set newChartSheet = wbk.Charts.Add
    With newChartSheet
        .Move After:=wbk.Worksheets(wbk.Worksheets.Count)
        .name = cleanName
        .ChartType = xlXYScatterLinesNoMarkers
    End With

    If Not rangeData Is Nothing Then
        newChartSheet.SetSourceData Source:=rangeData
    Else
        Do While newChartSheet.SeriesCollection.Count > 0
            newChartSheet.SeriesCollection(1).Delete
        Loop
    End If

    Set CreateChart = newChartSheet
    Exit Function
ifError:
    Set CreateChart = Nothing
End Function

Public Function PlotFromRange(ByVal data_src As Range, Optional ByVal ChartName As String = "", Optional ByVal nbSeriesByGroup As Integer = -1, Optional ByVal ColorStyle As String = "DefaultStyle", Optional ByVal FormattingOptions As String = "") As Boolean
    On Error GoTo ifError

    Dim tempChartName As String
    If ChartName = "" Then
        tempChartName = FalLang.RemoveChars(data_src.Parent.name, "_-. ")
    Else
        tempChartName = ChartName
    End If

    Dim tempFormattingOptions As String
    tempFormattingOptions = "Title=" & tempChartName & ";XTitle=" & data_src.Cells(1, 1).value & ";YTitle=(a.u);ChartType=" & xlXYScatterLinesNoMarkers & ";PlotBy=" & xlColumns & ";AutoLegend=0;" & FormattingOptions

    Set m_Chart = Me.CreateChart(tempChartName, data_src.Parent.Parent)
    If m_Chart Is Nothing Then GoTo ifError

    Me.AddSeriesFromRange data_src
    Me.DeleteMatchingSeries data_src.Cells(1, 1).value, False
    Me.ColorSeries 1, nbSeriesByGroup, ColorStyle
    Me.ApplyFormattingFromString tempFormattingOptions

    PlotFromRange = True
    Exit Function
ifError:
    PlotFromRange = False
End Function

Public Function PlotFromRangeWithFormatting(ByVal data_src As Range, Optional ByVal ChartName As String = "", Optional ByVal nbSeriesByGroup As Integer = 0, Optional ByVal ColorStyle As String = "DefautStyle", Optional ByVal FormattingOptions As String = "") As Boolean
    On Error GoTo ifError

    If Not Me.PlotFromRange(data_src, ChartName, nbSeriesByGroup, ColorStyle) Then
        Me.HandleError "Failed to create the initial chart from the data range."
        PlotFromRangeWithFormatting = False
        Exit Function
    End If

    If Not Me.ApplyFormattingFromString(FormattingOptions) Then
        Me.HandleError "Failed to apply formatting to the chart."
        PlotFromRangeWithFormatting = False
        Exit Function
    End If

    PlotFromRangeWithFormatting = True
    Exit Function
ifError:
    Me.HandleError "An unexpected error occurred in PlotFromRangeWithFormatting."
    PlotFromRangeWithFormatting = False
End Function

Public Function PlotFromWorksheet(ByVal wks_src As Worksheet, Optional ByVal DataTopLeftCell As String = "A1", Optional ByVal ColorStyle As String = "DefaultStyle", Optional ByVal plotSeries As Boolean = True, Optional ByVal nbseries As Integer = -1, Optional ByVal sPlot As String = " - Plot", Optional ByVal FormattingOptions As String = "") As Boolean
    ' @brief Automatically detects and plots data from a worksheet.
    On Error GoTo ifError

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
    chartTitle = FalLang.RemoveSubstrings(wks_src.name, ".txt;.xlsx;.mdm;xlsm; ")
    chartTitle = FalLang.ResizeString(chartTitle, 31)

    '--- Data Table Detection ---
    Dim LastLine As Long
    LastLine = FalWork.Get_Last_Row(wks_src, TopLeftCell.Column)
    firstLine = Me.FindFirstRowFromBottom(wks_src, TopLeftCell.Column, LastLine)

    LastCol_Index = FalWork.Get_Last_Column(wks_src, firstLine)
    Dim LastCol_Name As String, firstcol_name As String
    LastCol_Name = FalLang.col(LastCol_Index)
    firstcol_name = FalLang.col(TopLeftCell.Column)

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
    Chart_Name = FalLang.ResizeString(FalLang.RemoveChars(wks_src.name, "_-. "), 10) & " - " & sPlot & "(All)"
    Dim tempFormattingOptions As String
    tempFormattingOptions = "Title=" & chartTitle & ";XTitle=" & aLabel(1) & ";YTitle=(a.u);ChartType=" & xlXYScatterLinesNoMarkers & ";PlotBy=" & xlColumns & ";AutoLegend=0;" & FormattingOptions

    Set m_Chart = Me.CreateChart(Chart_Name, wks_src.Parent)
    If m_Chart Is Nothing Then GoTo ifError

    Me.AddSeriesFromRange wks_src.Range(firstcol_name & (firstLine - 1) & ":" & LastCol_Name & LastLine)
    Me.DeleteMatchingSeries wks_src.Range(firstcol_name & (firstLine - 1)).value, False
    Me.ColorSeries 1, ((NbColumn / nbseries) - 1), ColorStyle
    Me.ApplyFormattingFromString tempFormattingOptions

    PlotFromWorksheet = True
    Exit Function
ifError:
    PlotFromWorksheet = False
End Function

Public Function PlotFromArray(ByVal Arr2D As Variant, Optional ByVal ChartName As String = "Array Plot", Optional ByVal FormattingOptions As String = "") As Boolean
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
    If Not Me.PlotFromRange(tempRange, ChartName, , , FormattingOptions) Then
        GoTo cleanup
    End If

    PlotFromArray = True

cleanup:
    ' Clean up the temporary workbook
    Application.DisplayAlerts = False
    tempWb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    If Not PlotFromArray Then PlotFromArray = False ' Ensure it's false on error
    Exit Function
ifError:
    PlotFromArray = False
    Resume cleanup
End Function

Public Function AddTrendline(ByVal seriesIndex As Integer, Optional ByVal trendlineType As XlTrendlineType = xlLinear, Optional ByVal order As Integer = 2, Optional ByVal forecastForward As Double = 0, Optional ByVal forecastBackward As Double = 0, Optional ByVal displayEquation As Boolean = False, Optional ByVal displayRSquared As Boolean = False) As Boolean
    ' @brief Adds a trendline to a specified series on the chart.
    ' @param seriesIndex The index number of the series to add the trendline to.
    ' @param trendlineType (Optional) The type of trendline to add (e.g., xlLinear, xlPolynomial). Defaults to xlLinear.
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

    m_Chart.SeriesCollection(seriesIndex).ErrorBar Direction:=direction, Include:=include, Type:=type, Amount:=amount

    AddErrorBars = True
    Exit Function

ErrHandler:
    AddErrorBars = False
End Function

Private Function CreateChartFromRange(ByVal data As Range, Optional ByVal chartType As XlChartType = xlXYScatter, Optional ByVal chartName As String = "") As Boolean
    On Error GoTo ifError

    Dim tempName As String
    If chartName = "" Then
        tempName = "Chart_" & data.Parent.Name
    Else
        tempName = chartName
    End If

    ' Use the existing CreateChart helper to create the chart sheet
    Set m_Chart = Me.CreateChart(tempName, data.Parent.Parent)
    If m_Chart Is Nothing Then GoTo ifError

    ' Set the chart type and source data
    m_Chart.chartType = chartType
    m_Chart.SetSourceData Source:=data

    CreateChartFromRange = True
    Exit Function
ifError:
    CreateChartFromRange = False
End Function

Public Function CreateBarChart(ByVal data As Range, Optional ByVal chartName As String = "Bar Chart", Optional ByVal style As XlChartType = xlBarClustered) As Boolean
    ' @brief Creates a new bar chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of bar chart to create. Defaults to xlBarClustered.
    ' @return True if the chart was created successfully, False otherwise.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreateBarChart ActiveSheet.Range("A1:B5")
    CreateBarChart = Me.CreateChartFromRange(data, style, chartName)
End Function

Public Function CreatePieChart(ByVal data As Range, Optional ByVal chartName As String = "Pie Chart", Optional ByVal style As XlChartType = xlPie) As Boolean
    ' @brief Creates a new pie chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of pie chart to create. Defaults to xlPie.
    ' @return True if the chart was created successfully, False otherwise.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreatePieChart ActiveSheet.Range("A1:B5")
    CreatePieChart = Me.CreateChartFromRange(data, style, chartName)
End Function

Public Function CreateLineChart(ByVal data As Range, Optional ByVal chartName As String = "Line Chart", Optional ByVal style As XlChartType = xlLine) As Boolean
    ' @brief Creates a new line chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of line chart to create. Defaults to xlLine.
    ' @return True if the chart was created successfully, False otherwise.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreateLineChart ActiveSheet.Range("A1:B5")
    CreateLineChart = Me.CreateChartFromRange(data, style, chartName)
End Function

Public Function CreateAreaChart(ByVal data As Range, Optional ByVal chartName As String = "Area Chart", Optional ByVal style As XlChartType = xlArea) As Boolean
    ' @brief Creates a new area chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of area chart to create. Defaults to xlArea.
    ' @return True if the chart was created successfully, False otherwise.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreateAreaChart ActiveSheet.Range("A1:B5")
    CreateAreaChart = Me.CreateChartFromRange(data, style, chartName)
End Function

Public Function CreateSurfaceChart(ByVal data As Range, Optional ByVal chartName As String = "Surface Chart", Optional ByVal style As XlChartType = xlSurface) As Boolean
    ' @brief Creates a new surface chart from a data range.
    ' @param data The range containing the data to plot.
    ' @param chartName (Optional) The name for the new chart.
    ' @param style (Optional) The specific type of surface chart to create. Defaults to xlSurface.
    ' @return True if the chart was created successfully, False otherwise.
    ' @example
    ' Dim plot As New FalPlot
    ' plot.CreateSurfaceChart ActiveSheet.Range("A1:C10")
    CreateSurfaceChart = Me.CreateChartFromRange(data, style, chartName)
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
    ' @brief Returns a predefined array of color indices for chart series.
    ' @param Style The name of the color palette to return.
    Select Case Style
        Case "DefaultStyle": ColorPalette = Array(3, 5, 4, 7, 46, 45, 44, 43, 41, 11, 49, 1, 56, 48, 23, 22, 21, 18, 17, 40, 39, 38, 37, 36, 35, 34, 33)
        Case "RainbowStyle": ColorPalette = Array(3, 30, 21, 54, 13, 18, 38, 7, 39, 49, 11, 5, 41, 37, 33, 8, 42, 14, 50, 10, 4, 43, 12, 40, 36, 6, 44, 45, 46, 53)
        Case Else: ColorPalette = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56)
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
