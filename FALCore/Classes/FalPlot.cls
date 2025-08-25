VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FalPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

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
End Property

Public Property Get ChartTitle() As String
    ChartTitle = m_ChartTitle
End Property

Public Property Let ChartTitle(ByVal value As String)
    m_ChartTitle = value
End Property

Public Property Get Y1Title() As String
    Y1Title = m_Y1Title
End Property

Public Property Let Y1Title(ByVal value As String)
    m_Y1Title = value
End Property

Public Property Get Y2Title() As String
    Y2Title = m_Y2Title
End Property

Public Property Let Y2Title(ByVal value As String)
    m_Y2Title = value
End Property

Public Property Get X1Title() As String
    X1Title = m_X1Title
End Property

Public Property Let X1Title(ByVal value As String)
    m_X1Title = value
End Property

Public Property Get X2Title() As String
    X2Title = m_X2Title
End Property

Public Property Let X2Title(ByVal value As String)
    m_X2Title = value
End Property

Public Property Get X1ScaleType() As XlScaleType
    X1ScaleType = m_X1ScaleType
End Property

Public Property Let X1ScaleType(ByVal value As XlScaleType)
    m_X1ScaleType = value
End Property

Public Property Get Y1ScaleType() As XlScaleType
    Y1ScaleType = m_Y1ScaleType
End Property

Public Property Let Y1ScaleType(ByVal value As XlScaleType)
    m_Y1ScaleType = value
End Property

Public Property Get X2ScaleType() As XlScaleType
    X2ScaleType = m_X2ScaleType
End Property

Public Property Let X2ScaleType(ByVal value As XlScaleType)
    m_X2ScaleType = value
End Property

Public Property Get Y2ScaleType() As XlScaleType
    Y2ScaleType = m_Y2ScaleType
End Property

Public Property Let Y2ScaleType(ByVal value As XlScaleType)
    m_Y2ScaleType = value
End Property

Public Property Get HasLegend() As Boolean
    HasLegend = m_HasLegend
End Property

Public Property Let HasLegend(ByVal value As Boolean)
    m_HasLegend = value
End Property

Public Property Get HasTitle() As Boolean
    HasTitle = m_HasTitle
End Property

Public Property Let HasTitle(ByVal value As Boolean)
    m_HasTitle = value
End Property

Public Property Get X1HasTitle() As Boolean
    X1HasTitle = m_X1HasTitle
End Property

Public Property Let X1HasTitle(ByVal value As Boolean)
    m_X1HasTitle = value
End Property

Public Property Get Y1HasTitle() As Boolean
    Y1HasTitle = m_Y1HasTitle
End Property

Public Property Let Y1HasTitle(ByVal value As Boolean)
    m_Y1HasTitle = value
End Property

Public Property Get X2HasTitle() As Boolean
    X2HasTitle = m_X2HasTitle
End Property

Public Property Let X2HasTitle(ByVal value As Boolean)
    m_X2HasTitle = value
End Property

Public Property Get Y2HasTitle() As Boolean
    Y2HasTitle = m_Y2HasTitle
End Property

Public Property Let Y2HasTitle(ByVal value As Boolean)
    m_Y2HasTitle = value
End Property

Public Property Get X1ShowAxis() As Boolean
    X1ShowAxis = m_X1ShowAxis
End Property

Public Property Let X1ShowAxis(ByVal value As Boolean)
    m_X1ShowAxis = value
End Property

Public Property Get Y1ShowAxis() As Boolean
    Y1ShowAxis = m_Y1ShowAxis
End Property

Public Property Let Y1ShowAxis(ByVal value As Boolean)
    m_Y1ShowAxis = value
End Property

Public Property Get X1ShowGridLines() As Boolean
    X1ShowGridLines = m_X1ShowGridLines
End Property

Public Property Let X1ShowGridLines(ByVal value As Boolean)
    m_X1ShowGridLines = value
End Property

Public Property Get Y1ShowGridLines() As Boolean
    Y1ShowGridLines = m_Y1ShowGridLines
End Property

Public Property Let Y1ShowGridLines(ByVal value As Boolean)
    m_Y1ShowGridLines = value
End Property

Public Property Get FontName() As String
    FontName = m_FontName
End Property

Public Property Let FontName(ByVal value As String)
    m_FontName = value
End Property

Public Property Get TitleFontSize() As Integer
    TitleFontSize = m_TitleFontSize
End Property

Public Property Let TitleFontSize(ByVal value As Integer)
    m_TitleFontSize = value
End Property

Public Property Get XAxisTitleFontSize() As Integer
    XAxisTitleFontSize = m_XAxisTitleFontSize
End Property

Public Property Let XAxisTitleFontSize(ByVal value As Integer)
    m_XAxisTitleFontSize = value
End Property

Public Property Get YAxisTitleFontSize() As Integer
    YAxisTitleFontSize = m_YAxisTitleFontSize
End Property

Public Property Let YAxisTitleFontSize(ByVal value As Integer)
    m_YAxisTitleFontSize = value
End Property

Public Property Get XAxisTicksFontSize() As Integer
    XAxisTicksFontSize = m_XAxisTicksFontSize
End Property

Public Property Let XAxisTicksFontSize(ByVal value As Integer)
    m_XAxisTicksFontSize = value
End Property

Public Property Get YAxisTicksFontSize() As Integer
    YAxisTicksFontSize = m_YAxisTicksFontSize
End Property

Public Property Let YAxisTicksFontSize(ByVal value As Integer)
    m_YAxisTicksFontSize = value
End Property

Public Property Get LegendFontSize() As Integer
    LegendFontSize = m_LegendFontSize
End Property

Public Property Let LegendFontSize(ByVal value As Integer)
    m_LegendFontSize = value
End Property

Public Property Get LegendPosition() As XlLegendPosition
    LegendPosition = m_LegendPosition
End Property

Public Property Let LegendPosition(ByVal value As XlLegendPosition)
    m_LegendPosition = value
End Property

Public Property Get LegendInLayout() As Boolean
    LegendInLayout = m_LegendInLayout
End Property

Public Property Let LegendInLayout(ByVal value As Boolean)
    m_LegendInLayout = value
End Property

Public Property Get PlotAreaLineWeight() As Double
    PlotAreaLineWeight = m_PlotAreaLineWeight
End Property

Public Property Let PlotAreaLineWeight(ByVal value As Double)
    m_PlotAreaLineWeight = value
End Property

Public Property Get SeriesLineWeight() As Double
    SeriesLineWeight = m_SeriesLineWeight
End Property

Public Property Let SeriesLineWeight(ByVal value As Double)
    m_SeriesLineWeight = value
End Property

Public Property Get SeriesLineDashStyle() As MsoLineDashStyle
    SeriesLineDashStyle = m_SeriesLineDashStyle
End Property

Public Property Let SeriesLineDashStyle(ByVal value As MsoLineDashStyle)
    m_SeriesLineDashStyle = value
End Property

Public Property Get SeriesMarkerStyle() As XlMarkerStyle
    SeriesMarkerStyle = m_SeriesMarkerStyle
End Property

Public Property Let SeriesMarkerStyle(ByVal value As XlMarkerStyle)
    m_SeriesMarkerStyle = value
End Property

Public Property Get HasDataLabels() As Boolean
    HasDataLabels = m_HasDataLabels
End Property

Public Property Let HasDataLabels(ByVal value As Boolean)
    m_HasDataLabels = value
End Property

Public Property Get X1LabelNumberFormat() As String
    X1LabelNumberFormat = m_X1LabelNumberFormat
End Property

Public Property Let X1LabelNumberFormat(ByVal value As String)
    m_X1LabelNumberFormat = value
End Property

Public Property Get Y1LabelNumberFormat() As String
    Y1LabelNumberFormat = m_Y1LabelNumberFormat
End Property

Public Property Let Y1LabelNumberFormat(ByVal value As String)
    m_Y1LabelNumberFormat = value
End Property

Public Property Get X2LabelNumberFormat() As String
    X2LabelNumberFormat = m_X2LabelNumberFormat
End Property

Public Property Let X2LabelNumberFormat(ByVal value As String)
    m_X2LabelNumberFormat = value
End Property

Public Property Get Y2LabelNumberFormat() As String
    Y2LabelNumberFormat = m_Y2LabelNumberFormat
End Property

Public Property Let Y2LabelNumberFormat(ByVal value As String)
    m_Y2LabelNumberFormat = value
End Property

Public Property Get Y1Min() As Variant
    Y1Min = m_Y1Min
End Property

Public Property Let Y1Min(ByVal value As Variant)
    m_Y1Min = value
End Property

Public Property Get Y1Max() As Variant
    Y1Max = m_Y1Max
End Property

Public Property Let Y1Max(ByVal value As Variant)
    m_Y1Max = value
End Property

Public Property Get X1Min() As Variant
    X1Min = m_X1Min
End Property

Public Property Let X1Min(ByVal value As Variant)
    m_X1Min = value
End Property

Public Property Get X1Max() As Variant
    X1Max = m_X1Max
End Property

Public Property Let X1Max(ByVal value As Variant)
    m_X1Max = value
End Property

Public Property Get Y2Min() As Variant
    Y2Min = m_Y2Min
End Property

Public Property Let Y2Min(ByVal value As Variant)
    m_Y2Min = value
End Property

Public Property Get Y2Max() As Variant
    Y2Max = m_Y2Max
End Property

Public Property Let Y2Max(ByVal value As Variant)
    m_Y2Max = value
End Property

Public Property Get X2Min() As Variant
    X2Min = m_X2Min
End Property

Public Property Let X2Min(ByVal value As Variant)
    m_X2Min = value
End Property

Public Property Get X2Max() As Variant
    X2Max = m_X2Max
End Property

Public Property Let X2Max(ByVal value As Variant)
    m_X2Max = value
End Property

Public Property Get SquarePlot() As Boolean
    SquarePlot = m_SquarePlot
End Property

Public Property Let SquarePlot(ByVal value As Boolean)
    m_SquarePlot = value
End Property

Public Property Get CrossesAt() As Double
    CrossesAt = m_CrossesAt
End Property

Public Property Let CrossesAt(ByVal value As Double)
    m_CrossesAt = value
End Property

Public Property Get PlotBy() As XlRowCol
    PlotBy = m_PlotBy
End Property

Public Property Let PlotBy(ByVal value As XlRowCol)
    m_PlotBy = value
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

Private Sub HandleError(ByVal ErrMsg As String)
    MsgBox "An error occurred:" & vbCrLf & ErrMsg, vbOKOnly + vbCritical, "Operation Impossible"
End Sub

Public Function CreateChart(ChartName As String, wbk As Workbook, Optional rangeData As Variant) As Chart
    ' Creation du graphique.
    ChartName = FalLang.RemoveChars(ChartName, ":\/?*[];")
    ChartName = FalLang.ResizeString(ChartName, 31)
    If FalWork.SheetExist(wbk, ChartName) Then wbk.Charts(ChartName).Delete


    Dim newChartSheet As Chart
    Set newChartSheet = wbk.Charts.Add
    With newChartSheet
        .Move After:=wbk.Worksheets(Worksheets.count)
        If ChartName <> "" Then .name = ChartName
        .ChartType = xlXYScatterLinesNoMarkers
    End With
    If Not IsMissing(rangeData) And TypeName(rangeData) = "Range" Then
        SetSource newChartSheet, rangeData
    Else
        ClearSeries newChartSheet
    End If
    Set CreateChart = newChartSheet
End Function

Public Function ClearSeries(myChart As Chart) As Boolean
    If myChart.SeriesCollection.count > 0 Then
        Do While myChart.SeriesCollection.count > 0
            myChart.SeriesCollection(1).Delete
        Loop
    End If
    If myChart.SeriesCollection.count = 0 Then ClearSeries = True Else ClearSeries = False
End Function

Public Function SetSource(myChart As Chart, rangeData As Variant, Optional PlotBy As Integer = 2) As Variant
    If TypeName(rangeData) <> "Range" Then SetSource = CVErr(2001): Exit Function
    On Error Resume Next
    myChart.SetSourceData Source:=rangeData
    myChart.PlotBy = PlotBy
    SetSource = True
End Function

Public Function AddSeriesFromRange(myChart As Chart, data_src As Range) As Variant
    Dim seriesRange As Range
    Dim newSeries As series
    Dim i As Integer

    For i = 2 To data_src.Columns.count
        If myChart.FullSeriesCollection.count >= 255 Then GoTo ifError

        If data_src.Cells(1, i).value <> data_src.Cells(1, 1).value Then
            Set newSeries = myChart.SeriesCollection.newSeries

            Set seriesRange = data_src.Columns(i).Offset(1, 0).Resize(data_src.Rows.count - 1, 1)
            newSeries.values = seriesRange

            Set seriesRange = data_src.Columns(1).Offset(1, 0).Resize(data_src.Rows.count - 1, 1)
            newSeries.XValues = seriesRange

            newSeries.name = "='" & data_src.parent.name & "'!" & data_src.Cells(1, i).Address
        End If
    Next i
    AddSeriesFromRange = True

    Exit Function

ifError:
    AddSeriesFromRange = True
End Function

Function AddSeriesFromArray(Chart As Chart, Arr2D() As Variant) As Variant
    Dim newSeries       As series
    Dim i               As Integer

    For i = LBound(Arr2D) + 1 To UBound(Arr2D)
        If Chart.FullSeriesCollection.count >= 256 Then Exit Function

        Set newSeries = Chart.SeriesCollection.newSeries

        newSeries.values = Application.index(Arr2D, 0, i + 1)

        newSeries.XValues = Application.index(Arr2D, 0, 1)

        newSeries.name = Arr2D(1, i + 1)
    Next i
    AddSeriesFromArray = True
End Function
