# FALCore Classes

This directory contains the class modules (`.cls`) for the FALCore library.

## FalPlot

`FalPlot.cls` is a powerful class for creating and manipulating charts in Excel. It provides an object-oriented interface to a wide range of chart functionalities.

### Example Usage

Here is a basic example of how to use the `FalPlot` class:

```vba
Sub PlotExample()
    Dim plot As New FalPlot
    Dim dataRange As Range

    ' Set the data source
    Set dataRange = ThisWorkbook.ActiveSheet.Range("A1:B10")

    ' Create a new chart and assign it to the plot object
    Set plot.Chart = ThisWorkbook.ActiveSheet.Shapes.AddChart2(240, xlXYScatterLines).Chart
    plot.SetSourceData dataRange

    ' Customize the chart using properties
    plot.ChartTitle = "My Awesome Plot"
    plot.X1Title = "Time (s)"
    plot.Y1Title = "Voltage (V)"
    plot.HasLegend = False

    ' Apply all formatting at once
    plot.FormatChart
End Sub
```

For more details on all the available properties and methods, please refer to the extensive documentation within the `FalPlot.cls` file itself.

### Theming Engine

The `FalPlot` class includes a simple yet powerful theming engine that allows you to save the visual style of a chart and apply it to other charts. This is extremely useful for maintaining a consistent look and feel across your projects.

Themes are saved as simple text files (`.theme`) in a `FALCore/Themes` directory, which is created automatically in the same folder as your workbook.

#### Saving a Theme

First, format a chart exactly how you want it, either manually or using the `FalPlot` properties. Then, you can save its style as a new theme.

```vba
Sub SaveChartTheme()
    Dim plot As New FalPlot

    ' You must have a chart selected, or assign one to the plot object
    If ActiveChart Is Nothing Then
        MsgBox "Please select a chart first."
        Exit Sub
    End If

    Set plot.Chart = ActiveChart

    ' Save the theme with a descriptive name
    If plot.SaveThemeFromChart("MyCorporateTheme") Then
        MsgBox "Theme 'MyCorporateTheme' saved successfully!"
    End If
End Sub
```

#### Applying a Theme

Once a theme is saved, you can apply it to any other chart.

```vba
Sub ApplyChartTheme()
    Dim plot As New FalPlot
    Dim dataRange As Range

    ' Data for our new chart
    Set dataRange = ThisWorkbook.ActiveSheet.Range("D1:E10")

    ' Create a new chart
    Set plot.Chart = ThisWorkbook.ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Chart
    plot.SetSourceData dataRange

    ' Apply the saved theme
    If plot.ApplyTheme("MyCorporateTheme") Then
        ' You can still make specific tweaks after applying the theme
        plot.ChartTitle = "New Chart with Corporate Theme"
        plot.HasLegend = True
    End If
End Sub
```
