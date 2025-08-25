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
