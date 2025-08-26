# FALCore VBA Suite

**Version: 1.1.0**
**Author: Florent ALBANY**

---

## Introduction

FALCore is a comprehensive library of VBA modules designed to accelerate application development in Microsoft Excel. It provides a collection of robust, reusable, and well-documented functions for common and advanced programming tasks, allowing developers to focus on their core application logic instead of reinventing the wheel.

## Key Features

- **Hybrid Approach**: The library offers both procedural modules and object-oriented classes, providing flexibility for different programming styles.
- **Modular Design**: The suite is organized into distinct modules and classes, each focusing on a specific area (Files, Worksheets, Arrays, Plotting, etc.).
- **Robust & Reusable**: Functions and methods are built with error handling and are designed to be easily integrated into any VBA project.
- **Well-Documented**: All public members include detailed header comments explaining their purpose, parameters, and usage.
- **Consistent Naming**: The library follows a clear `Fal...` prefix convention, providing a clean namespace.

## Architecture Overview

FALCore is structured into two main components:

- **`/FALCore/Classes`**: Contains powerful, object-oriented class modules (`.cls`) for complex tasks.
- **`/FALCore/Modules`**: Contains a wide range of procedural helper modules (`.bas`).

### Classes Overview

- **`FalPlot.cls`**: A powerful class for creating and manipulating charts. It provides an object-oriented interface for plotting data and customizing every aspect of a chart's appearance.

### Modules Overview

- **`FalPlot.bas`**: A backward-compatible wrapper for the `FalPlot` class. It provides a simple, procedural interface for common plotting tasks.
- **`FalArray.bas`**: An advanced toolkit for array manipulation.
- **`FalCSV.bas`**: A module for working with CSV files.
- **`FalFile.bas`**: A powerful set of utilities for file and folder operations.
- **`FalLang.bas`**: A module for language-related functions.
- **`FalLog.bas`**: A flexible logging utility.
- **`FalUtils.bas`**: A collection of utility functions.
- **`FalWork.bas`**: A comprehensive collection of functions for managing Workbooks and Worksheets.
- **`FalXls.bas`**: A module for Excel-specific functions, including project-level utilities like creating summary sheets and exporting/importing all VBA components.

## Installation

To use the FALCore suite in your project, follow these steps:

1. In the VBA Editor (`Alt+F11`), right-click in the Project Explorer and select **Import File...**.
2. Navigate to the `FALCore/Modules` directory and select all the `.bas` files.
3. Go to **Tools -> References** in the VBA Editor.
4. Ensure that **"Microsoft Scripting Runtime"** is checked. This is required for `Dictionary` objects and the `FileSystemObject` used in `FalFile` and `FalWork`.

## Quick Start Example

Here is a simple example demonstrating how to use several modules from the FALCore suite together.

```vba
Sub FALCore_Demo()
    ' 1. Initialize the logger to show all messages in the Immediate Window
    FalLog.InitializeLogger Level:=llDebug, Destination:=ldImmediate

    FalLog.LogMessage llInfo, "Demo.Start", "FALCore demo started."

    ' 2. Create a new workbook and a worksheet using FalWork
    Dim wbk As Workbook
    Set wbk = FalWork.Create_Workbook(MakeVisible:=True)
    If wbk Is Nothing Then
        FalLog.LogMessage llError, "Demo.Create", "Failed to create workbook."
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = FalWork.Create_Worksheet("DemoSheet", wbk)
    FalLog.LogMessage llDebug, "Demo.Setup", "Workbook and worksheet created successfully."

    ' 3. Create a 2D array with FalArray and write it to the sheet with FalWork
    Dim dataArray As Variant
    dataArray = FalArray.a2D_Create(NumRows:=3, NumCols:=4, FillValue:="Test")

    ' Write the array to the worksheet as values
    If FalWork.Write_Array_To_Worksheet(dataArray, ws.Range("A1")) Then
        FalLog.LogMessage llInfo, "Demo.Write", "Successfully wrote 2D array to DemoSheet!A1."
    Else
        FalLog.LogMessage llError, "Demo.Write", "Failed to write array to worksheet."
    End If

    ' 4. Clean up (optional)
    ' wbk.Close SaveChanges:=False
    FalLog.LogMessage llInfo, "Demo.End", "FALCore demo finished."
End Sub
```

## Author

- **Florent ALBANY**

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
