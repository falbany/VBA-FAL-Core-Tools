# FALCore VBA Suite

**Version: 1.0.0**  
**Author: Florent ALBANY**

---

## Introduction

FALCore is a comprehensive library of VBA modules designed to accelerate application development in Microsoft Excel. It provides a collection of robust, reusable, and well-documented functions for common and advanced programming tasks, allowing developers to focus on their core application logic instead of reinventing the wheel.

## Key Features

- **Modular Design**: The suite is organized into distinct modules, each focusing on a specific area (Files, Worksheets, Arrays, Logging). These modules can be used independently or together.
- **Robust & Reusable**: Functions are built with error handling and are designed to be easily integrated into any VBA project.
- **Well-Documented**: All public functions include detailed header comments explaining their purpose, parameters, and usage, making the library easy to learn and use.
- **Consistent Naming**: The library follows a clear `Fal...` prefix convention for all module names, providing a clean namespace that prevents naming collisions with other libraries or host application functions.

## Modules Overview

The FALCore suite is organized into the following modules:

- **`FALCore.bas`**: The central "About" module for the suite, containing version information and a general description of the library.
- **`FalFile.bas`**: A powerful set of utilities for file and folder operations, including reading, writing, copying, moving, sorting, and zipping files.
- **`FalWork.bas`**: A comprehensive collection of functions for creating, manipulating, and managing Excel Workbooks and Worksheets.
- **`FalArray.bas`**: An advanced toolkit for creating, manipulating, and querying 1D, 2D, 3D, and 4D arrays. Includes features like JSON conversion, regression analysis, and complex data slicing.
- **`FalLog.bas`**: A flexible logging utility with configurable debug levels (Error, Warning, Info, Debug) and multiple output destinations (Immediate Window, Text File).

## Installation

To use the FALCore suite in your project, follow these steps:

1. In the VBA Editor (`Alt+F11`), right-click in the Project Explorer and select **Import File...**.
2. Navigate to the `Modules` directory and select all the `.bas` files.
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

Copyright (c) 2023 Florent ALBANY.

The following may be altered and reused as you wish so long as the copyright notice is left unchanged (including Author, Website and Copyright). It may not be sold/resold or reposted on other sites (links back to this site are allowed).
