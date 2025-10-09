# FALCore VBA Suite

**Version: 1.2.0**
**Author: Florent ALBANY**

---

## Introduction

FALCore is a comprehensive library of VBA modules designed to accelerate application development in Microsoft Excel. It provides a collection of robust, reusable, and well-documented functions for common and advanced programming tasks, allowing developers to focus on their core application logic instead of reinventing the wheel.

## Key Features

- **Hybrid Approach**: The library offers both procedural modules and object-oriented classes, providing flexibility for different programming styles.
- **Modular Design**: The suite is organized into distinct modules and classes, each focusing on a specific area (Files, Worksheets, Arrays, Plotting, etc.).
- **Extensible with Submodules**: Includes a curated set of powerful third-party VBA libraries as git submodules for advanced functionality like JSON handling and high-performance operations.
- **Robust & Reusable**: Functions and methods are built with error handling and are designed to be easily integrated into any VBA project.
- **Well-Documented**: All public members include detailed header comments explaining their purpose, parameters, and usage.
- **Consistent Naming**: The library follows a clear `Fal...` prefix convention, providing a clean namespace.

## Architecture Overview

FALCore is structured into three main components:

- **`/FALCore/Classes`**: Contains powerful, object-oriented class modules (`.cls`) for complex tasks.
- **`/FALCore/Modules`**: Contains a wide range of procedural helper modules (`.bas`).
- **`/modules`**: Contains third-party VBA libraries managed as Git submodules.

### Classes Overview

- **`FalPlot.cls`**: A powerful class for creating and manipulating charts. It provides an object-oriented interface for plotting data and customizing every aspect of a chart's appearance.
- **`clsMDM.cls`**: A robust class for handling Keysight's MDM (Measurement Data Model) files. It supports reading, creating, and exporting MDM data to various formats, including MDM strings, 2D arrays for Excel, and JSON.

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

### Submodules Overview

This project includes several powerful third-party libraries managed as Git submodules. See the `modules/README.md` for more details.

- **VBA-Dictionary**: A cross-platform `Dictionary` object.
- **VBA-JSON**: A JSON parser and converter.
- **VBA-StringBuilder**: For high-performance string concatenation.
- **VBA-Better-Array**: A more powerful array class.
- **VBA-Log**: A flexible logging framework.

## Installation

This project is designed to be built from source files using the included build script.

### Step 1: Clone the Repository and Submodules

First, clone the repository. Because this project uses Git submodules, you must initialize them after cloning.

```bash
# Clone the main repository
git clone <repository_url>
cd <repository_name>

# Initialize and fetch the submodules
git submodule update --init --recursive
```

### Step 2: Build the VBA Project

The `BuildFALCoreProject.bas` script automates the process of importing all necessary source files into a single, functional Excel workbook.

1.  **Open Microsoft Excel** and create a new, blank workbook.
2.  **Save the workbook** as `FALCore.xlsm` in the root directory of this project.
3.  **Open the VBA Editor** (`Alt` + `F11`).
4.  In the VBA Editor, go to **Insert > Module**.
5.  Copy the entire content of `BuildFALCoreProject.bas` and paste it into the new module.
6.  Go to **Tools -> References** and ensure that **"Microsoft Scripting Runtime"** is checked.
7.  Run the `BuildProject` macro.
8.  When prompted, select the root folder of the project.

The script will import all modules and classes, creating a complete, runnable project in `FALCore.xlsm`.

## Quick Start Example

The `demoMDM.bas` module (included by the build script) provides a comprehensive example of how to use the `clsMDM` class. To run it:

1.  Open the built `FALCore.xlsm` file.
2.  Open the VBA Editor (`Alt` + `F11`).
3.  Open the Immediate Window (`Ctrl` + `G`).
4.  In any module, type `RunAdvancedMdmDemo` and press Enter, or place your cursor inside the `RunAdvancedMdmDemo` sub and press `F5`.

The demo will create an MDM object in memory, print its string and JSON representations to the Immediate Window, and create a new worksheet with the exported data.

## Author

- **Florent ALBANY**

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.