# FALCore VBA Library

Welcome to the FALCore VBA Library, a collection of powerful and reusable modules and classes designed to accelerate Microsoft Excel development.

## Project Structure

The library is organized into two main folders:

- **/Classes**: Contains object-oriented class modules (`.cls` files) that provide advanced functionalities.
  - `FalPlot`: A powerful class for creating and manipulating charts.
- **/Modules**: Contains standard procedural modules (`.bas` files) with a wide range of helper functions.

## Getting Started

To use the FALCore library in your VBA project, you can import the desired `.cls` and `.bas` files.

Please refer to the `README.md` file in each subdirectory for more detailed information about the specific classes and modules available.

## Simplified Module Import

To simplify the process of importing all the FALCore modules and classes at once, you can use the `VBAProjectImporter.bas` module.

### How to Use

1.  First, manually import only the `FALCore/VBAProjectImporter.bas` file into your VBA project.
2.  **Important:** You must enable programmatic access to the VBA project. In Excel, go to `File > Options > Trust Center > Trust Center Settings > Macro Settings`, and check the box for **"Trust access to the VBA project object model"**.
3.  Run the `ImportVBAFiles` macro from the `VBAProjectImporter` module.
4.  You will be prompted to select the `FALCore` folder.
5.  The macro will automatically import all `.bas` and `.cls` files from the `/Classes` and `/Modules` subdirectories into your project.
