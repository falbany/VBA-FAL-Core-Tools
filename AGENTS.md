# Agent Instructions

This document provides instructions for AI agents working with the FALCore VBA Suite.

## Project Overview

FALCore is a comprehensive library of VBA modules and classes designed to accelerate application development in Microsoft Excel. The project is managed as a collection of source files (`.bas`, `.cls`) that are intended to be built into a functional `.xlsm` workbook.

## Project Structure

- **`/FALCore/Classes/`**: Contains powerful, object-oriented class modules (`.cls`).
- **`/FALCore/Modules/`**: Contains a wide range of procedural helper modules (`.bas`).
- **`/FALCore/demo/`**: Contains demonstration modules that showcase how to use the classes and modules.
- **`/modules/`**: Contains third-party VBA libraries managed as Git submodules. See `modules/AGENTS.md` for specific instructions on using these.
- **`BuildFALCoreProject.bas`**: A script used to automate the assembly of the VBA project from all source files.
- **`README.md`**: The main project documentation.

## Development Workflow

This project is not developed directly within an Excel file. Instead, the source code is maintained as individual text files, and a build process is used to create a functional workbook.

### 1. Initializing Submodules
If you have just cloned the repository, you **must** initialize the submodules first:
```bash
git submodule update --init --recursive
```

### 2. Making Code Changes
Modify the existing `.bas` and `.cls` files or create new ones as needed. Follow the naming and code style conventions outlined below.

### 3. Building the Project
As an AI agent, you cannot run the build process yourself. However, you should understand it. The `BuildFALCoreProject.bas` script is designed to be run from within a host Excel workbook. It imports all necessary project files, including the submodules, to create a single, runnable `FALCore.xlsm` file.

## Naming Conventions

- All public functions and subs in the core modules should be prefixed with `Fal`.
- Core module names should be prefixed with `Fal`.

## Code Style and Conventions

### Documentation and Comments

- All public `Function`, `Sub`, and `Property` definitions must be preceded by a documentation block.
- The documentation block must use `@` annotations: `@brief`, `@param <name>`, `@return`, `@example`.

### Class Structure

Class modules (`.cls` files) should be organized into the following logical sections, in order:

1.  Header and `Option Explicit`
2.  Private Member Variables
3.  Public Properties
4.  Initialization (`Class_Initialize`)
5.  Public `Create...` or `Add...` Methods
6.  Other Public Methods
7.  Private Helper Functions

### Naming

- Use descriptive names for functions and variables.
- Private helper functions should be prefixed with `prv`.

## Key Dependencies

- **`clsMDM`**: This class depends on the `VBA-Dictionary` and `VBA-StringBuilder` submodules.
- **`FalPlot`**: `FalLang`, `FalWork`, `FalArray`, `FalFile`
- **`FalCSV`**: `FalFile`, `FalArray`
- **`FalFile`**: `FALCore`, `FalLang`, `FalWork`, `FalArray`
- **`FalWork`**: `FalFile`