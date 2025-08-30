# Agent Instructions

This document provides instructions for AI agents working with the FALCore VBA Suite.

## Project Overview

FALCore is a library of VBA modules for Microsoft Excel. The goal is to provide a set of reusable functions to accelerate development.

## Project Structure

- `FALCore/Modules/`: Contains the VBA modules (`.bas` files).
- `Other/Classes/`: Contains other VBA classes (`.cls` files).
- `README.md`: The main documentation file.
- `LICENSE`: The project license.
- `CHANGELOG.md`: The project changelog.
- `CONTRIBUTING.md`: Contribution guidelines.

## Naming Conventions

- All public functions and subs in the modules should be prefixed with `Fal`.
- Module names should be prefixed with `Fal`.

## Code Style and Conventions

### Documentation and Comments

- All public `Function`, `Sub`, and `Property` definitions should be preceded by a documentation block.
- The documentation block should use `@` annotations.
- Use `@brief` for a short, one-line description of the function's purpose.
- Use `@param <name>` to describe each parameter.
- Use `@return` to describe the function's return value.
- Use `@example` to provide a clear example of how to use the code.
- Use `@see` to reference related functions or modules.

### Class Structure

For better readability and maintainability, class modules (`.cls` files) should be organized into the following logical sections, in order:

1.  Header and `Option Explicit`
2.  Private Member Variables
3.  Public Properties
4.  Initialization (`Class_Initialize`)
5.  Public `Create...` Methods (if any)
6.  Other Public Methods
7.  Private Helper Functions

### Naming

- Use descriptive names for functions and variables (e.g., `CreateChartFromWorksheet` is preferred over `PlotFromWks`).
- Private helper functions should be prefixed with `prv`.

## Development Process

1.  **Discuss Changes**: Before making any changes, please open an issue to discuss the proposed changes.
2.  **Create a Branch**: Create a new branch for your changes.
3.  **Make Changes**: Make your changes in the new branch.
4.  **Update Documentation**: If you add or change any functionality, please update the `README.md` file and any relevant code comments.
5.  **Submit a Pull Request**: When your changes are complete, submit a pull request.

## Testing

There is no formal test suite for this project yet. When making changes, please manually test them in Microsoft Excel to ensure that they work as expected.

## Module Dependencies

- **`FalPlot`**: `FalLang`, `FalWork`, `FalArray`, `FalFile`
- **`FalCSV`**: `FalFile`, `FalArray`
- **`FalFile`**: `FALCore`, `FalLang`, `FalWork`, `FalArray`
- **`FalWork`**: `FalFile`
