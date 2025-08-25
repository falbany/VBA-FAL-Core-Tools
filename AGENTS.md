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
