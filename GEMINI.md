# Gemini Agent Instructions

This document provides instructions for the Gemini AI agent working with the FALCore VBA Suite.

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

1.  **Understand the Goal**: Before making any changes, make sure you understand the user's request.
2.  **Explore the Code**: Explore the existing code to understand how it works.
3.  **Make Changes**: Make your changes to the code.
4.  **Update Documentation**: If you add or change any functionality, please update the `README.md` file and any relevant code comments.
5.  **Verify Your Work**: Before submitting your changes, please verify that they work as expected. Since there is no automated test suite, you should manually review your changes.

## Submitting Your Work

When you are finished, please submit your work with a clear and concise commit message.
