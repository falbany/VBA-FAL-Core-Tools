# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Created `FalPlot2.cls`, a new class with a powerful, extensible theming engine for creating modern, publication-quality charts.
- Implemented a robust, file-based theming system for `FalPlot2` with the ability to save and load themes.
- Added several predefined, integrated themes to `FalPlot2`, such as "IEEE-Publication" and "Modern-Dashboard".
- Added a new `CreateSummarySheet` function to `FalXls.bas` that creates a summary of all sheets (worksheets and charts) in a workbook.
- Standard project files: `.gitignore`, `CHANGELOG.md`, `LICENSE`, `CONTRIBUTING.md`.
- `FalXls.bas` module with a function to generate a worksheet summary and many other features.
- `README.md` files in `FALCore/Modules` and `FALCore/Classes` directories.
- Comprehensive documentation to all public methods and properties in `FalPlot.cls`.
- New methods to `FalPlot.cls` to cover all functionality from the original module.

### Changed
- Major refactoring of `FalPlot.cls` to improve error handling, performance, and naming consistency.
- Refactored the API of `FalPlot2.cls` to be more consistent and maintainable, including standardizing all `Create...` functions to return a Chart object.
- Merged `CreateChartSummarySheet` and `CreateSummarySheet` in `FalXls.bas` into a single, more capable function.
- Removed dependencies on `FalLang` and `FalFile` from `FalPlot.cls` and `FalPlot2.cls` to make them more self-contained.
- Refactored `FalPlot.bas` module into an object-oriented `FalPlot.cls` class. `FalPlot.bas` is now a backward-compatible wrapper.
- Renamed all modules in `FALCore/Modules/` to the `Fal<...>` convention.
- Renamed methods in `FalPlot.cls` to a consistent `PascalCase` convention.
- Translated the entire codebase, including comments and function names, from Spanish to English.

### Fixed
- An issue where `FalPlot.cls` was missing most of its properties and methods.
- An issue where internal calls in `FalPlot.cls` were not updated after method renaming.
- Corrected a bug where creating a new chart would overwrite an existing chart with the same name. The name uniqueness check now accounts for both worksheets and chart sheets.

### Deprecated

### Removed

### Security
