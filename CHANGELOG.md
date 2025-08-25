# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Standard project files: `.gitignore`, `CHANGELOG.md`, `LICENSE`, `CONTRIBUTING.md`.
- `FalXls.bas` module with a function to generate a worksheet summary and many other features.
- `README.md` files in `FALCore/Modules` and `FALCore/Classes` directories.
- Comprehensive documentation to all public methods and properties in `FalPlot.cls`.
- New methods to `FalPlot.cls` to cover all functionality from the original module.

### Changed
- Refactored `FalPlot.bas` module into an object-oriented `FalPlot.cls` class. `FalPlot.bas` is now a backward-compatible wrapper.
- Renamed all modules in `FALCore/Modules/` to the `Fal<...>` convention.
- Renamed methods in `FalPlot.cls` to a consistent `PascalCase` convention.
- Translated the entire codebase, including comments and function names, from Spanish to English.

### Fixed
- An issue where `FalPlot.cls` was missing most of its properties and methods.
- An issue where internal calls in `FalPlot.cls` were not updated after method renaming.

### Changed

### Deprecated

### Removed

### Fixed

### Security
