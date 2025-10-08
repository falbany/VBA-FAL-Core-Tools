# VBA Submodules

This directory contains several useful VBA libraries included as Git submodules. These modules provide enhanced functionality for common programming tasks.

## Included Modules

### 1. [VBA-Better-Array](https://github.com/Senipah/VBA-Better-Array)
A VBA class that provides a more powerful and flexible array-like object. It includes methods for sorting, slicing, splicing, and other common array manipulations, making it easier to work with collections of data.

### 2. [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)
A library for parsing and converting JSON data. It can transform a JSON string into VBA `Dictionary` and `Collection` objects, and can convert VBA data structures back into a JSON string. This is essential for working with web APIs.

### 3. [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)
A cross-platform (Windows and Mac) drop-in replacement for the `Scripting.Dictionary` object. It provides a familiar key-value store that is crucial for many data manipulation tasks, especially when `Scripting.Dictionary` is not available.

### 4. [VBA-Log](https://github.com/VBA-tools/VBA-Log)
A simple logging utility for VBA. It allows for logging messages at different levels (Debug, Info, Warn, Error) to the Immediate Window or to custom callback functions, which helps in debugging and monitoring applications.

### 5. [VBA-StringBuilder](https://github.com/retailcoder/VBA-StringBuilder)
A class designed for efficient string concatenation. It provides a `StringBuilder` object that is significantly faster than using the standard `&` operator when building large strings in a loop, improving performance.