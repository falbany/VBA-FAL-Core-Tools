# Agent Instructions for VBA Submodules

This directory contains several VBA libraries managed as Git submodules. They provide powerful, reusable functionality that should be leveraged whenever possible to accelerate development and improve code quality.

## Initializing Submodules

If you have just cloned the repository, the submodule directories will be empty. You **must** run the following command in your bash terminal to initialize and fetch the content of all submodules:

```bash
git submodule update --init --recursive
```

## Available Submodules and Usage

Before writing new code, always consider if one of these submodules can solve your problem more efficiently.

### 1. [VBA-Better-Array](./VBA-Better-Array/)
- **Purpose**: Provides a `BetterArray` class that offers advanced array manipulation features not available in standard VBA arrays.
- **When to use**: Use this for complex array operations such as sorting, filtering, slicing, splicing, or joining. It is more powerful and often more readable than writing custom array logic.

### 2. [VBA-JSON](./VBA-JSON/)
- **Purpose**: A robust library for converting JSON strings to VBA objects (`Dictionary`, `Collection`) and vice-versa.
- **When to use**: This is essential for any task involving web APIs or reading/writing JSON configuration files. Use the `JsonConverter` module to handle all JSON operations.

### 3. [VBA-Dictionary](./VBA-Dictionary/)
- **Purpose**: A cross-platform (Windows/Mac) implementation of a dictionary (key-value store). It serves as a drop-in replacement for the Windows-only `Scripting.Dictionary`.
- **When to use**: **Always prefer this `Dictionary` class over `Scripting.Dictionary`** to ensure your code is compatible with all environments. It is a fundamental tool for data manipulation.

### 4. [VBA-Log](./VBA-Log/)
- **Purpose**: A simple but effective logging framework. It allows you to output log messages to the Immediate Window or custom listeners.
- **When to use**: Use this for debugging and tracing application flow. Implement `Logger.LogDebug` or `Logger.LogInfo` calls to track important events or variable states instead of relying solely on `Debug.Print`.

### 5. [VBA-StringBuilder](./VBA-StringBuilder/)
- **Purpose**: An efficient tool for building large strings.
- **When to use**: Whenever you are concatenating more than a few strings together, especially inside a loop. Using the `StringBuilder` class is significantly more performant than using the `&` operator repeatedly. The `clsMDM` class in this project already uses this pattern, which should be followed.