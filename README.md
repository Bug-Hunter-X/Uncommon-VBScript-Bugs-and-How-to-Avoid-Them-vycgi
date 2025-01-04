# Uncommon VBScript Bugs and Solutions

This repository demonstrates some less frequently encountered bugs in VBScript and provides solutions to mitigate them.  VBScript, while simple, can still lead to unexpected behavior if not handled carefully.  This document illustrates issues related to late binding, type mismatches, exceptions, and potential memory leaks.

## Bugs Covered

* **Late Binding and Type Mismatches:** The lack of compile-time type checking makes late binding error-prone.  The provided examples showcase how to handle potential errors gracefully.
* **Implicit Type Conversion:**  VBScript's implicit type conversions can cause comparisons to yield unexpected results.  This example shows the correct way to avoid unexpected behavior.
* **Unhandled Exceptions:**  Improper exception handling can cause script termination.  The example details how to handle exceptions using `On Error Resume Next` in a structured way.
* **Memory Leaks (Potential):**  While less frequent, creating many objects without releasing them can lead to memory pressure.  The document highlights best practices for managing objects.

## Solutions

The accompanying VBScript files demonstrate proper error handling, explicit type conversion, and structured exception handling to address each of the mentioned issues.  This demonstrates how to write robust and reliable VBScript code.