Late Binding and Type Mismatches: VBScript's late binding can lead to runtime errors if an object's method or property doesn't exist or if a type mismatch occurs during a procedure call.  Example:

```vbscript
Set obj = CreateObject("Some.Object")
' No error checking if Some.Object exists
result = obj.NonExistentMethod()
```
This will fail if `Some.Object` or `NonExistentMethod` is not found.  Explicit error handling is crucial for robust VBScript.

Implicit Type Conversions and Unexpected Behavior: VBScript's automatic type conversion can cause subtle issues. For example, comparing a string to a number might not work as expected. 

```vbscript
if "10" = 10 then
  msgbox "Equal"
else
  msgbox "Not Equal"
end if
```
This will show "Not Equal." Always use explicit type conversions when comparing different data types.

Unhandled Exceptions: VBScript can throw runtime exceptions that can halt script execution unless handled with `On Error Resume Next` or `On Error GoTo`.  Improper handling can lead to unexpected termination and data loss. Example: 

```vbscript
On Error Resume Next
Set fs = CreateObject("Scripting.FileSystemObject")
file = fs.OpenTextFile("nonexistent.txt")
' Error is ignored;  better to have explicit error handling.
```
Use structured exception handling for more maintainable code.

Memory Leaks: Though less common in VBScript compared to languages with manual memory management, large loops or improper object handling can lead to memory issues, especially when creating many objects without releasing them.