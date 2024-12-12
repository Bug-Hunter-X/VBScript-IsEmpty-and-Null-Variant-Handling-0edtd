# VBScript IsEmpty and Null Variant Handling

This repository demonstrates a subtle bug in VBScript related to the `IsEmpty` function and how it interacts with `Null` Variant values. The issue occurs when `IsEmpty` is used to check for missing arguments, but those arguments might be `Null` instead of simply empty.  This can cause unexpected `Type mismatch` errors.

The `bug.vbs` file shows the problematic code. The `bugSolution.vbs` file shows the corrected code demonstrating how to robustly handle `Null` values.

## How to Reproduce

1.  Open `bug.vbs`
2.  Run the script and call the function `f` with at least one `Null` parameter. You should see a `Type mismatch` error.
3.  Open `bugSolution.vbs` and run this version. It should handle `Null` gracefully.