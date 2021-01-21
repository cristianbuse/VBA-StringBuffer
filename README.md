# VBA-StringBuffer

A string buffer is like a String, but can be modified more easily using the built-in methods.
StringBuffer is a VBA Class that allows faster String Append/Concatenation than regular VBA concatenation (by using the ```Mid``` statement) and also provides other useful methods like Insert, Delete, Replace. The naming and the structure of the methods mimic a Java-like String Buffer.

## Installation

Just import the following code module in your VBA Project:

* **StringBuffer.cls**

## Usage
Create a new instance of the StringBuffer class and you are ready to use its methods:

```vba
Dim buff As New StringBuffer
```

Example1:
```vba
Sub BufferMethodsDemo()
    Dim buff As New StringBuffer
    '
    buff.Append "ABFGH"
    Debug.Print buff.Value 'ABFGH
    '
    buff.Insert 3, "CDE"
    Debug.Print buff 'ABCDEFGH
    '
    buff.Reverse
    Debug.Print buff 'HGFEDCBA
    '
    buff.Replace 2, 2, "XX"
    Debug.Print buff 'HXXEDCBA
    '
    buff.Reverse
    Debug.Print buff 'ABCDEXXH
    '
    buff.Delete 6, 2
    Debug.Print buff 'ABCDEH
    '
    Debug.Print buff.Substring(2, 3) 'BCD
End Sub
```

Example2 (see DEMO_General Module for full code):
```vba
Public Function TabDelimitedTextFrom2DArray(arr As Variant) As String
    Const methodName As String = "TabDelimitedTextFrom2DArray"
    '
    If GetDimensionsCount(arr) <> 2 Then
        Err.Raise 5, MODULE_NAME & "." & methodName, "Expected 2D Array"
    End If
    '
    Dim i As Long
    Dim j As Long
    Dim lowerRowBound As Long: lowerRowBound = LBound(arr, 1)
    Dim upperRowBound As Long: upperRowBound = UBound(arr, 1)
    Dim lowerColBound As Long: lowerColBound = LBound(arr, 2)
    Dim upperColBound As Long: upperColBound = UBound(arr, 2)
    Dim buff As New StringBuffer
    '
    On Error GoTo ErrorHandler
    For i = lowerRowBound To upperRowBound
        For j = lowerColBound To upperColBound
            If IsNull(arr(i, j)) Then
                buff.Append "NULL"
            Else
                buff.Append CStr(arr(i, j))
            End If
            If j < upperColBound Then buff.Append vbTab
        Next j
        If i < upperRowBound Then buff.Append vbNewLine
    Next i
    '
    TabDelimitedTextFrom2DArray = buff.StringValue
Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & methodName, "Invalid Array Value"
End Function
```

## Notes
* The StringBuffer is extremely useful for operations that involve lots of append operations. For example, it can append 1 character at a time for a million times in about 15 milliseconds on a Windows OS.
* You can download the available Demo Workbook. There is a Worksheet that allows you to test the Append speed and also another Worksheet with a few saved speed results for comparison.

## License
MIT License

Copyright (c) 2019 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.