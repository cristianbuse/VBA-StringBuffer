# VBA-StringBuffer

A string buffer is like a String, but can be modified.
StringBuffer is a VBA Class that allows faster String Append/Concatenation than regular VBA concatenation and also provides other useful methods like Insert, Delete, Replace.

## Installation

Just import the following code module in your VBA Project:

* **StringBuffer.cls**

## Usage
Create a new instance of the StringBuffer class and you are ready to use its methods:
```vba
Dim buff As New StringBuffer
```

Example1 (see DEMO_General Module for full code):
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

Example2:
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

## Notes
* The StringBuffer is extremely useful for operations that involve lots of append or insert operations. For example, it can append 1 character at a time for a million times in about 15 milliseconds on a Windows OS.
* You can download the available Demo Workbook. There are 2 Worksheets that allow you to test the Append and Insert speeds and also 2 Worksheets with a few saved speed results for comparison.

## License
Copyright (C) 2019 Cristian Buse

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see [http://www.gnu.org/licenses/](http://www.gnu.org/licenses/) or
[GPLv3](https://choosealicense.com/licenses/gpl-3.0/).