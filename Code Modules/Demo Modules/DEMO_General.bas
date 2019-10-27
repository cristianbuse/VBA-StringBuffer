Attribute VB_Name = "DEMO_General"
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "DEMO_General"

'*******************************************************************************
'Converts an array of Variants into a tab delimited text
'*******************************************************************************
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
    TabDelimitedTextFrom2DArray = buff.Value
Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & "." & methodName, "Invalid Array Value"
End Function

'*******************************************************************************
'Returns the Number of dimensions for an input array
'*******************************************************************************
Public Function GetDimensionsCount(inputArray As Variant) As Long
    Dim dimensionIndex As Long
    Dim dimensionBound As Long
    '
    'Increse the dimension index and loop until an error occurs
    On Error GoTo FinalDimension
    For dimensionIndex = 1 To 60000
       dimensionBound = LBound(inputArray, dimensionIndex)
    Next dimensionIndex
Exit Function
FinalDimension:
    GetDimensionsCount = dimensionIndex - 1
End Function

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
