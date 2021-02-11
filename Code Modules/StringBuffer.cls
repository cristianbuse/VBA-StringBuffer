VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' VBA StringBuffer
'''-------------------------------------------------
''' https://github.com/cristianbuse/VBA-StringBuffer
'''-------------------------------------------------
'''
''' Copyright (c) 2019 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to deal
''' in the Software without restriction, including without limitation the rights
''' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
''' copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in all
''' copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
''' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
''' SOFTWARE.
'''=============================================================================
''==============================================================================
'' Description:
''    * A 'Java-like' string buffer
'' The main methods are (both using the 'Mid' statement for speed):
''    * Append (much faster than using regular VBA concatenation - & operator)
''    * Insert
'' Other useful methods:
''    * Delete
''    * Replace
''    * Reset
''    * Reverse
''    * Substring
'' Retrieve value by calling property:
''    * Value (default class member - can be omitted)
'' Notes:
''    * The buffer's capacity should only be increased by using the
''      EnsureCapacity method
''==============================================================================

Option Explicit

'Internal BUFFER Struct
Private Type BUFFER_STRUCT
    text_ As String
    endIndex_ As Long
    capacity_ As Long
End Type

'Class members
Private m_buffer As BUFFER_STRUCT

'*******************************************************************************
'Append new text to buffer
'*******************************************************************************
Public Sub Append(ByRef textToAppend As String)
    Dim addedLength As Long: addedLength = VBA.Len(textToAppend)
    If addedLength = 0 Then Exit Sub
    Dim newEndIndex As Long: newEndIndex = m_buffer.endIndex_ + addedLength
    '
    'EnsureCapacity already checks for: newEndIndex > m_buffer.capacity_
    'When many appends are done (ex. 1 character at a time for a million times)
    '   then the extra 'If' is faster than exiting the scope of this function
    '   i.e. avoid pushing a new stack frame at the top of the call stack
    If newEndIndex > m_buffer.capacity_ Then EnsureCapacity newEndIndex
    '
    'Replace unused characters with the new text after the last used position
    'The Mid Statement can be used to quickly replace characters
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mid-statement
    Mid$(m_buffer.text_, m_buffer.endIndex_ + 1, addedLength) = textToAppend
    '
    'Store last character position
    m_buffer.endIndex_ = newEndIndex
End Sub

'*******************************************************************************
'Delete a portion of the text
'*******************************************************************************
Public Sub Delete(ByVal startIndex As Long, ByVal length_ As Long)
    'Validate Input
    If startIndex < 1 Or startIndex > m_buffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Delete", "Invalid startIndex"
    ElseIf length_ < 0 Then
        Err.Raise 5, TypeName(Me) & ".Delete", "Invalid length_"
    ElseIf length_ = 0 Then
        Exit Sub 'Nothing to delete
    End If
    '
    'Check if a simple shift of the endIndex would suffice
    If startIndex + length_ > m_buffer.endIndex_ Then
        'Ignoring characters that were marked for deletion
        m_buffer.endIndex_ = startIndex - 1
        Exit Sub
    End If
    '
    Dim shiftLength As Long
    '
    shiftLength = m_buffer.endIndex_ - startIndex - length_ + 1
    '
    'Shift Text Left
    Mid$(m_buffer.text_, startIndex, shiftLength) _
        = VBA.Mid$(m_buffer.text_, startIndex + length_, shiftLength)
    '
    'Update last character position
    m_buffer.endIndex_ = m_buffer.endIndex_ - length_
End Sub

'*******************************************************************************
'Extend buffer size if needed
'*******************************************************************************
Public Sub EnsureCapacity(ByVal minimumCapacity As Long)
    'Maximum string length allowed by VBA for a dynamic-length string
    Const MAX_CAPACITY As Long = &H7FFFFFFF '2,147,483,647 (dec)
    '
    If minimumCapacity > m_buffer.capacity_ Then
        Dim oldCapacity As Long: oldCapacity = m_buffer.capacity_
        '
        m_buffer.capacity_ = minimumCapacity * 2
        If m_buffer.capacity_ > MAX_CAPACITY Then m_buffer.capacity_ = MAX_CAPACITY
        '
        m_buffer.text_ = m_buffer.text_ & VBA.Space$(m_buffer.capacity_ - oldCapacity)
    End If
End Sub

'*******************************************************************************
'Insert new text into buffer
'*******************************************************************************
Public Sub Insert(ByVal startIndex As Long, ByRef textToInsert As String)
    'Validate Input
    If startIndex < 1 Or startIndex > m_buffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Insert", "Invalid startIndex"
    End If
    '
    Dim addedLength As Long: addedLength = VBA.Len(textToInsert)
    If addedLength = 0 Then Exit Sub 'Nothing to insert
    Dim newEndIndex As Long: newEndIndex = m_buffer.endIndex_ + addedLength
    Dim shiftLength As Long: shiftLength = m_buffer.endIndex_ - startIndex + 1
    '
    'EnsureCapacity already checks for: newEndIndex > m_buffer.capacity_
    'When many appends are done (ex. 1 character at a time for a million times)
    '   then the extra 'If' is faster than exiting the scope of this function
    '   i.e. avoid pushing a new stack frame at the top of the call stack
    If newEndIndex > m_buffer.capacity_ Then EnsureCapacity newEndIndex
    '
    'Shift Text Right
    Mid$(m_buffer.text_, startIndex + addedLength, shiftLength) _
        = VBA.Mid$(m_buffer.text_, startIndex, shiftLength)
    '
    'Replace unused characters with the new text starting at startIndex
    'The Mid Statement can be used to quickly replace characters
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mid-statement
    Mid$(m_buffer.text_, startIndex, addedLength) = textToInsert
    '
    'Update last character position
    m_buffer.endIndex_ = newEndIndex
End Sub

'*******************************************************************************
'Replace a portion of the buffer with a given text
'*******************************************************************************
Public Sub Replace(ByVal startIndex As Long, ByVal length_ As Long, ByRef replacementText As String)
    'Validate Input
    If startIndex < 1 Or startIndex > m_buffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Replace", "Invalid startIndex"
    ElseIf length_ < 0 Then
        Err.Raise 5, TypeName(Me) & ".Replace", "Invalid length_"
    ElseIf length_ = 0 Then
        Exit Sub 'Nothing to replace
    End If
    '
    Dim usedLength As Long
    '
    'Compute usable length
    If startIndex + length_ > m_buffer.endIndex_ + 1 Then
        usedLength = m_buffer.endIndex_ + 1 - startIndex
    Else
        usedLength = length_
    End If
    '
    'Replace characters with the new text
    'The Mid Statement can be used to quickly replace characters
    'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mid-statement
    Mid$(m_buffer.text_, startIndex, usedLength) = replacementText
End Sub

'*******************************************************************************
'Reset buffer members for (re)usage
'*******************************************************************************
Public Sub Reset()
    m_buffer.text_ = vbNullString
    m_buffer.endIndex_ = 0
    m_buffer.capacity_ = 0
End Sub

'*******************************************************************************
'Reverses the contained string
'*******************************************************************************
Public Sub Reverse()
    If m_buffer.endIndex_ > 0 Then
        Me.Replace 1, m_buffer.endIndex_, VBA.StrReverse(Me.Value)
    End If
End Sub

'*******************************************************************************
'Returns a substring
'*******************************************************************************
Public Function Substring(ByVal startIndex As Long, ByVal length_ As Long) As String
    'Validate Input
    If startIndex < 1 Or startIndex > m_buffer.endIndex_ Then
        Err.Raise 9, TypeName(Me) & ".Substring", "Invalid startIndex"
    ElseIf length_ < 0 Then
        Err.Raise 5, TypeName(Me) & ".Substring", "Invalid length_"
    ElseIf length_ = 0 Then
        Exit Function
    End If
    '
    Dim usedLength As Long
    '
    'Compute usable length
    If startIndex + length_ > m_buffer.endIndex_ + 1 Then
        usedLength = m_buffer.endIndex_ + 1 - startIndex
    Else
        usedLength = length_
    End If
    '
    Substring = VBA.Mid$(m_buffer.text_, startIndex, usedLength)
End Function

'===============================================================================
'Returns the capacity of the string i.e. total length of buffer
'===============================================================================
Public Property Get Capacity() As Long
    Capacity = m_buffer.capacity_
End Property

'===============================================================================
'Returns the length of the string i.e. total number of used characters
'===============================================================================
Public Property Get Length() As Long
    Length = m_buffer.endIndex_
End Property

'===============================================================================
'Get the Used String
'Default class member. 'strBuffer.Value' can be also called as 'strBuffer'
'Open class in a text editor to see: Attribute [procName].VB_UserMemId = 0
'===============================================================================
Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    If m_buffer.endIndex_ > 0 Then
        Value = VBA.Left$(m_buffer.text_, m_buffer.endIndex_)
    End If
End Property
