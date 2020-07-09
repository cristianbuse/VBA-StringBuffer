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
''    * A string buffer is like a String, but can be modified by using a couple
''      of useful exposed methods
'' The main methods are:
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

'*******************************************************************************
'API for copying a number of bytes (byteLength) in memory starting from the
'   Source Memory Address to the Destination Memory Address
'*******************************************************************************
#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (ByVal destinationAddress As LongPtr, ByVal sourceAddress As LongPtr, ByVal byteLength As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "usr/lib/libc.dylib" Alias "memmove" (ByVal destinationAddress As Long, ByVal sourceAddress As Long, ByVal byteLength As Long) As Long
    #End If
#Else 'Windows
    'https://msdn.microsoft.com/en-us/library/mt723419(v=vs.85).aspx
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destinationAddress As LongPtr, ByVal sourceAddress As LongPtr, ByVal byteLength As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal destinationAddress As Long, ByVal sourceAddress As Long, ByVal byteLength As Long)
    #End If
#End If

'Internal BUFFER Struct
Private Type BUFFER_STRUCT
    text_ As String
    endIndex_ As Long
    capacity_ As Long
    #If VBA7 Then
        pointer_ As LongPtr
    #Else
        pointer_ As Long
    #End If
End Type

'Class members
Private m_buffer As BUFFER_STRUCT
Private m_useCopyMemForLargeChunks As Boolean

'Shifting large chunks of text is faster through external DLL Functions that
'   operate directly in memory, after a certain size
'Note that the chosen value is not the optimal value in all cases but was chosen
'   based on billions of iterations with different scenarios
Private Const LARGE_CHUNK_SIZE  As Long = 8192

'The internal String is Unicode with 2 bytes per character
'Any shifting using CopyMemory will operate on the intenal String only
'   See Insert and Delete methods
Private Const BYTES_PER_CHAR As Long = 2

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Class Constructor
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()
    m_useCopyMemForLargeChunks = True
    Reset
End Sub

'*******************************************************************************
'Reset buffer members for (re)usage
'*******************************************************************************
Public Sub Reset()
    m_buffer.text_ = vbNullString
    m_buffer.endIndex_ = 0
    m_buffer.capacity_ = 0
    m_buffer.pointer_ = 0
End Sub

'*******************************************************************************
'Append new text to buffer
'Note that 'textToAppend' is considered Unicode
'*******************************************************************************
Public Sub Append(ByRef textToAppend As String)
    Dim addedLength As Long: addedLength = VBA.Len(textToAppend)
    If addedLength = 0 Then Exit Sub
    Dim newEndIndex As Long: newEndIndex = m_buffer.endIndex_ + addedLength
    '
    'EnsureCapacity already checks for: newEndIndex > m_buffer.capacity_
    'When many appends are done (ex. 1 character at a time for a million times)
    '   then the extra 'If' is faster than exiting the scope of this function
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
Public Sub Delete(ByRef startIndex As Long, ByRef length_ As Long)
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
    If m_useCopyMemForLargeChunks And shiftLength > LARGE_CHUNK_SIZE Then
        CopyMemory UnsignedAddition(m_buffer.pointer_, (startIndex - 1) * BYTES_PER_CHAR) _
                 , UnsignedAddition(m_buffer.pointer_, (startIndex + length_ - 1) * BYTES_PER_CHAR) _
                 , shiftLength * BYTES_PER_CHAR
    Else
        Mid$(m_buffer.text_, startIndex, shiftLength) _
            = VBA.Mid$(m_buffer.text_, startIndex + length_, shiftLength)
    End If
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
        m_buffer.pointer_ = VBA.StrPtr(m_buffer.text_) 'Likely to have changed
    End If
End Sub

'*******************************************************************************
'Insert new text into buffer
'Note that 'textToInsert' is considered Unicode
'*******************************************************************************
Public Sub Insert(ByRef startIndex As Long, ByRef textToInsert As String)
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
    If newEndIndex > m_buffer.capacity_ Then EnsureCapacity newEndIndex
    '
    'Shift Text Right
    If m_useCopyMemForLargeChunks And shiftLength > LARGE_CHUNK_SIZE Then
        CopyMemory UnsignedAddition(m_buffer.pointer_, (startIndex + addedLength - 1) * BYTES_PER_CHAR) _
                 , UnsignedAddition(m_buffer.pointer_, (startIndex - 1) * BYTES_PER_CHAR) _
                 , shiftLength * BYTES_PER_CHAR
    Else
        Mid$(m_buffer.text_, startIndex + addedLength, shiftLength) _
            = VBA.Mid$(m_buffer.text_, startIndex, shiftLength)
    End If
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
'Note that 'replacementText' is considered Unicode
'*******************************************************************************
Public Sub Replace(ByRef startIndex As Long, ByRef length_ As Long, ByRef replacementText As String)
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
Public Function Substring(ByRef startIndex As Long, ByRef length_) As String
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

'===============================================================================
'Get/Set the option of using the CopyMemory DLL functions to speed up shifting
'   large chunks of text (see Insert and Delete methods)
'===============================================================================
Public Property Get UseCopyMemoryForLargeChunks() As Boolean
    UseCopyMemoryForLargeChunks = m_useCopyMemForLargeChunks
End Property
Public Property Let UseCopyMemoryForLargeChunks(ByVal newValue As Boolean)
    m_useCopyMemForLargeChunks = newValue
End Property

'*******************************************************************************
'Unsigned Addition
'
'VBA does not allow the declaration of unsigned integers. The integers are
'   always signed and can store both positive and negative numbers.
'
'-------------------------------------------------
'Basic information on bits and bytes
'-------------------------------------------------
'Bit: a basic unit of information used in computing.
'   Can only have one of two values: 0 or 1
'Nibble: a set of 4 bits
'   Can have binary values from 0000 to 1111 (0 to 15 in decimal notation)
'   Ex. 1001 (bin) = 1*2^3 + 0*2^2 + 0*2^1 + 1*2^0 = 8 + 0 + 0 + 1 = 9 (dec)
'Byte: unit of digital information that consists of 8 bits (or 2 nibbles)
'   Can have binary values from 00000000 to 11111111 (0 to 255 in decimal)
'   In VBA a Byte is an unsigned type
'
'-------------------------------------------------
'Signed VBA Integer Data Types
'-------------------------------------------------
'Integer: 2 Bytes (16 bits)
'   Can store values from -32,768 to 32,767 (decimal)
'Long: 4 Bytes (32 bits)
'   Can store values from -2,147,483,648 to 2,147,483,647 (decimal)
'LongLong: 8 Bytes (or 64 bits)
'   Values from -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 (decimal)
'   Available only in x64 versions of Applications using VBA
'
'In VBA the "Two's complement" mathematical operation method is used to
'   represent both negative and positive numbers (signed numbers)
'   See: https://en.wikipedia.org/wiki/Two%27s_complement
'   Ex. for 0101 (5), we reverse bits and add 1 and we get 1011 (-5)
'In signed integers the left-most bit is used to indicate the sign, so, a value
'   of 1 in the left-most bit indicates a negative number and a value of 0 in
'   the left-most bit indicates a non-negative number (zero or positive)
'
'-------------------------------------------------
'Unsigned vs Signed example
'-------------------------------------------------
'A 2-byte unsigned Integer would have binary values from 0000 0000 0000 0000 to
'   1111 1111 1111 1111 (0 to 65,535 in decimal)
'A 2-byte signed Integer has binary values from 0000 0000 0000 0000 to
'   0111 1111 1111 1111 (positive 0 to 32,767 in decimal) and binary values from
'   1000 0000 0000 0000 to 1111 1111 1111 1111 (negative -32,768 to -1 in decimal)
'
'-------------------------------------------------
'Hexadecimal notation
'-------------------------------------------------
'Hex notation is a base-16 system where each digit is a value between 0 an 15
'In Decimal, each digit is between 0 an 9 and in Binary each digit is 0 or 1
'
'In hex notation values from 0 to 9 are the same as in Decimal but values
'   10 to 15 are written as A to F
'So, a nibble corresponds to a digit in hex (0 to F) and a byte can be written
'   as a 2 digit hex number with values from 00 to FF
'
'In VBA the hex numbers are prefixed by &H characters. 00 -> &H00; FF -> &HFF
'   Ex.: &H7E = 7*16^1 + 15*16^0 = 126 (dec) or 0111 1110 in binary
'
'Hex notation provides a very convenient way to write byte values
'
'If a 2-byte Integer Type would be unsigned then it's values 0 to 65,535 could be
'   written as &H0000 to &HFFFF but because the Integer Type is signed then it's
'   values -32,768 to 32,767 could be written in hex as follows:
'       1000000000000000 to 1111111111111111 (-32,768 to -1) as &H8000 to &HFFFF
'       0000000000000000 to 0111111111111111 (0 to 32,767) as &H0000 to &H7FFF
'
'-------------------------------------------------
'LongPtr
'-------------------------------------------------
'https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/longptr-data-type
'Not a true data type. It transforms to Long(x32) or LongLong(x64)
'
'Long:
'   -2,147,483,648 to -1 (dec) corresponds to &H80000000 to &HFFFFFFFF (hex)
'    0 to +2,147,483,647 (dec) corresponds to &H00000000 to &H7FFFFFFF (hex)
'LongLong:
'    0 to +9,223,372,036,854,775,807 corresponds to &H0000000000000000 to &H7FFFFFFFFFFFFFFF
'   -9,223,372,036,854,775,808 to -1 corresponds to &H8000000000000000 to &HFFFFFFFFFFFFFFFF
'   https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/longlong-data-type
'
'Notes:
'   To declare a value as Long in VBA the following suffix is used: &
'   To declare a value as LongLong in VBA the following suffix is used: ^
'   Example:
'       &H8000  (Integer) = -32,768
'       &H8000& (Long)    = +32,768
'       &H80000000  (Long) =  -2,147,483,648 (can't be Integer; "&" not needed)
'       &H80000000& (Long) =  -2,147,483,648
'       &H80000000^ (LongLong) =  +2,147,483,648 (64-bit platforms only)
'       &H8000000000000000 (LongLong) = will not compile!
'       &H8000000000000000^ (LongLong) = -9,223,372,036,854,775,808
'
'-------------------------------------------------
'Memory address
'-------------------------------------------------
'Memory addresses are fixed-length sequences of digits conventionally displayed
'   and manipulated as unsigned integers
'
'When given a memory address as a signed integer and a positive increment, in
'   order to find the correct address+increment, the 2 values must take
'   into account the limits described above.
'
'-------------------------------------------------
'Overflow Example
'-------------------------------------------------
'   Assume the 2 Long Integer numbers:
'       A = &H7FFFFFFD (hex) = 2147483645 (dec)
'       B = &H0000000C (hex) =         12 (dec)
'   If the 2 integers would be Unsigned then their sum would be:
'       S = A + B = &H7FFFFFFD + &H0000000C = &H80000009
'       or in decimal
'       S = A + B = 2147483645 +         12 = +2147483657
'   But because the 2 integers are Signed then their sum exceeds the limit
'       of 2,147,483,648 available in a Long data type
'   In VBA, the signed number S = &H80000009 = -2147483639
'
'-------------------------------------------------
'Aim of the function
'-------------------------------------------------
'The "UnsignedAddition" function avoids overflow errors as in the example above
'   by adding the minimum negative value as needed
'*******************************************************************************
#If VBA7 Then
Private Function UnsignedAddition(val1 As LongPtr, val2 As LongPtr) As LongPtr
#Else
Private Function UnsignedAddition(val1 As Long, val2 As Long) As Long
#End If
    'The minimum negative integer value of a Long Integer in VBA
    #If Win64 Then
    Const minNegative As LongLong = &H8000000000000000^ '-9,223,372,036,854,775,808 (dec)
    #Else
    Const minNegative As Long = &H80000000 '-2,147,483,648 (dec)
    #End If
    '
    If val1 > 0 Then
        If val2 > 0 Then
            'Overflow could occur
            If (val1 + minNegative + val2) < 0 Then
                'The sum will not overflow
                UnsignedAddition = val1 + val2
            Else
                'Example for Long data type (x32):
                '   &H7FFFFFFD + &H0000000C =  &H80000009
                '   2147483645 +         12 = -2147483639
                UnsignedAddition = val1 + minNegative + val2 + minNegative
            End If
        Else 'Val2 <= 0
            'Sum cannot overflow
            UnsignedAddition = val1 + val2
        End If
    Else 'Val1 <= 0
        If val2 > 0 Then
            'Sum cannot overflow
            UnsignedAddition = val1 + val2
        Else 'Val2 <= 0
            'Overflow could occur
            On Error GoTo ErrorHandler
            UnsignedAddition = val1 + val2
        End If
    End If
Exit Function
ErrorHandler:
    Err.Raise 6, TypeName(Me) & ".UnsignedAddition", "Overflow"
End Function
