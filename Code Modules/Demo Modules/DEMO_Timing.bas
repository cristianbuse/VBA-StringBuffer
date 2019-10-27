Attribute VB_Name = "DEMO_Timing"
Option Explicit
Option Private Module

Public Function TestAppendTimeNoBuffer(ByVal wordsCount As Long, wordLength As Long) As Variant
    'Check Input
    If wordsCount < 1 Then GoTo FailInput
    If wordLength < 1 Then GoTo FailInput
    If wordsCount * wordLength > &H7FFFFFFF Then GoTo FailInput
    '
    Dim i As Long
    Dim resultString As String
    Dim word As String: word = String$(wordLength, "A")
    Dim tStart As Date: tStart = TimeNow
    '
    For i = 1 To wordsCount
        resultString = resultString & word 'regular VB concatenation with new memory allocation
    Next i
    '
    TestAppendTimeNoBuffer = TimeToSeconds(TimeNow - tStart)
Exit Function
FailInput:
    TestAppendTimeNoBuffer = CVErr(xlErrValue)
End Function

Public Function TestAppendTimeWithBuffer(ByVal wordsCount As Long, wordLength As Long) As Variant
    'Check Input
    If wordsCount < 1 Then GoTo FailInput
    If wordLength < 1 Then GoTo FailInput
    If wordsCount * wordLength > 2147483647 Then GoTo FailInput
    '
    Dim i As Long
    Dim resultString As String
    Dim word As String: word = String$(wordLength, "A")
    Dim buff As New StringBuffer
    Dim tStart As Date: tStart = TimeNow
    '
    For i = 1 To wordsCount
        buff.Append word
    Next i
    resultString = buff 'or buff.Value
    Set buff = Nothing
    '
    TestAppendTimeWithBuffer = TimeToSeconds(TimeNow - tStart)
Exit Function
FailInput:
    TestAppendTimeWithBuffer = CVErr(xlErrValue)
End Function

Public Function TestInsertTimeWithCopyMemory(ByVal wordsCount As Long, wordLength As Long) As Variant
    'Check Input
    If wordsCount < 1 Then GoTo FailInput
    If wordLength < 1 Then GoTo FailInput
    If wordsCount * wordLength > 2147483647 Then GoTo FailInput
    '
    Dim i As Long
    Dim resultString As String
    Dim word As String: word = String$(wordLength, "A")
    Dim buff As New StringBuffer
    Dim tStart As Date: tStart = TimeNow
    '
    buff.Append "AAA"
    For i = 1 To wordsCount
        buff.Insert 2, word
    Next i
    Set buff = Nothing
    '
    TestInsertTimeWithCopyMemory = TimeToSeconds(TimeNow - tStart)
Exit Function
FailInput:
    TestInsertTimeWithCopyMemory = CVErr(xlErrValue)
End Function

Public Function TestInsertTimeWithoutCopyMemory(ByVal wordsCount As Long, wordLength As Long) As Variant
    'Check Input
    If wordsCount < 1 Then GoTo FailInput
    If wordLength < 1 Then GoTo FailInput
    If wordsCount * wordLength > 2147483647 Then GoTo FailInput
    '
    Dim i As Long
    Dim resultString As String
    Dim word As String: word = String$(wordLength, "A")
    Dim buff As New StringBuffer
    Dim tStart As Date: tStart = TimeNow
    '
    buff.Append "AAA"
    buff.UseCopyMemoryForLargeChunks = False 'The difference from TestInsertTimeWithCopyMemory
    For i = 1 To wordsCount
        buff.Insert 2, word
    Next i
    Set buff = Nothing
    '
    TestInsertTimeWithoutCopyMemory = TimeToSeconds(TimeNow - tStart)
Exit Function
FailInput:
    TestInsertTimeWithoutCopyMemory = CVErr(xlErrValue)
End Function

'*******************************************************************************
'Timing
'*******************************************************************************
Public Function TimeNow() As Date
    Dim t As Double
    '
    #If Mac Then
        Dim varTemp As Variant
        '
        varTemp = Evaluate("=Now()") 'Resolution of 0.01 seconds
        If VBA.IsError(varTemp) Then
            t = VBA.Now() 'Resolution of 1 second
        Else
            t = varTemp
        End If
        t = t - Int(t)
    #Else
        t = VBA.Timer / 86400
    #End If
    TimeNow = t
End Function
Public Function TimeToSeconds(ByVal time_ As Date) As Double
    'There are 86,400 seconds in a day (24h * 60m * 60s)
    'Convert from day fraction (time) to seconds
    TimeToSeconds = Round(time_ * 86400, 3)
End Function
