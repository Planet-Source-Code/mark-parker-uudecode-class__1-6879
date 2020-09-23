Attribute VB_Name = "TestModule"
Option Explicit

Function Result_Filename(ByVal Filename As String) As String
    Dim Length As Long
    Dim Temp_Str As String
    Dim i As Long
    Dim Done As Boolean
    Dim Left_Str As String
    Dim Right_Str As String
    
    Done = False
    
    Temp_Str = Filename
    
    Length = Len(Temp_Str)
    
    i = Length
    
    While Done = False And i <> 0
    Right_Str = Right(Temp_Str, 1)
    If Right_Str = " " Then
    Done = True
    End If
    Temp_Str = Left(Temp_Str, i - 1)
    i = i - 1
    Wend
    
    Result_Filename = Right(Filename, Length - Len(Temp_Str) - 1)
End Function

Function HeaderFilename(ByVal HeaderString As String) As String
    HeaderString = Trim(HeaderString)
    While InStr(HeaderString, " ") <> 0
        HeaderString = Right(HeaderString, InStr(HeaderString, " ") + 1)
    Wend
    HeaderFilename = Trim(HeaderString)
End Function

Sub Testcrlf()
If vbCrLf = Chr(13) + Chr(10) Then MsgBox "True"
End Sub


Function SplitOffNextLine(ByRef Data As String, ByRef NextLine As String)
    If InStr(Data, Chr(13)) Or InStr(Data, Chr(13) + Chr(10)) Then  'There's a cr(unix) or a crlf(dos)
        If Left(Data, 1) = Chr(13) Then Data = Right(Data, Len(Data) - 1)   'kill any leading...
        If Left(Data, 1) = Chr(10) Then Data = Right(Data, Len(Data) - 1)
        NextLine = ""
        While Left(Data, 1) <> Chr(13)
            NextLine = NextLine + Left(Data, 1)
            Data = Right(Data, Len(Data) - 1)
        Wend
        If Left(Data, 1) = Chr(13) Then Data = Right(Data, Len(Data) - 1)
        If Left(Data, 1) = Chr(10) Then Data = Right(Data, Len(Data) - 1)
    End If
End Function

Private Function FastSplitter(ByRef Data As String, ByRef NextLine As String)
    If InStr(Data, Chr(13)) Then  'There's a carriage return in the line
        NextLine = ""
        NextLine = Left(Data, InStr(Data, Chr(13)) - 1)
        Data = Right(Data, Len(Data) - InStr(Data, Chr(13)))
        If Left(Data, 1) = Chr(10) Then Data = Right(Data, Len(Data) - 1)
    End If
End Function

Sub testsplit()
    Dim teststr As String
    Dim timeholder As Single
    Dim timetotal As Single
    Dim theline As String
    Dim i As Integer
    
    For i = 1 To 32000
        teststr = String(25, ".") + "-" + vbCrLf + "-" + String(25, ".")
        timeholder = Timer
        SplitOffNextLine teststr, theline
        'FastSplitter teststr, theline
        'MsgBox "'" + teststr + "'"
        'MsgBox "'" + theline + "'"
        timetotal = timetotal + (Timer - timeholder)
    Next i
    MsgBox timetotal
End Sub

Sub ANDtest()
    Dim i, j As Integer
    
    i = 64
    j = 65
    
    MsgBox Str((i And j))
End Sub

Sub WriteVals()
    Dim i As Byte
    Dim j As Long
    
    Open "C:\Code\UU Class\Test Files\Priorit.uue" For Binary Access Read As #1
    Open "C:\Code\UU Class\Test Files\priorit.txt" For Output As #2
    
    For j = 1 To 100
        Get #1, , i
        Print #2, i
    Next j
    
    Close #1
    Close #2
End Sub
