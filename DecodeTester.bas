Attribute VB_Name = "DecodeTester"
Option Explicit

Dim asctime, checktime, calctime, chrtime As Single
Dim TimeHolder As Single

'Takes sets of 4 characters (ASCII) and returns sets of 3 characters (binary)
Private Function DecodeString(ByVal InString As String, ByVal Bytes As Long) As String
    Dim OutString As String
    Dim i As Long
    
    Dim x0, x1, x2 As Long   'These are the chars that will be spit out
    Dim y0, y1, y2, y3 As Long   'These are what we got in
    
    For i = 1 To Len(InString) Step 4
        y0 = Asc(Mid(InString, i, 1))           'Get 4 chars and put into 'y's
        y1 = Asc(Mid(InString, i + 1, 1))
        y2 = Asc(Mid(InString, i + 2, 1))
        y3 = Asc(Mid(InString, i + 3, 1))
        
        If (y0 = 96) Then y0 = 32               'If char is 96 then set to 32
        If (y1 = 96) Then y1 = 32
        If (y2 = 96) Then y2 = 32
        If (y3 = 96) Then y3 = 32
        
        x0 = ((y0 - 32) * 4) + ((y1 - 32) \ 16) 'Calculate the 3 chars
        x1 = ((y1 Mod 16) * 16) + ((y2 - 32) \ 4)
        x2 = ((y2 Mod 4) * 64) + (y3 - 32)
        
        Select Case Bytes
            Case 2
                OutString = OutString + Chr(x0) + Chr(x1)
            Case 1
                OutString = OutString + Chr(x0)
            Case Else
                OutString = OutString + Chr(x0) + Chr(x1) + Chr(x2)
        End Select
        Bytes = Bytes - 3
    Next i
    DecodeString = OutString
End Function

'Takes sets of 4 characters (ASCII) and returns sets of 3 characters (binary)
Private Function FastDecodeString(ByVal InString As String, ByVal Bytes As Long) As String
    Dim OutString As String
    Dim i As Long
    
    Dim x0, x1, x2 As Long   'These are the chars that will be spit out
    Dim y0, y1, y2, y3 As Long   'These are what we got in
    
    For i = 1 To Len(InString) Step 4
        y0 = Asc(Mid(InString, i, 1))           'Get 4 chars and put into 'y's
        y1 = Asc(Mid(InString, i + 1, 1))
        y2 = Asc(Mid(InString, i + 2, 1))
        y3 = Asc(Mid(InString, i + 3, 1))
        
        If (y0 = 96) Then y0 = 32               'If char is 96 then set to 32
        If (y1 = 96) Then y1 = 32
        If (y2 = 96) Then y2 = 32
        If (y3 = 96) Then y3 = 32
        
        x0 = ((y0 - 32) * 4) + ((y1 - 32) \ 16) 'Calculate the 3 chars
        x1 = ((y1 Mod 16) * 16) + ((y2 - 32) \ 4)
        x2 = ((y2 Mod 4) * 64) + (y3 - 32)
        
        OutString = OutString + Chr(x0) + Chr(x1) + Chr(x2)
    Next i
    If Len(OutString) > Bytes Then
        FastDecodeString = Left(OutString, Bytes)
    Else
        FastDecodeString = OutString
    End If
End Function

'Takes sets of 4 characters (ASCII) and returns sets of 3 characters (binary)
Private Function NewDecodeString(ByVal InString As String, ByVal Bytes As Long) As String
    Dim OutString As String
    Dim i As Long
    
    Dim x0, x1, x2 As Long   'These are the chars that will be spit out
    Dim y0, y1, y2, y3 As Long   'These are what we got in
    
    For i = 1 To Len(InString) Step 4
        y0 = Asc(Mid(InString, i, 1))           'Get 4 chars and put into 'y's
        y1 = Asc(Mid(InString, i + 1, 1))
        y2 = Asc(Mid(InString, i + 2, 1))
        y3 = Asc(Mid(InString, i + 3, 1))
        
        'If (y0 = 96) Then y0 = 32               'If char is 96 then set to 32
        'If (y1 = 96) Then y1 = 32
        'If (y2 = 96) Then y2 = 32
        'If (y3 = 96) Then y3 = 32
       
        'x0 = ((y0 - 32) * 4) + ((y1 - 32) \ 16) 'Calculate the 3 chars
        'x1 = ((y1 Mod 16) * 16) + ((y2 - 32) \ 4)
        'x2 = ((y2 Mod 4) * 64) + (y3 - 32)
        
        y0 = y0 - 32
        y1 = y1 - 32
        y2 = y2 - 32
        y3 = y3 - 32
        
        x0 = ((y0 And 63) * 2 ^ 2) Or ((y1 And 48) / 2 ^ 4)
        x1 = ((y1 And 15) * 2 ^ 4) Or ((y2 And 60) / 2 ^ 2)
        x2 = ((y2 And 3) * 2 ^ 6) Or ((y3 And 63))
        
        OutString = OutString + Chr(x0) + Chr(x1) + Chr(x2)
    Next i
    If Len(OutString) > Bytes Then
        NewDecodeString = Left(OutString, Bytes)
    Else
        NewDecodeString = OutString
    End If
End Function

Private Sub DecodeTest()
    Dim TimeHolder As Single
    Dim i As Integer
    Dim dummy As String
    Dim slowtime, fasttime As Single
    
    TimeHolder = Timer
    For i = 1 To 25000
        dummy = FastDecodeString("_P``_P```/__`/\```#_`/\`__\``/___P``````````````````````````", 45)
    Next i
    slowtime = Timer - TimeHolder
    
    TimeHolder = Timer
    For i = 1 To 25000
        dummy = NewDecodeString("_P``_P```/__`/\```#_`/\`__\``/___P``````````````````````````", 45)
    Next i
    fasttime = Timer - TimeHolder
    
    MsgBox "Old decode: " + Trim(Str(slowtime)) + vbCrLf + "New decode: " + Trim(Str(fasttime))
End Sub

Sub DecodeTimes()
    Dim i As Integer
    Dim dummy As String
    
    For i = 1 To 25000
        dummy = FastDecodeString("_P``_P```/__`/\```#_`/\`__\``/___P``````````````````````````", 45)
    Next i
    MsgBox "Asc() time: " + Trim(Str(asctime)) + vbCrLf + "Check time: " + Trim(Str(checktime)) + vbCrLf + "Calc time: " + Trim(Str(calctime)) + vbCrLf + "Chr() time: " + Trim(Str(chrtime))
End Sub
