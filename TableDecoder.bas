Attribute VB_Name = "TableDecoder"
Option Explicit

'Takes sets of 4 characters (ASCII) and returns sets of 3 characters (binary)
Private Function DecodeString(ByVal InString As String, ByVal Bytes As Long) As String
    Dim OutString As String
    Dim i As Long
    Dim CharArray()
    Dim x0, x1, x2 As Long      'These are the chars that will be spit out
    Dim y0, y1, y2, y3 As Long  'These are what we got in
    
    ReDim CharArray(Len(InString))
    
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
        'DoEvents
    Next i
    If Len(OutString) > Bytes Then
        DecodeString = Left(OutString, Bytes)
    Else
        DecodeString = OutString
    End If
End Function

'Takes sets of 3 characters (binary) and returns sets of 4 characters (ASCII)
Private Function EncodeString(ByVal InString As String) As String
    Dim OutString As String
    
    Dim i As Integer
    
    Dim y0, y1, y2, y3 As Integer
    
    Dim x0, x1, x2 As Integer
    
    'It's Very Important to pad the InString to make the len a multiple of 3.
    'This can add 1 or 2 extra NULL character to the end of the file,
    'resulting in a different file size. No harm, it's for easier
    'implementation. We could chop the file back down, upon uudecoding...
    If Len(InString) Mod 3 <> 0 Then
        InString = InString & String(3 - Len(InString) Mod 3, Chr$(0))
    End If
    
    For i = 1 To Len(InString) Step 3
        x0 = Asc(Mid(InString, i, 1))
        x1 = Asc(Mid(InString, i + 1, 1))
        x2 = Asc(Mid(InString, i + 2, 1))
        
        y0 = (x0 \ 4 + 32)
        y1 = ((x0 Mod 4) * 16) + (x1 \ 16 + 32)
        y2 = ((x1 Mod 16) * 4) + (x2 \ 64 + 32)
        y3 = (x2 Mod 64) + 32
        
        If (y0 = 32) Then y0 = 96
        If (y1 = 32) Then y1 = 96
        If (y2 = 32) Then y2 = 96
        If (y3 = 32) Then y3 = 96
        
        OutString = OutString + Chr(y0) + Chr(y1) + Chr(y2) + Chr(y3)
    Next i
    EncodeString = OutString
End Function

Sub MakeUUETable()
    Dim i, j, k, l As Integer
    Dim FilePos As Long
    
    'ProgForm.PBar.Value = 0
    ProgForm.Show
    Open "C:\Code\UU Class\Test Files\Table.uue" For Binary Access Write As #1
    FilePos = 1
    For i = 33 To 96
        For j = 33 To 96
            For k = 33 To 96
                For l = 33 To 96
                    Put #1, FilePos, EncodeString(Chr(i) + Chr(j) + Chr(k) + Chr(l))
                    FilePos = FilePos + 4
                    DoEvents
                Next l
            Next k
        Next j
        ProgForm.PBar.Value = i
    Next i
    Close #1
    Unload ProgForm
End Sub

Sub CheckVals()
    Dim i As Long
    Dim InByte As String * 1
    Dim maxx, minn As Integer
    
    minn = 255: maxx = 0
    ProgForm.PBar.Value = 0
    ProgForm.Show
    Open "C:\Code\UU Class\Test Files\Table.uue" For Binary Access Read As #1
    ProgForm.PBar.Max = LOF(1)
    For i = 1 To LOF(1)
        Get #1, i, InByte
        'If Asc(InByte) > maxx Then maxx = Asc(InByte)
        If Asc(InByte) < minn Then minn = Asc(InByte)
        DoEvents
        ProgForm.PBar.Value = i
    Next i
    Close #1
    Unload ProgForm
    MsgBox "Max: " + Trim(Str(maxx)) + vbCrLf + "Min: " + Trim(Str(minn))
End Sub

Sub MakeArray()
    Dim DecodeArray(33 To 96, 33 To 96, 33 To 96, 33 To 96) As String * 3
    Dim i, j, k, l As Integer
    
    ProgForm.Show
    For i = 33 To 96
        For j = 33 To 96
            For k = 33 To 96
                For l = 33 To 96
                    DecodeArray(i, j, k, l) = DecodeString(Chr(i) + Chr(j) + Chr(k) + Chr(l), 3)
                    DoEvents
                Next l
            Next k
        Next j
        ProgForm.PBar.Value = i
    Next i
End Sub
