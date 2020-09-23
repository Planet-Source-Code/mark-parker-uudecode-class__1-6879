Attribute VB_Name = "ArrayEncode"
Option Explicit

Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (dst As Any, src As Any, ByVal cb As Long)

Public Sub StringToByteArray(ByVal StringIn As String, ByteArray() As Byte)
    Dim lBytes As Long
    
    If Len(StringIn) = 0 Then Exit Sub
    lBytes = Len(StringIn)
    ReDim ByteArray(lBytes - 1)
    
    RtlMoveMemory ByteArray(0), ByVal StringIn, lBytes
End Sub

Public Sub ByteArrayToString(ByteArray() As Byte, StringOut As String)
  Dim lBytes As Long

  If LBound(ByteArray) > 0 Then Exit Sub ' lBound MUST be 0
  lBytes = UBound(ByteArray) + 1
  StringOut = String$(lBytes, 0)
  
  RtlMoveMemory ByVal StringOut, ByteArray(0), lBytes
End Sub

Private Function ArrayDecodeString(ByVal InString As String, ByVal Bytes As Long) As String
    Dim OutString As String
    Dim i As Long
    Dim UnCodedArray() As Byte
    Dim CodedArray() As Byte
    
    StringToByteArray InString, UnCodedArray()
    ReDim CodedArray((Len(InString) / 4) * 3)
    
    For i = 0 To (Len(InString) / 4) - 1  'would be 0 to 14
        If (UnCodedArray(i * 4 + 0) = 96) Then UnCodedArray(i * 4 + 0) = 32         'If char is 96 then set to 32
        If (UnCodedArray(i * 4 + 1) = 96) Then UnCodedArray(i * 4 + 1) = 32
        If (UnCodedArray(i * 4 + 2) = 96) Then UnCodedArray(i * 4 + 2) = 32
        If (UnCodedArray(i * 4 + 3) = 96) Then UnCodedArray(i * 4 + 3) = 32
        
        CodedArray(i * 3 + 0) = ((UnCodedArray(i * 4 + 0) - 32) * 4) + ((UnCodedArray(i * 4 + 1) - 32) \ 16) 'Calculate the 3 chars
        CodedArray(i * 3 + 1) = ((UnCodedArray(i * 4 + 1) Mod 16) * 16) + ((UnCodedArray(i * 4 + 2) - 32) \ 4)
        CodedArray(i * 3 + 2) = ((UnCodedArray(i * 4 + 2) Mod 4) * 64) + (UnCodedArray(i * 4 + 3) - 32)
    Next i
    ByteArrayToString CodedArray(), OutString
    ArrayDecodeString = Left(OutString, Bytes)
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

Sub atest()
    Dim dec1 As String
    Dim dec2 As String
    
    dec1 = FastDecodeString("_P``_P```/__`/\```#_`/\`__\``/___P``````````````````````````", 45)
    dec2 = ArrayDecodeString("_P``_P```/__`/\```#_`/\`__\``/___P``````````````````````````", 45)
    
    If dec1 <> dec2 Then MsgBox "Decode Error!"
    
End Sub

Sub ADecodeTest()
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
        dummy = ArrayDecodeString("_P``_P```/__`/\```#_`/\`__\``/___P``````````````````````````", 45)
    Next i
    fasttime = Timer - TimeHolder
    
    MsgBox "Fast decode: " + Trim(Str(slowtime)) + vbCrLf + "Array decode: " + Trim(Str(fasttime))
End Sub

