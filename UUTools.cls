VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UUTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (dst As Any, src As Any, ByVal cb As Long)

Public Event DecodeProgress(ByVal Percent As Single, ByVal Total As Long)
Public Event EncodeProgress(ByVal Percent As Single, ByVal Total As Long)

Const PortionSize As Integer = 45
Const BufLen As Integer = 1024

Private DestFileName As String

'***********************************************************************************
'*                      Encoding/decoding wrappers begin here.                     *
'*                      --------------------------------------                     *
'* These are the public methods for the class. All other subs/functions should be  *
'* declared as private. These methods provide an easy interface to the main encode/*
'* decode functions found below.                                                   *
'*                                                                                 *
'*      ProgressiveDecode   : decodes a stream on-the-fly.                         *
'*          Returns TRUE if decoding is taking place,                              *
'*          Returns FALSE if not decoding (decoder is searching for an attachment) *
'*          Returns a filename if a file was just successfully decoded.            *
'*      UUDecodeFile        : decodes a uuencoded file.                            *
'*          Returns the number of files that were decoded.                         *
'*      UUEncodeFile        : encodes a binary to a uuencoded file.                *
'*          Returns TRUE if encoding was successful,                               *
'*          Returns FALSE if encoding was not successful                           *
'***********************************************************************************

Public Function ProgressiveDecode(ByVal Data As String, Optional ByVal DestDir As String = "") As Variant
    Static InAFile As Boolean
    Static CumulativeData As String
    Static CurFilePos As Long
    
    Static WriteBuffer As String
    
    Dim StringLine As String
    Dim FileNum As Long
    Dim Bytes As Long
    
    If InAFile = False Then
        While (LCase(Left(Data, 5) <> "begin")) And Len(Data) > 5  'Skip to the "begin"
            Data = Right(Data, Len(Data) - 1)
        Wend
        If LCase(Left(Data, 5)) <> "begin" Then      'See if a file was found
            ProgressiveDecode = False
            Exit Function
        End If
        FastSplitter Data, StringLine
        DestFileName = DestDir + HeaderFilename(StringLine) 'Ok, we're at a valid file
        CurFilePos = 1
        InAFile = True
    End If
    
    If InAFile Then
        CumulativeData = CumulativeData + Data          'Add new data to buffer
        While InStr(CumulativeData, Chr(13))            'We have enough to decode some
            FastSplitter CumulativeData, StringLine
            If LCase(Left(StringLine, 3)) = "end" Then  'Stop at the end of a file
                InAFile = False
                WriteToFile DestFileName, WriteBuffer, CurFilePos
                ProgressiveDecode = DestFileName
                DestFileName = ""
                CurFilePos = 1
                Exit Function
            End If
            Bytes = Asc(Left(StringLine, 1)) - 32       'Calc how many bytes to decode  (OPT 1)
            StringLine = Right(StringLine, Len(StringLine) - 1)
            If Len(StringLine) > 2 Then
                WriteBuffer = WriteBuffer + DecodeString(StringLine, Bytes)
                If Len(WriteBuffer) > BufLen Then            'Buffer up to BufLen bytes
                    WriteToFile DestFileName, WriteBuffer, CurFilePos
                End If
            End If
        Wend
        ProgressiveDecode = True
    End If
End Function

Public Function UUDecodeFile(ByVal InputFilename As String, Optional ByVal DestDir As String = "") As Integer
    Dim InString As String
    Dim OutString As String
    Dim OutFileName As String
    Dim CurrentPos As Long
    Dim WriteBuffer As String
    Dim FilesFound As Long
    Dim Bytes As Long
    Dim FNumIn As Long
    
    FNumIn = FreeFile
    Open InputFilename For Input As #FNumIn ' file opened for reading
    
    FilesFound = 0
    While Not EOF(1)
        Line Input #FNumIn, InString
        If Left(InString, 6) = "begin " Then
            FilesFound = FilesFound + 1
            If DestDir = "" Then DestDir = StripDirname(InputFilename)  'If an output path wasn't suppplied,
            OutFileName = DestDir + HeaderFilename(InString)            'put the output file in the current dir.
            CurrentPos = 1
            Line Input #FNumIn, InString
            While LCase(Left(InString, 3)) <> "end"
                Bytes = Asc(Left(InString, 1)) - 32     'Calc how many bytes to decode
                InString = Right(InString, Len(InString) - 1)
                If Len(InString) > 2 Then
                    WriteBuffer = WriteBuffer + DecodeString(InString, Bytes)
                    If Len(WriteBuffer) > BufLen Then            'Buffer up to BufLen bytes
                        RaiseEvent DecodeProgress(CurrentPos / LOF(1), CurrentPos)
                        WriteToFile OutFileName, WriteBuffer, CurrentPos
                    End If
                End If
                Line Input #FNumIn, InString
            Wend
            RaiseEvent DecodeProgress(CurrentPos / LOF(1), CurrentPos)
            WriteToFile OutFileName, WriteBuffer, CurrentPos
        End If
    Wend
    Close #FNumIn
    UUDecodeFile = FilesFound
End Function

Public Function UUEncodeFile(ByVal InputFilename As String, ByVal OutputFilename As String) As Boolean
    Dim TotalParts As Long
    Dim Remain As Long
    Dim InString As String
    Dim CurrentPos As Long
    Dim i As Long
    Dim FNumIn As Long
    Dim FNumOut As Long
    
    On Error GoTo ErrorHandler
    
    'Open the original file as binary read
    FNumIn = FreeFile
    Open InputFilename For Binary Access Read Shared As #FNumIn
    
    'Open the target file as binary write
    FNumOut = FreeFile
    Open OutputFilename For Binary Access Write As #FNumOut
    
    'Write the standard uuencode file header
    Put #FNumOut, , "begin 644 " + StripFilename(InputFilename) + vbCrLf

    'TotalParts equals the number of full sized chunks of "PortionSize" bytes.
    TotalParts = LOF(1) \ PortionSize

    'Remainder equals how many remaining bytes are at the end of the file.
    Remain = LOF(1) Mod PortionSize

    'CurrentPos is the current file position
    CurrentPos = 1

    'Set up InString as a buffer the size of PortionSize
    InString = String(PortionSize, 0)

    'For loop to read the portions one by one
    For i = 1 To TotalParts
        Get #FNumIn, CurrentPos, InString
        'use the ENC() for standard uuencode compatibility, pad "M"
        Put #FNumOut, , ENC(PortionSize) + EncodeString(InString) + vbCrLf
        CurrentPos = CurrentPos + PortionSize
        RaiseEvent EncodeProgress(CurrentPos / LOF(1), CurrentPos)
    Next
        
    'Set up InString again as a buffer the size of the remaining bytes
    InString = String(Remain, 0)
       
    'get the remaining bytes toward end of the file
    Get #FNumIn, CurrentPos, InString
        
    'get the remaining bytes size and calculate ENC() for the last line
    Put #FNumOut, , ENC(LOF(1) - CurrentPos + 1) + EncodeString(InString) + vbCrLf
    
    'put "end" for standard uuencode compatibility
    Put #FNumOut, , ENC(0) + vbCrLf + "end" + vbCrLf
    
    'All Done!
    Close #FNumIn
    Close #FNumOut
    UUEncodeFile = True
    Exit Function
    
ErrorHandler:
    Close #FNumIn
    Close #FNumOut
    UUEncodeFile = False
End Function

'***********************************************************************************
'*                    Main encode/decode functions begin here.                     *
'*                    ----------------------------------------                     *
'*                                                                                 *
'*      DecodeString        : decodes a set of 4 ASCII bytes to 3 binary bytes.    *
'*      EncodeString        : encodes a set of 3 binary bytes to 4 ASCII bytes.    *
'***********************************************************************************

'Takes sets of 4 characters (ASCII) and returns sets of 3 characters (binary)
Private Function DecodeString(ByVal InString As String, ByVal Bytes As Long) As String
    Dim OutString As String
    Dim i As Long
    Dim UnCodedArray() As Byte
    Dim CodedArray() As Byte
    
    'Convert the string to an array of bytes
    StringToByteArray InString, UnCodedArray()
    
    'Set up the array to recieve the decoded bytes
    ReDim CodedArray((Len(InString) / 4) * 3)
    
    For i = 0 To (Len(InString) / 4) - 1  'For i = 0 to 14
        'Check for characters that should be spaces and convert them back
        If (UnCodedArray(i * 4 + 0) = 96) Then UnCodedArray(i * 4 + 0) = 32
        If (UnCodedArray(i * 4 + 1) = 96) Then UnCodedArray(i * 4 + 1) = 32
        If (UnCodedArray(i * 4 + 2) = 96) Then UnCodedArray(i * 4 + 2) = 32
        If (UnCodedArray(i * 4 + 3) = 96) Then UnCodedArray(i * 4 + 3) = 32
        
        'Decode the 3 bytes
        CodedArray(i * 3 + 0) = ((UnCodedArray(i * 4 + 0) - 32) * 4) + ((UnCodedArray(i * 4 + 1) - 32) \ 16) 'Calculate the 3 chars
        CodedArray(i * 3 + 1) = ((UnCodedArray(i * 4 + 1) Mod 16) * 16) + ((UnCodedArray(i * 4 + 2) - 32) \ 4)
        CodedArray(i * 3 + 2) = ((UnCodedArray(i * 4 + 2) Mod 4) * 64) + (UnCodedArray(i * 4 + 3) - 32)
    Next i
    ByteArrayToString CodedArray(), OutString
    DecodeString = Left(OutString, Bytes)
End Function

'Takes sets of 3 characters (binary) and returns sets of 4 characters (ASCII)
Private Function EncodeString(ByVal InString As String) As String
    Dim OutString As String
    Dim i As Integer
    Dim UnCodedArray() As Byte
    Dim CodedArray() As Byte
    
    'Pad will null characters if necessary
    If Len(InString) Mod 3 <> 0 Then
        InString = InString & String(3 - Len(InString) Mod 3, Chr$(0))
    End If
    
    'Convert string to a byte array. This is MUCH faster than the Asc/Chr combo.
    StringToByteArray InString, UnCodedArray()
    
    'Make sure our output array is the correct size
    ReDim CodedArray(((Len(InString) / 3) * 4) - 1)
    
    For i = 0 To (Len(InString) / 3) - 1
        'Encode 4 bytes at a time
        CodedArray(i * 4 + 0) = UnCodedArray(i * 3 + 0) \ 4 + 32
        CodedArray(i * 4 + 1) = ((UnCodedArray(i * 3 + 0) Mod 4) * 16) + (UnCodedArray(i * 3 + 1) \ 16 + 32)
        CodedArray(i * 4 + 2) = ((UnCodedArray(i * 3 + 1) Mod 16) * 4) + (UnCodedArray(i * 3 + 2) \ 64 + 32)
        CodedArray(i * 4 + 3) = (UnCodedArray(i * 3 + 2) Mod 64) + 32
        
        'Check for spaces and eliminate them
        If CodedArray(i * 4 + 0) = 32 Then CodedArray(i * 4 + 0) = 96
        If CodedArray(i * 4 + 1) = 32 Then CodedArray(i * 4 + 1) = 96
        If CodedArray(i * 4 + 2) = 32 Then CodedArray(i * 4 + 2) = 96
        If CodedArray(i * 4 + 3) = 32 Then CodedArray(i * 4 + 3) = 96
    Next i
    ByteArrayToString CodedArray(), OutString
    EncodeString = OutString
End Function

'***********************************************************************************
'*                         Utility functions begin here.                           *
'***********************************************************************************

'Splits off the next line from Data into NextLine
Private Function FastSplitter(ByRef Data As String, ByRef NextLine As String)
    If InStr(Data, Chr(13)) Then  'There's a carriage return in the line
        If Left(Data, 1) = Chr(13) Then Data = Right(Data, Len(Data) - 1)   'kill any leading...
        If Left(Data, 1) = Chr(10) Then Data = Right(Data, Len(Data) - 1)
        NextLine = ""
        NextLine = Left(Data, InStr(Data, Chr(13)) - 1)
        Data = Right(Data, Len(Data) - InStr(Data, Chr(13)))
        If Left(Data, 1) = Chr(10) Then Data = Right(Data, Len(Data) - 1)
    End If
End Function

'Returns the filename part from a full filename.
Private Function StripFilename(ByVal Filename As String) As String
    While InStr(Filename, "\") <> 0
        Filename = Right(Filename, Len(Filename) - InStr(Filename, "\"))
    Wend
    StripFilename = Filename
End Function

'Returns the directory part from a filename.
Private Function StripDirname(ByVal Filename As String) As String
    While Right(Filename, 1) <> "\"
        Filename = Left(Filename, Len(Filename) - 1)
    Wend
    StripDirname = Filename
End Function

'Returns the filename from the header line of a uue file.
Private Function HeaderFilename(ByVal HeaderString As String) As String
    HeaderString = Right(HeaderString, Len(HeaderString) - Len("begin "))
    HeaderString = Right(HeaderString, Len(HeaderString) - InStr(HeaderString, " "))
    HeaderFilename = Trim(HeaderString)
End Function

'Writes data to a file from a buffer.
Private Sub WriteToFile(ByVal FileNam As String, ByRef Buffer As String, ByRef Position As Long)
    Dim FileNum As Long
    FileNum = FreeFile
    Open FileNam For Binary Access Write As #FileNum
    Put #FileNum, Position, Buffer
    Close #FileNum
    Position = Position + Len(Buffer)
    Buffer = ""
End Sub

'Converts a string to an array of bytes. Just like Asc() on every character.
Private Sub StringToByteArray(ByVal StringIn As String, ByteArray() As Byte)
    Dim lBytes As Long
    
    If Len(StringIn) = 0 Then Exit Sub
    lBytes = Len(StringIn)
    ReDim ByteArray(lBytes - 1)
    
    RtlMoveMemory ByteArray(0), ByVal StringIn, lBytes
End Sub

'Converts an array of bytes to a string.
Private Sub ByteArrayToString(ByteArray() As Byte, StringOut As String)
  Dim lBytes As Long

  If LBound(ByteArray) > 0 Then Exit Sub 'lBound MUST be 0
  lBytes = UBound(ByteArray) + 1
  StringOut = String$(lBytes, 0)
  
  RtlMoveMemory ByVal StringOut, ByteArray(0), lBytes
End Sub

'Gives a character representing the number of bytes to be decoded from a line. "M" is a full line (45 bytes)
Private Function ENC(ByVal i As Integer) As String
    If i = 0 Then
        ENC = "`"
    Else
        i = i + 32
        ENC = Chr(i)
    End If
End Function

