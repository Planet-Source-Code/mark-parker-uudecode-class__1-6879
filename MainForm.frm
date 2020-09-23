VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UU Class Project Tester"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton DoProgDec 
      Caption         =   "Progressive Decode"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   1080
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar EncProg 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton DoDecode 
      Caption         =   "Decode"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton DoEncode 
      Caption         =   "Encode"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton PickOutFile 
      Caption         =   "..."
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton PickInFile 
      Caption         =   "..."
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox OutFile 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox InFile 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Text            =   "C:\Code\UU Class\Test Files\Aolback.uue"
      Top             =   120
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog Dialogs 
      Left            =   4440
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar DecProg 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label 
      Caption         =   "Decode Progress:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "Encode Progress:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Output Filename:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      Caption         =   "Input Filename:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents UUCoder As UUTools
Attribute UUCoder.VB_VarHelpID = -1

Private Sub DoDecode_Click()
    Dim timeholder As Double
    
    If InFile <> "" Then
        timeholder = Timer
        UUCoder.UUDecodeFile InFile.Text
        MsgBox Timer - timeholder
    End If
    InFile = ""
    DecProg.Value = 0
End Sub

Private Sub DoEncode_Click()
    Dim timeholder As Double
    
    If InFile <> "" And OutFile <> "" Then
        timeholder = Timer
        UUCoder.UUEncodeFile InFile.Text, OutFile.Text
        MsgBox Timer - timeholder
    End If
    InFile = ""
    OutFile = ""
    EncProg.Value = 0
End Sub

Private Sub DoProgDec_Click()
    Dim i As Long
    Dim TotalParts As Long
    Dim CurrentPos As Long
    Dim Remain As Long
    Dim Buffer As String
    
    Dim timeholder As Double
    
    Open InFile.Text For Binary Access Read Shared As #1
    TotalParts = LOF(1) \ 1000
    Remain = LOF(1) Mod 1000
    CurrentPos = 1
    Buffer = String(1000, 0)
    
    timeholder = Timer
    
    For i = 1 To TotalParts
        Buffer = String(1000, 0)
        Get #1, CurrentPos, Buffer
        UUCoder.ProgressiveDecode Buffer, "C:\Code\UU Class\Test Files\"
        CurrentPos = CurrentPos + 1000
    Next
    Buffer = String(Remain, 0)
    Get #1, CurrentPos, Buffer
    UUCoder.ProgressiveDecode Buffer, "C:\Code\UU Class\Test Files\"
    
    MsgBox "Total Time: " & Timer - timeholder
    
    Close #1
End Sub

Private Sub Form_Load()
    Set UUCoder = New UUTools
    EncProg.Min = 0
    EncProg.Max = 1000
    EncProg.Value = 0
    DecProg.Min = 0
    DecProg.Max = 1000
    DecProg.Value = 0
End Sub

Private Sub PickInFile_Click()
    Dialogs.Filename = ""
    Dialogs.ShowOpen
    InFile = Dialogs.Filename
End Sub

Private Sub PickOutFile_Click()
    Dialogs.Filename = ""
    Dialogs.ShowSave
    OutFile = Dialogs.Filename
End Sub

Private Sub UUCoder_DecodeProgress(ByVal Percent As Single, ByVal Total As Long)
    DecProg.Value = Percent * 1000
    DoEvents
End Sub

Private Sub UUCoder_EncodeProgress(ByVal Percent As Single, ByVal Total As Long)
    EncProg.Value = Percent * 1000
    DoEvents
End Sub
