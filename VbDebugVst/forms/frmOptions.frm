VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBlockSize 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   3345
      TabIndex        =   5
      Top             =   435
      Width           =   1500
   End
   Begin VB.TextBox txtSampleRate 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1725
      TabIndex        =   3
      Top             =   435
      Width           =   1500
   End
   Begin VB.TextBox txtTempo 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   105
      TabIndex        =   1
      Top             =   435
      Width           =   1500
   End
   Begin VbDebugVst.ctlButton cmdSave 
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   873
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00E0E0E0&
      Height          =   405
      Index           =   2
      Left            =   3330
      Top             =   420
      Width           =   1530
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00E0E0E0&
      Height          =   405
      Index           =   1
      Left            =   1710
      Top             =   420
      Width           =   1530
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   90
      Top             =   420
      Width           =   1530
   End
   Begin VB.Label lblBlockSize 
      BackColor       =   &H00404040&
      Caption         =   "Block Size:"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   3345
      TabIndex        =   4
      Top             =   75
      Width           =   1500
   End
   Begin VB.Label lblSampleRate 
      BackColor       =   &H00404040&
      Caption         =   "Sample Rate:"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1725
      TabIndex        =   2
      Top             =   75
      Width           =   1500
   End
   Begin VB.Label lblTempo 
      BackColor       =   &H00404040&
      Caption         =   "Tempo:"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   1500
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmOptions.frm - non-persistent options
' // by The trick, 2022
' //

Option Explicit

Private Sub cmdSave_Click()
    Dim dTempo      As Double
    Dim lBlockSize  As Long
    Dim lSampleRate As Long
    
    If StrToDbl(txtTempo.Text, dTempo) < 0 Then
        MsgBox "Invalid tempo value", vbCritical
        Exit Sub
    End If
    
    If StrToLng(txtBlockSize.Text, False, lBlockSize) < 0 Then
        MsgBox "Invalid blocksize value", vbCritical
        Exit Sub
    End If
    
    If StrToLng(txtSampleRate.Text, False, lSampleRate) < 0 Then
        MsgBox "Invalid blocksize value", vbCritical
        Exit Sub
    End If
    
    If dTempo < 60 Or dTempo > 240 Then
        MsgBox "Tempo value is out of range [60...240]", vbCritical
        Exit Sub
    End If
    
    If lBlockSize < 256 Or lBlockSize > 16384 Then
        MsgBox "BlockSize value is out of range [256...16384]", vbCritical
        Exit Sub
    End If
    
    If lSampleRate < 8000 Or lSampleRate > 48000 Then
        MsgBox "SampleRate value is out of range [8000...48000]", vbCritical
        Exit Sub
    End If
    
    If dTempo <> Tempo Then
        Tempo = dTempo
        frmSongEdit.UpdateAudio
    End If
    
    If lBlockSize <> BlockSize Then
        BlockSize = lBlockSize
    End If
    
    If lSampleRate <> SampleRate Then
        SampleRate = lSampleRate
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    txtTempo.Text = Tempo
    txtBlockSize.Text = BlockSize
    txtSampleRate.Text = SampleRate
    
End Sub
