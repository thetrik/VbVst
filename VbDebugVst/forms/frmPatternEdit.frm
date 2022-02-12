VERSION 5.00
Begin VB.Form frmPatternEdit 
   BackColor       =   &H00404040&
   Caption         =   "Pattern editor (hold Ctrl to edit velocity)"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   Icon            =   "frmPatternEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   Begin VB.PictureBox picContainer 
      Align           =   1  'Align Top
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   497
      TabIndex        =   1
      Top             =   0
      Width           =   7455
      Begin VbDebugVst.ctlSwitch chkSnap 
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   120
         Width           =   1395
         _extentx        =   2461
         _extenty        =   450
         value           =   -1
         caption         =   "Snap"
      End
   End
   Begin VbDebugVst.ctlPianoRollExt ctlPianoRoll 
      Height          =   6315
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   7395
      _extentx        =   13044
      _extenty        =   11139
   End
End
Attribute VB_Name = "frmPatternEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmPatternEdit.frm - pattern editor
' // by The trick, 2022
' //

Option Explicit

Public Sub UpdatePlayback()
    
    ctlPianoRoll.PlaybackPos = (PlaybackPos - Int(PlaybackPos)) * 16
    
End Sub

Private Sub chkSnap_OnClick( _
            ByVal bNewValue As Boolean)
    ctlPianoRoll.Snap = bNewValue
End Sub

Private Sub ctlPianoRoll_PatternChanged()
    Dim cUC As ctlPianoRollExt
    
    Set cUC = ctlPianoRoll.Object
    
    g_tSong.tPattern = cUC.Pattern
    
End Sub

Private Sub Form_Resize()

    If Me.ScaleHeight > picContainer.ScaleHeight + 100 And Me.ScaleWidth > 200 Then
        ctlPianoRoll.Move 0, picContainer.ScaleHeight, Me.ScaleWidth, Me.ScaleHeight - picContainer.ScaleHeight
    End If
    
End Sub
