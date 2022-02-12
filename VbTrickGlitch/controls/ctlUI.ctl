VERSION 5.00
Begin VB.UserControl ctlUI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   8535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11475
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   765
   Begin VbTrickGlitch.ctlSwitch swhLocked 
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   1440
      Width           =   2655
      _extentx        =   4683
      _extenty        =   450
      caption         =   "Locked:"
   End
   Begin VbTrickGlitch.ctlKnob knbSpeed 
      Height          =   1035
      Left            =   6120
      TabIndex        =   2
      Top             =   300
      Width           =   1035
      _extentx        =   1826
      _extenty        =   1826
      largechange     =   10
      smallchange     =   1
      max             =   100
   End
   Begin VbTrickGlitch.ctlSlots sltSlots 
      Height          =   1755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      _extentx        =   10716
      _extenty        =   3096
   End
   Begin VbTrickGlitch.ctlPianoRoll pnrEdit 
      Height          =   6735
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   11475
      _extentx        =   20241
      _extenty        =   11880
   End
   Begin VbTrickGlitch.ctlKnob knbBeats 
      Height          =   1035
      Left            =   7200
      TabIndex        =   4
      Top             =   300
      Width           =   1035
      _extentx        =   1826
      _extenty        =   1826
      zerovalue       =   1
      largechange     =   1
      smallchange     =   1
      max             =   8
      min             =   1
      value           =   4
   End
   Begin VbTrickGlitch.ctlKnob knbDivisions 
      Height          =   1035
      Left            =   8280
      TabIndex        =   6
      Top             =   300
      Width           =   1035
      _extentx        =   1826
      _extenty        =   1826
      zerovalue       =   1
      largechange     =   1
      smallchange     =   1
      max             =   8
      min             =   1
      value           =   4
   End
   Begin VbTrickGlitch.ctlKnob knbPitch 
      Height          =   1035
      Left            =   9360
      TabIndex        =   8
      Top             =   300
      Width           =   1035
      _extentx        =   1826
      _extenty        =   1826
      largechange     =   100
      smallchange     =   1
      max             =   1200
      min             =   -1200
   End
   Begin VbTrickGlitch.ctlSwitch swhSnap 
      Height          =   255
      Left            =   8820
      TabIndex        =   11
      Top             =   1440
      Width           =   2595
      _extentx        =   4577
      _extenty        =   450
      value           =   -1
      caption         =   "Snap:"
   End
   Begin VbTrickGlitch.ctlKnob knbSmooth 
      Height          =   1035
      Left            =   10440
      TabIndex        =   12
      Top             =   300
      Width           =   1035
      _extentx        =   1826
      _extenty        =   1826
      largechange     =   10
      smallchange     =   1
      max             =   100
   End
   Begin VB.Label lblSmooth 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Smooth:"
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
      Height          =   195
      Left            =   10440
      TabIndex        =   13
      Top             =   60
      Width           =   1035
   End
   Begin VB.Line linSeparator 
      BorderColor     =   &H00808080&
      X1              =   408
      X2              =   760
      Y1              =   92
      Y2              =   92
   End
   Begin VB.Label lblPitch 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Pitch:"
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
      Height          =   195
      Left            =   9360
      TabIndex        =   9
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label lblDivisions 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Divisions:"
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
      Height          =   195
      Left            =   8280
      TabIndex        =   7
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label lblBeats 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Beats:"
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
      Height          =   195
      Left            =   7200
      TabIndex        =   5
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Speed:"
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
      Height          =   195
      Left            =   6120
      TabIndex        =   3
      Top             =   60
      Width           =   1035
   End
End
Attribute VB_Name = "ctlUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // ctlUI.ctl - GUI for VbTrickGlitch plugin
' // by The trick, 2022
' //

Option Explicit

' // This variable is shared between control and class
Private m_tSharedData() As tSharedData
Private m_tSABankDesc   As SAFEARRAY1D

Public Property Get hWnd() As OLE_HANDLE
    hWnd = UserControl.hWnd
End Property

Public Sub SetBankShared( _
           ByVal pBank As PTR)
           
    With m_tSABankDesc
        .cbElements = Len(m_tSharedData(0))
        .cDims = 1
        .fFeatures = FADF_AUTO
        .rgsabound(0).cElements = 1
        .pvData = pBank
    End With
    
    PutMemPtr ByVal ArrPtr(m_tSharedData), VarPtr(m_tSABankDesc)
    
End Sub

' // Change UI state
' // Update UI according to current state
Public Sub StateChanged()
    Dim cPianoRoll  As ctlPianoRoll
    Dim dNormalPos  As Double
    
    On Error GoTo exit_proc
    
    With m_tSharedData(0)

        If .eChStateEffect And SCM_SLOT_CURRENT Then
            sltSlots.CurrentSlot = .lCurrentSlot
        End If
        
        If .eChStateEffect And SCM_SLOT_PLAYBACK Then
            sltSlots.PlaybackSlot = .lPlaybackSlot
        End If
        
        With .tPresets(.lCurPreset)
            
            If m_tSharedData(0).eChStateEffect And SCM_PROGRAM Then
            
                Set cPianoRoll = pnrEdit.object
                
                cPianoRoll.Pattern = .tSlots(m_tSharedData(0).lActiveSlot).tPattern
                knbPitch.Value = .tSlots(m_tSharedData(0).lActiveSlot).fPitch * 100
                knbSpeed.Value = .tSlots(m_tSharedData(0).lActiveSlot).fSpeed * 100
                knbSmooth.Value = .tSlots(m_tSharedData(0).lActiveSlot).fSmooth * 100
                
            Else
                If m_tSharedData(0).lCurrentSlot = sltSlots.ActiveSlot Then
                    
                    If (m_tSharedData(0).eChStateEffect And SCM_PITCH) Then
                        knbPitch.Value = .tSlots(m_tSharedData(0).lCurrentSlot).fPitch * 100
                    End If
                    
                    If m_tSharedData(0).eChStateEffect And SCM_SPEED Then
                        knbSpeed.Value = .tSlots(m_tSharedData(0).lCurrentSlot).fSpeed * 100
                    End If
                    
                    If m_tSharedData(0).eChStateEffect And SCM_SMOOTH Then
                        knbSmooth.Value = .tSlots(m_tSharedData(0).lCurrentSlot).fSmooth * 100
                    End If
                    
                End If
            End If
            
        End With
        
        ' // Update playback pos
        dNormalPos = (.dPlaybackPos - (Int(.dPlaybackPos / pnrEdit.Beats) * pnrEdit.Beats)) * 4
        pnrEdit.PlaybackPos = dNormalPos
        sltSlots.PlaybackPos = dNormalPos / (pnrEdit.Beats * 4)
        
exit_proc:

    End With
    
End Sub

Private Sub knbBeats_Changed( _
            ByVal lNewValue As Long)
            
    pnrEdit.Beats = lNewValue

    m_tSharedData(0).tPresets(m_tSharedData(0).lCurPreset).tSlots(sltSlots.ActiveSlot).tPattern.lLengthPerBeats = lNewValue
    ' // This is not automated

End Sub

Private Sub knbDivisions_Changed( _
            ByVal lNewValue As Long)
    pnrEdit.Divisions = lNewValue
    ' // This is not automated
End Sub

Private Sub knbPitch_Changed( _
            ByVal lNewValue As Long)

    m_tSharedData(0).tPresets(m_tSharedData(0).lCurPreset).tSlots(sltSlots.ActiveSlot).fPitch = lNewValue / 100
    
    If m_tSharedData(0).bRecordMode Then
        m_tSharedData(0).eChStateUI = m_tSharedData(0).eChStateUI Or SCM_PITCH
    End If

End Sub

Private Sub knbSmooth_Changed( _
            ByVal lNewValue As Long)
            
    m_tSharedData(0).tPresets(m_tSharedData(0).lCurPreset).tSlots(sltSlots.ActiveSlot).fSmooth = lNewValue / 100
    
    If m_tSharedData(0).bRecordMode Then
        m_tSharedData(0).eChStateUI = m_tSharedData(0).eChStateUI Or SCM_SMOOTH
    End If
    
End Sub

Private Sub knbSpeed_Changed( _
            ByVal lNewValue As Long)
            
    m_tSharedData(0).tPresets(m_tSharedData(0).lCurPreset).tSlots(sltSlots.ActiveSlot).fSpeed = lNewValue / 100
    
    If m_tSharedData(0).bRecordMode Then
        m_tSharedData(0).eChStateUI = m_tSharedData(0).eChStateUI Or SCM_SPEED
    End If

End Sub

Private Sub pnrEdit_PatternChanged()
    Dim cPianoRoll  As ctlPianoRoll
    
    Set cPianoRoll = pnrEdit.object

    m_tSharedData(0).tPresets(m_tSharedData(0).lCurPreset).tSlots(sltSlots.ActiveSlot).tPattern = cPianoRoll.Pattern
    m_tSharedData(0).tPresets(m_tSharedData(0).lCurPreset).tSlots(sltSlots.ActiveSlot).tPattern.lLengthPerBeats = pnrEdit.Beats

End Sub

' // User clicked on slot
Private Sub sltSlots_OnSlotChange( _
            ByVal lNewIndex As Long)
    Dim cPianoRoll  As ctlPianoRoll
    
    Set cPianoRoll = pnrEdit.object

    With m_tSharedData(0).tPresets(m_tSharedData(0).lCurPreset)
    
        knbPitch.Value = .tSlots(lNewIndex).fPitch * 100
        knbSpeed.Value = .tSlots(lNewIndex).fSpeed * 100
        knbBeats.Value = .tSlots(lNewIndex).tPattern.lLengthPerBeats
        knbSmooth.Value = .tSlots(lNewIndex).fSmooth * 100
        
        cPianoRoll.Pattern = .tSlots(lNewIndex).tPattern
        
    End With
    
    If m_tSharedData(0).bRecordMode Then
        m_tSharedData(0).eChStateUI = m_tSharedData(0).eChStateUI Or SCM_SLOT_ACTIVE
    End If
    
    m_tSharedData(0).lActiveSlot = lNewIndex
    m_tSharedData(0).lCurrentSlot = lNewIndex ' // Current slot is set in record mode
        
End Sub

Private Sub swhLocked_OnClick( _
            ByVal bNewValue As Boolean)
    pnrEdit.Locked = bNewValue
End Sub

Private Sub swhSnap_OnClick( _
            ByVal bNewValue As Boolean)
    pnrEdit.Snap = bNewValue
End Sub

Private Sub UserControl_Terminate()
    PutMemPtr ByVal ArrPtr(m_tSharedData), NULL_PTR
End Sub
