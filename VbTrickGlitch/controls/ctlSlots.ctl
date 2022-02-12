VERSION 5.00
Begin VB.UserControl ctlSlots 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmrCaptureEdit 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1800
      Top             =   2400
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   675
      Left            =   1140
      MaxLength       =   16
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   1755
   End
End
Attribute VB_Name = "ctlSlots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // ctlSlots.ctl - patterns slots control
' // by The trick, 2022
' //

Option Explicit

Private Enum eSlotType
    ST_NORMAL       ' // Not active slot (for editing)
    ST_ACTIVE = 1   ' // Active slot for editing
    ST_PLAYBACK = 2 ' // Playback slot
    ST_CURRENT = 4  ' // Current slot in automation
End Enum

Private Const NUM_SLOTS_HORZ    As Long = 8
Private Const NUM_SLOTS_VERT    As Long = 5

Public Event OnSlotChange( _
             ByVal lNewIndex As Long)

Private m_lPlayBackSlot As Long
Private m_lActiveSlot   As Long
Private m_lCurrentSlot  As Long
Private m_fPlaybackPos  As Single
Private m_lSlotHeight   As Long
Private m_lSlotWidth    As Long
Private m_hClipRgn      As HANDLE
Private m_sSlotsNames() As String
Private m_lTextHeight   As Long

Public Property Get CurrentSlot() As Long
    CurrentSlot = m_lCurrentSlot
End Property
Public Property Let CurrentSlot( _
                    ByVal lIndex As Long)

    If lIndex = m_lCurrentSlot Then
        Exit Property
    End If

    RedrawSlotChangeState lIndex, ST_CURRENT Or GetSlotState(lIndex), _
                          m_lCurrentSlot, GetSlotState(m_lCurrentSlot) And Not ST_CURRENT
                          
    m_lCurrentSlot = lIndex
    
End Property

Public Property Get ActiveSlot() As Long
    ActiveSlot = m_lActiveSlot
End Property
Public Property Let ActiveSlot( _
                    ByVal lIndex As Long)

    If lIndex = m_lActiveSlot Then
        Exit Property
    End If

    RedrawSlotChangeState lIndex, ST_ACTIVE Or GetSlotState(lIndex), _
                          m_lActiveSlot, GetSlotState(m_lActiveSlot) And Not ST_ACTIVE
                          
    m_lActiveSlot = lIndex
    
End Property

Public Property Get PlaybackSlot() As Long
    PlaybackSlot = m_lPlayBackSlot
End Property
Public Property Let PlaybackSlot( _
                    ByVal lIndex As Long)

    If lIndex = m_lPlayBackSlot Then
        Exit Property
    End If

    RedrawSlotChangeState lIndex, ST_PLAYBACK Or GetSlotState(lIndex), _
                          m_lPlayBackSlot, GetSlotState(m_lPlayBackSlot) And Not ST_PLAYBACK
                          
    m_lPlayBackSlot = lIndex
    
End Property

Public Property Get PlaybackPos() As Single
    PlaybackPos = m_fPlaybackPos
End Property
Public Property Let PlaybackPos( _
                    ByVal fValue As Single)
     
    If fValue = m_fPlaybackPos Then
        Exit Property
    End If
    
    m_fPlaybackPos = fValue
    
    RedrawSlotChangeState m_lPlayBackSlot, GetSlotState(m_lPlayBackSlot), -1, 0

End Property

Public Sub Refresh()
    Dim hdc     As HANDLE
    Dim lIndex  As Long
    Dim lPosX   As Long
    Dim lPosY   As Long
    Dim eType   As eSlotType
    
    hdc = GetUCDC
    
    SetDCBrushColor hdc, &H404040
    
    PatBlt hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, vbPatCopy
    
    For lIndex = 0 To NUM_SLOTS_HORZ * NUM_SLOTS_VERT - 1
        
        lPosY = (lIndex \ NUM_SLOTS_HORZ) * m_lSlotHeight
        lPosX = (lIndex Mod NUM_SLOTS_HORZ) * m_lSlotWidth
        
        If lIndex = m_lActiveSlot Then
            eType = ST_ACTIVE
        Else
            eType = ST_NORMAL
        End If
        
        If lIndex = m_lPlayBackSlot Then
            eType = eType Or ST_PLAYBACK
        End If
        
        DrawSlot hdc, lPosX, lPosY, eType, m_sSlotsNames(lIndex)
        
    Next
    
    ReleaseUCDC hdc
    
    InvalidateRect UserControl.hWnd, ByVal 0&, 0
    
End Sub

Private Function GetSlotState( _
                 ByVal lIndex As Long) As eSlotType
                     
    If lIndex = m_lActiveSlot Then
        GetSlotState = ST_ACTIVE
    Else
        GetSlotState = ST_NORMAL
    End If
    
    If lIndex = m_lPlayBackSlot Then
        GetSlotState = GetSlotState Or ST_PLAYBACK
    End If
    
    If lIndex = m_lCurrentSlot Then
        GetSlotState = GetSlotState Or ST_CURRENT
    End If
    
End Function

Private Sub RedrawSlotChangeState( _
            ByVal lIndex1 As Long, _
            ByVal eState1 As eSlotType, _
            ByVal lIndex2 As Long, _
            ByVal eState2 As eSlotType)
    Dim tSlotArea   As RECT
    Dim hdc         As HANDLE
    
    hdc = GetUCDC
    
    If lIndex1 <> -1 Then
    
        tSlotArea = GetSlotRect(lIndex1)
        DrawSlot hdc, tSlotArea.Left, tSlotArea.Top, eState1, m_sSlotsNames(lIndex1)
        InvalidateRect UserControl.hWnd, tSlotArea, 0
        
    End If
    
    If lIndex2 <> -1 Then
    
        tSlotArea = GetSlotRect(lIndex2)
        DrawSlot hdc, tSlotArea.Left, tSlotArea.Top, eState2, m_sSlotsNames(lIndex2)
        InvalidateRect UserControl.hWnd, tSlotArea, 0
        
    End If
    
    ReleaseUCDC hdc
    
End Sub

Private Sub DrawSlot( _
            ByVal hdc As HANDLE, _
            ByVal lX As Long, _
            ByVal lY As Long, _
            ByVal eType As eSlotType, _
            ByRef sName As String)
    Dim lBackColor  As Long
    Dim lBackColor2 As Long
    Dim lForeColor  As Long
    Dim lPos        As Long
    Dim tRC         As RECT
    
    Select Case eType And ST_ACTIVE
    Case ST_NORMAL
        lBackColor = &H303030
        lBackColor2 = &H808080
        lForeColor = &HE0E0E0
    Case ST_ACTIVE
        lBackColor = &HC06080
        lBackColor2 = &HF09090
        lForeColor = &HE0E0E0
    End Select
    
    SelectObject hdc, GetStockObject(DC_BRUSH)
    SetDCBrushColor hdc, lBackColor
    SetDCPenColor hdc, lForeColor
    
    SetRect tRC, lX + 2, lY + 2, lX + m_lSlotWidth - 2, lY + m_lSlotHeight - 2
    
    If eType And ST_PLAYBACK Then
    
        lPos = m_fPlaybackPos * (tRC.Right - tRC.Left)

        If lPos > tRC.Right - tRC.Left Then
            lPos = tRC.Right - tRC.Left
        End If
        
        SelectClipRgn hdc, m_hClipRgn
        OffsetClipRgn hdc, tRC.Left, tRC.Top
                
        SetDCBrushColor hdc, lBackColor2
        
        If lPos > 0 Then
            PatBlt hdc, tRC.Left, tRC.Top, lPos, (tRC.Bottom - tRC.Top), vbPatCopy
        End If
        
        If lPos < m_lSlotWidth - 4 Then
            
            SetDCBrushColor hdc, lBackColor
            PatBlt hdc, tRC.Left + lPos, tRC.Top, tRC.Left + lPos + (tRC.Right - tRC.Left), (tRC.Bottom - tRC.Top), vbPatCopy
            
        End If
        
        SelectClipRgn hdc, 0
        SelectObject hdc, GetStockObject(NULL_BRUSH)
        
        SetDCPenColor hdc, &H80F0F0
        
        RoundRect hdc, tRC.Left, tRC.Top, tRC.Right, tRC.Bottom, 5, 5
        SelectObject hdc, GetStockObject(DC_BRUSH)
        
    Else
        RoundRect hdc, tRC.Left, tRC.Top, tRC.Right, tRC.Bottom, 5, 5
    End If
    
    If eType And ST_CURRENT Then
        
        SetDCPenColor hdc, &H4040FF
        SelectObject hdc, GetStockObject(NULL_BRUSH)
        RoundRect hdc, tRC.Left + 1, tRC.Top + 1, tRC.Right - 1, tRC.Bottom - 1, 3, 3
        
    End If
    
    OffsetRect tRC, 0, (m_lSlotHeight - 4 - m_lTextHeight) / 2
    DrawText hdc, sName, Len(sName), tRC, DT_CENTER Or DT_END_ELLIPSIS
    
End Sub

Private Sub EndNameEdit( _
            Optional bCancel As Boolean)
    
    If GetCapture = txtEdit.hWnd Then
        ReleaseCapture
    End If
    
    If Not bCancel Then
        m_sSlotsNames(m_lActiveSlot) = txtEdit.Text
    End If
    
    txtEdit.Visible = False
    RedrawSlotChangeState m_lActiveSlot, ST_ACTIVE Or GetSlotState(m_lActiveSlot), -1, 0
    ReleaseCaptureEditBox
    
End Sub

Private Sub tmrCaptureEdit_Timer()
    Dim tCurPos As POINT
    
    If GetFocus <> txtEdit.hWnd Then
        EndNameEdit
    Else
        
        If GetCapture <> UserControl.hWnd Then
            
            GetCursorPos tCurPos
            ScreenToClient UserControl.hWnd, tCurPos
            
            Select Case ChildWindowFromPointEx(UserControl.hWnd, tCurPos.x, tCurPos.y, CWP_SKIPDISABLED Or CWP_SKIPINVISIBLE)
            Case 0, UserControl.hWnd
                SetCapture UserControl.hWnd
            End Select
            
        End If
        
    End If
    
End Sub

Private Sub txtEdit_KeyPress( _
            ByRef KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        EndNameEdit True
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        EndNameEdit
        KeyAscii = 0
    End If
End Sub

Private Sub CaptureEditBox()
    tmrCaptureEdit.Enabled = True
End Sub

Private Sub ReleaseCaptureEditBox()

    If GetCapture = UserControl.hWnd Then
        ReleaseCapture
    End If
    
    tmrCaptureEdit.Enabled = False
    
End Sub

Private Sub txtEdit_Validate( _
            ByRef bCancel As Boolean)
    EndNameEdit
End Sub

Private Sub UserControl_DblClick()
    Dim tPt         As POINT
    Dim tSlotArea   As RECT
    
    GetCursorPos tPt
    ScreenToClient UserControl.hWnd, tPt
    
    tSlotArea = GetSlotRect(m_lActiveSlot)
    InflateRect tSlotArea, -2, -2
    
    If PtInRect(tSlotArea, tPt.x, tPt.y) = 0 Then Exit Sub
    
    ' // Edit name
    txtEdit.Move tSlotArea.Left, tSlotArea.Top, tSlotArea.Right - tSlotArea.Left, tSlotArea.Bottom - tSlotArea.Top
    txtEdit.Text = m_sSlotsNames(m_lActiveSlot)
    txtEdit.SelStart = 0
    txtEdit.SelLength = Len(m_sSlotsNames(m_lActiveSlot))
    txtEdit.Visible = True
    txtEdit.SetFocus
    
    CaptureEditBox
    
End Sub

Private Sub UserControl_Initialize()
    Dim lIndex  As Long
    
    ReDim m_sSlotsNames(NUM_SLOTS_HORZ * NUM_SLOTS_VERT - 1)
    
    For lIndex = 0 To UBound(m_sSlotsNames)
        m_sSlotsNames(lIndex) = "Slot " & CStr(lIndex)
    Next
    
End Sub

Private Sub UserControl_MouseDown( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    Dim tPt         As POINT
    Dim lSlotIndex  As Long
    Dim tSlotArea   As RECT
    
    tPt.x = fX
    tPt.y = fY
    
    If txtEdit.Visible Then
        EndNameEdit
    End If
    
    lSlotIndex = ((tPt.x \ m_lSlotWidth) Mod NUM_SLOTS_HORZ) + ((tPt.y \ m_lSlotHeight) * NUM_SLOTS_HORZ)
    
    If lSlotIndex < 0 Or lSlotIndex >= NUM_SLOTS_VERT * NUM_SLOTS_HORZ Then
        Exit Sub
    End If
    
    tSlotArea = GetSlotRect(lSlotIndex)
    
    If tPt.x >= tSlotArea.Left + 2 And tPt.x < tSlotArea.Right - 2 And _
        tPt.y >= tSlotArea.Top + 2 And tPt.y < tSlotArea.Bottom - 2 Then
        
        If lSlotIndex <> m_lActiveSlot Then

            RedrawSlotChangeState lSlotIndex, ST_ACTIVE Or GetSlotState(lSlotIndex), _
                                  m_lActiveSlot, GetSlotState(m_lActiveSlot) And Not ST_ACTIVE
                                  
            m_lActiveSlot = lSlotIndex
            
        End If
        
        RaiseEvent OnSlotChange(m_lActiveSlot)
        
    End If
    
End Sub

Private Function GetUCDC() As HANDLE

    GetUCDC = UserControl.hdc
    
    SaveDC GetUCDC
    
    SelectObject GetUCDC, GetStockObject(DC_PEN)
    SelectObject GetUCDC, GetStockObject(DC_BRUSH)
    
End Function

Private Sub ReleaseUCDC( _
            ByVal hdc As HANDLE)
    RestoreDC hdc, -1
End Sub

Private Function GetSlotRect( _
                 ByVal lIndex As Long) As RECT
    With GetSlotRect
        .Left = (lIndex Mod NUM_SLOTS_HORZ) * m_lSlotWidth
        .Top = (lIndex \ NUM_SLOTS_HORZ) * m_lSlotHeight
        .Right = .Left + m_lSlotWidth
        .Bottom = .Top + m_lSlotHeight
    End With
End Function

Private Sub UserControl_Resize()

    m_lSlotWidth = UserControl.ScaleWidth \ NUM_SLOTS_HORZ
    m_lSlotHeight = UserControl.ScaleHeight \ NUM_SLOTS_VERT
    
    m_lTextHeight = UserControl.TextHeight("0")
    
    If m_hClipRgn Then
        DeleteObject m_hClipRgn
    End If
    
    m_hClipRgn = CreateRoundRectRgn(0, 0, m_lSlotWidth - 3, m_lSlotHeight - 3, 5, 5)
    
    Refresh
    
End Sub

Private Sub UserControl_Terminate()
    If m_hClipRgn Then
        DeleteObject m_hClipRgn
    End If
End Sub
