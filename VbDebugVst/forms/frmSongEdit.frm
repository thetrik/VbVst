VERSION 5.00
Begin VB.Form frmSongEdit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "Song editor"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
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
   Icon            =   "frmSongEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   825
End
Attribute VB_Name = "frmSongEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmSongEdit.frm - song editor
' // by The trick, 2022
' //

Option Explicit

' // 1st track - audio
' // 2nd track - pattern
' // 3rd track - events

Private Type tTrack
    sName       As String
    tPanelArea  As RECT
    tTrackArea  As RECT
End Type

Private m_tTracks(2)        As tTrack
Private m_tTimelineArea     As RECT
Private m_fBarWidth         As Single
Private m_cCurPencil        As StdPicture
Private m_cCurEraser        As StdPicture
Private m_lCurrentEvent     As Long
Private m_bDrawMode         As Boolean
Private m_tLastDrawPos      As Point
Private m_hBufBmp           As Handle
Private m_hBufDC            As Handle
Private m_lLastPlaybackPos  As Long
Private m_bLoaded           As Boolean

Public Sub ClearTrack()
    Dim lIndex  As Long
    
    g_tSong.tEvents(m_lCurrentEvent).bInit = False
    
    For lIndex = 0 To NUM_OF_BARS - 1
        RedrawTrackItem 2, lIndex
    Next
    
    InvalidateRect Me.hWnd, m_tTracks(2).tTrackArea, 0
    
End Sub

Public Sub SelectEvent( _
           ByVal lIndex As Long)
    m_lCurrentEvent = lIndex
    Redraw
End Sub

Public Sub UpdateAudio()
    If Not m_bLoaded Then
        Exit Sub
    Else
        Redraw
    End If
End Sub

Public Sub UpdatePluginInfo()
    If Not m_bLoaded Then
        Exit Sub
    Else
        Redraw
    End If
End Sub

Public Sub UpdatePlayback()
    Dim hdc     As Handle
    Dim lPos    As Long
    Dim tRC     As RECT
    
    If Not m_bLoaded Then
        Exit Sub
    End If
    
    hdc = Me.hdc
    
    SaveDC hdc
    SelectObject hdc, GetStockObject(DC_PEN)
    SetDCPenColor hdc, &H40A0A0
    lPos = BarsToPix(PlaybackPos)
    
    SetRect tRC, m_lLastPlaybackPos, 0, m_lLastPlaybackPos + 1, Me.ScaleHeight
    InvalidateRect Me.hWnd, tRC, 0
    
    RedrawPlayback hdc, lPos, True
    
    RestoreDC hdc, -1
    
    SetRect tRC, lPos, 0, lPos + 1, Me.ScaleHeight
    InvalidateRect Me.hWnd, tRC, 0
    
End Sub

Private Sub Redraw()
    Dim hdc As Handle
    Dim tRC As RECT
    
    hdc = Me.hdc
    
    SaveDC hdc
    
    SelectObject hdc, GetStockObject(DC_PEN)
    SelectObject hdc, GetStockObject(DC_BRUSH)
    
    SetRect tRC, 0, 0, Me.ScaleWidth, Me.ScaleHeight
    RedrawTracks hdc, tRC
    RedrawTimeLine hdc, tRC
    RedrawPlayback hdc, BarsToPix(PlaybackPos)
    RestoreDC hdc, -1
    
    InvalidateRect Me.hWnd, ByVal NULL_PTR, 0
    
End Sub

Private Sub RedrawPlayback( _
            ByVal hdc As Handle, _
            ByVal lPos As Long, _
            Optional ByVal bRestorePrev As Boolean)

    If bRestorePrev Then
        BitBlt hdc, m_lLastPlaybackPos, 0, 1, Me.ScaleHeight, m_hBufDC, 0, 0, vbSrcCopy
    End If
    
    m_lLastPlaybackPos = lPos

    BitBlt m_hBufDC, 0, 0, 1, Me.ScaleHeight, hdc, lPos, 0, vbSrcCopy
    
    MoveToEx hdc, lPos, 0, ByVal NULL_PTR
    LineTo hdc, lPos, Me.ScaleHeight
    
End Sub

Private Sub RedrawTimeLine( _
            ByVal hdc As Handle, _
            ByRef tRect As RECT)
    Dim lIndex      As Long
    Dim tArea       As RECT
    Dim tPos        As Point
    Dim sLbl        As String
    Dim lAvgSymW    As Long
    
    If IntersectRect(tArea, tRect, m_tTimelineArea) = 0 Then
        Exit Sub
    End If
    
    IntersectClipRect hdc, tArea.Left, tArea.Top, tArea.Right, tArea.Bottom
    
    SetDCBrushColor hdc, &H606060
    SetDCPenColor hdc, &HE0E0E0
    
    Rectangle hdc, m_tTimelineArea.Left, m_tTimelineArea.Top, m_tTimelineArea.Right, m_tTimelineArea.Bottom

    lIndex = Int((tArea.Left - m_tTimelineArea.Left) / m_fBarWidth)
    
    If lIndex = 0 Then
        lIndex = 1
    End If
    
    lAvgSymW = Me.TextWidth("0")
    
    Do While lIndex < NUM_OF_BARS
        
        tPos.x = Int(lIndex * m_fBarWidth) + m_tTimelineArea.Left
        tPos.y = m_tTimelineArea.Top
        
        If PtInRect(tArea, tPos.x, tPos.y) = 0 Then
            Exit Do
        End If
        
        sLbl = CStr(lIndex)
        
        TextOut hdc, tPos.x - lAvgSymW \ 2, tPos.y, sLbl, Len(sLbl)
        
        MoveToEx hdc, tPos.x, m_tTimelineArea.Bottom, ByVal NULL_PTR
        LineTo hdc, tPos.x, m_tTimelineArea.Bottom - 8
        
        lIndex = lIndex + 1
        
    Loop
    
    SelectClipRgn hdc, NULL_PTR
    
End Sub

Private Sub RedrawTracks( _
            ByVal hdc As Handle, _
            ByRef tRect As RECT)
    Dim tArea       As RECT
    Dim lTrackIndex As Long
    
    For lTrackIndex = 0 To UBound(m_tTracks)
    
        If IntersectRect(tArea, tRect, m_tTracks(lTrackIndex).tPanelArea) Then
            ' // Need redraw panel
            RedrawPanel hdc, lTrackIndex
        End If
        
        If IntersectRect(tArea, tRect, m_tTracks(lTrackIndex).tTrackArea) Then
            ' // Need redraw track
            RedrawTrack hdc, lTrackIndex, tArea
        End If
        
    Next
    
End Sub

Private Sub RedrawTrackItem( _
            ByVal lTrackIndex As Long, _
            ByVal lBarIndex As Long)
    Dim lBars   As Long
    Dim tRect   As RECT
    Dim hdc     As Handle
    
    If lTrackIndex = 0 Then
        lBars = -Int(-g_tSong.tAudioFile.dLenPerBars)
    Else
        lBars = 1
    End If
    
    tRect.Left = Int(lBarIndex * m_fBarWidth) + m_tTracks(lTrackIndex).tTrackArea.Left
    tRect.Top = m_tTracks(lTrackIndex).tTrackArea.Top
    tRect.Right = tRect.Left + -Int(-m_fBarWidth * lBars)
    tRect.Bottom = m_tTracks(lTrackIndex).tTrackArea.Bottom
    
    hdc = Me.hdc
    
    SaveDC hdc
    SelectObject hdc, GetStockObject(DC_PEN)
    SelectObject hdc, GetStockObject(DC_BRUSH)
    
    RedrawTrack hdc, lTrackIndex, tRect
    
    RestoreDC hdc, -1
    
    InvalidateRect Me.hWnd, tRect, 0
    
End Sub

Private Sub RedrawTrack( _
            ByVal hdc As Handle, _
            ByVal lIndex As Long, _
            ByRef tRect As RECT)
    Dim tArea       As RECT
    Dim lBarIndex   As Long
    Dim tPos        As Point
    
    With m_tTracks(lIndex)
        
        If IntersectRect(tArea, tRect, .tTrackArea) = 0 Then
            Exit Sub
        End If
        
        IntersectClipRect hdc, tArea.Left, tArea.Top, tArea.Right, tArea.Bottom
        
        SetDCBrushColor hdc, &H404040
        SetDCPenColor hdc, &HE0E0E0
        
        PatBlt hdc, tArea.Left, tArea.Top, tArea.Right - tArea.Left, tArea.Bottom - tArea.Top, vbPatCopy

        If tArea.Bottom >= .tTrackArea.Bottom Then
            MoveToEx hdc, .tTrackArea.Left, .tTrackArea.Top, ByVal NULL_PTR
            LineTo hdc, .tTrackArea.Right, .tTrackArea.Top
        End If
        
        ' // Draw bars
        lBarIndex = Int((tArea.Left - m_tTimelineArea.Left) / m_fBarWidth)
        
        If lBarIndex = 0 Then
            lBarIndex = 1
        End If
        
        Do While lBarIndex < NUM_OF_BARS
            
            tPos.x = Int(lBarIndex * m_fBarWidth) + m_tTimelineArea.Left
            tPos.y = m_tTimelineArea.Bottom
            
            If tPos.x > tArea.Right Then
                Exit Do
            End If
            
            If (lBarIndex Mod 4) = 0 Then
                SetDCPenColor hdc, &H808080
            Else
                SetDCPenColor hdc, &H606060
            End If
            
            MoveToEx hdc, tPos.x, tPos.y, ByVal NULL_PTR
            LineTo hdc, tPos.x, tArea.Bottom
            
            lBarIndex = lBarIndex + 1
            
        Loop
        
        Select Case lIndex
        Case 0: RedrawAudioEvents hdc, tArea
        Case 1: RedrawMidiEvents hdc, tArea
        Case 2: RedrawAutomation hdc, tArea
        End Select
        
        SelectClipRgn hdc, NULL_PTR
        
    End With
    
End Sub

Private Sub RedrawAutomationRange( _
            ByVal lX1 As Long, _
            ByVal lX2 As Long)
    Dim tRect   As RECT
    Dim hdc     As Handle
    Dim lTemp   As Long
    
    If lX2 < lX1 Then
    
        lTemp = lX1
        lX1 = lX2
        lX2 = lTemp
        
    End If
        
    SetRect tRect, lX1, m_tTracks(2).tTrackArea.Top, lX2 + 1, m_tTracks(2).tTrackArea.Bottom
    
    hdc = Me.hdc
    
    SaveDC hdc
    SelectObject hdc, GetStockObject(DC_PEN)
    SelectObject hdc, GetStockObject(DC_BRUSH)
    
    RedrawAutomation hdc, tRect
    
    RestoreDC hdc, -1
    
    InvalidateRect Me.hWnd, tRect, 0
    
End Sub

Private Sub RedrawAutomation( _
            ByVal hdc As Handle, _
            ByRef tRect As RECT)
    Dim lX          As Long
    Dim lY          As Long
    Dim lIndex      As Long
    Dim tArea       As RECT
    Dim dValue      As Double
    Dim lHeight     As Long
    Dim lColor      As Long
    Dim lNextBarPos As Long
    Dim lBarIndex   As Long
    
    If Not g_tSong.tEvents(m_lCurrentEvent).bInit Then
        Exit Sub
    End If
    
    If IntersectRect(tArea, tRect, m_tTracks(2).tTrackArea) = 0 Then
        Exit Sub
    End If
    
    lHeight = (m_tTracks(2).tTrackArea.Bottom - m_tTracks(2).tTrackArea.Top) - 2
    
    lBarIndex = Int((tArea.Left - m_tTimelineArea.Left) / m_fBarWidth)
    lNextBarPos = lBarIndex * m_fBarWidth
    
    If lNextBarPos < tArea.Left Then
        lBarIndex = lBarIndex + 1
        lNextBarPos = Int(lBarIndex * m_fBarWidth)
    End If
    
    For lX = tArea.Left To tArea.Right - 1
        
        lIndex = Int((lX - m_tTracks(2).tTrackArea.Left) / m_fBarWidth * 4 * EVENTS_QUANTIZATION)
        dValue = g_tSong.tEvents(m_lCurrentEvent).dEvents(lIndex)
        lY = m_tTracks(2).tTrackArea.Bottom - dValue * lHeight
        
        MoveToEx hdc, lX, m_tTracks(2).tTrackArea.Top + 1, ByVal NULL_PTR
        
        If (lX - m_tTimelineArea.Left) = lNextBarPos Then
        
            lBarIndex = lBarIndex + 1
            lNextBarPos = Int(lBarIndex * m_fBarWidth)
            SetDCPenColor hdc, &H606060
            
        Else
            SetDCPenColor hdc, &H404040
        End If
        
        LineTo hdc, lX, lY
        
        lColor = RGB(dValue * &H50 + &H40, dValue * &H50 + &H40, dValue * &H50 + &HA0)
        
        SetDCPenColor hdc, lColor
        
        LineTo hdc, lX, m_tTracks(2).tTrackArea.Bottom
        
    Next

End Sub

Private Sub RedrawMidiEvents( _
            ByVal hdc As Handle, _
            ByRef tRect As RECT)
    Dim lStartBar       As Long
    Dim lNumOfBars      As Long
    Dim lBarIndex       As Long
    Dim lBarHeight      As Long
    Dim lBarPosX        As Long
    
    If IsRectEmpty(tRect) Then
        Exit Sub
    End If
    
    SetDCBrushColor hdc, &H6060A0
    SetDCPenColor hdc, &H9090F0
    IntersectClipRect hdc, tRect.Left, tRect.Top, tRect.Right, tRect.Bottom
    
    lStartBar = Int((tRect.Left - m_tTimelineArea.Left) / m_fBarWidth)
    lNumOfBars = -Int(-(tRect.Right - m_tTimelineArea.Left) / m_fBarWidth) - lStartBar + 1
    
    If lStartBar + lNumOfBars > NUM_OF_BARS Then
        lNumOfBars = NUM_OF_BARS - lStartBar
    End If
    
    lBarHeight = m_tTracks(1).tTrackArea.Bottom - m_tTracks(1).tTrackArea.Top
    
    For lBarIndex = lStartBar To lStartBar + lNumOfBars - 1
        
        If Not g_tSong.bMidiTrack(lBarIndex) Then
            GoTo continue
        End If
        
        lBarPosX = Int(m_fBarWidth * lBarIndex) + m_tTimelineArea.Left

        Rectangle hdc, lBarPosX, m_tTracks(1).tTrackArea.Top, -Int(-lBarPosX - m_fBarWidth), m_tTracks(1).tTrackArea.Top + lBarHeight
continue:

    Next
    
    SelectClipRgn hdc, NULL_PTR
    
End Sub

Private Sub RedrawAudioEvents( _
            ByVal hdc As Handle, _
            ByRef tRect As RECT)
    Dim lStartBar       As Long
    Dim lNumOfBars      As Long
    Dim lBarIndex       As Long
    Dim lStartSample    As Long ' // Sample means both left+right if present
    Dim lNumOfSamples   As Long
    Dim lStartEvent     As Long
    Dim lSampleOffset   As Long
    Dim bHasEvent       As Boolean
    Dim lBarPosX        As Long
    Dim lBarWidth       As Long
    Dim bDrawEnd        As Boolean
    Dim lPixPos         As Long
    Dim lPixPosY        As Long
    Dim lSampleIndex    As Long
    Dim lSampleValue    As Long
    Dim lSampleStep     As Long
    Dim lBarHeight      As Long
    
    If IsRectEmpty(tRect) Then
        Exit Sub
    End If
    
    SetDCBrushColor hdc, &HA06060
    SetDCPenColor hdc, &HF09090
    
    lStartBar = Int((tRect.Left - m_tTimelineArea.Left) / m_fBarWidth)
    lNumOfBars = Int((tRect.Right - m_tTimelineArea.Left) / m_fBarWidth) - lStartBar + 1
    lBarHeight = m_tTracks(0).tTrackArea.Bottom - m_tTracks(0).tTrackArea.Top
    
    If lStartBar + lNumOfBars > NUM_OF_BARS Then
        lNumOfBars = NUM_OF_BARS - lStartBar
    End If
    
    If lStartBar > 0 Then
        lStartEvent = GetAudioEventStartIndexInMap(lStartBar)
    Else
        lStartEvent = -1
    End If
    
    For lBarIndex = lStartBar To lStartBar + lNumOfBars - 1
        
        If g_tSong.bAudioTrack(lBarIndex) Then
            lStartSample = 0
            lStartEvent = lBarIndex
        ElseIf lStartEvent >= 0 Then
            lStartSample = BarsToSamples(lBarIndex - lStartEvent)
        Else
            GoTo continue
        End If
        
        If lStartSample >= g_tSong.tAudioFile.lNumOfSamples Then
            lStartEvent = -1
            GoTo continue
        End If
        
        lBarPosX = Int(m_fBarWidth * lBarIndex) + m_tTimelineArea.Left
        
        If tRect.Left > lBarPosX Then
            lSampleOffset = BarsToSamples((tRect.Left - lBarPosX) / m_fBarWidth)
            lBarPosX = tRect.Left
        Else
            lSampleOffset = 0
        End If

        lStartSample = lStartSample + lSampleOffset
        
        If tRect.Right > lBarPosX + m_fBarWidth Then
            lNumOfSamples = BarsToSamples(1)
        Else
            lNumOfSamples = BarsToSamples((tRect.Right - lBarPosX) / m_fBarWidth)
        End If
        
        If lNumOfSamples + lStartSample >= g_tSong.tAudioFile.lNumOfSamples Then
            lNumOfSamples = g_tSong.tAudioFile.lNumOfSamples - lStartSample
            bDrawEnd = True
            lBarWidth = SamplesToBars(lNumOfSamples) * m_fBarWidth
        Else
            bDrawEnd = False
            lBarWidth = -Int(-m_fBarWidth)
        End If
        
        If lBarWidth = 0 Then
            GoTo continue
        End If
        
        PatBlt hdc, lBarPosX, m_tTracks(0).tTrackArea.Top, lBarWidth, lBarHeight, vbPatCopy
        
        If lStartSample = 0 Then
            MoveToEx hdc, lBarPosX, m_tTracks(0).tTrackArea.Top, ByVal NULL_PTR
            LineTo hdc, lBarPosX, m_tTracks(0).tTrackArea.Bottom
        End If

        If bDrawEnd Then
            MoveToEx hdc, lBarPosX + lBarWidth, m_tTracks(0).tTrackArea.Top, ByVal NULL_PTR
            LineTo hdc, lBarPosX + lBarWidth, m_tTracks(0).tTrackArea.Bottom
        End If
        
        MoveToEx hdc, lBarPosX, m_tTracks(0).tTrackArea.Top, ByVal NULL_PTR
        LineTo hdc, lBarPosX + lBarWidth, m_tTracks(0).tTrackArea.Top
        MoveToEx hdc, lBarPosX, m_tTracks(0).tTrackArea.Bottom - 1, ByVal NULL_PTR
        LineTo hdc, lBarPosX + lBarWidth, m_tTracks(0).tTrackArea.Bottom - 1
        
        With g_tSong.tAudioFile
            
            lSampleStep = lNumOfSamples \ lBarWidth
            lSampleIndex = lStartSample
            
            For lPixPos = 1 To lBarWidth - 1

                If .bIsMono Then
                
                    lSampleValue = Abs(.fSamples(lSampleIndex) * (lBarHeight \ 2 - 3))
                    
                    If lSampleValue = 0 Then
                        lSampleValue = 1
                    End If
                    
                    lPixPosY = m_tTracks(0).tTrackArea.Top + (lBarHeight - lSampleValue) \ 2
                    MoveToEx hdc, lBarPosX + lPixPos, lPixPosY, ByVal NULL_PTR
                    LineTo hdc, lBarPosX + lPixPos, lPixPosY + lSampleValue
                    
                Else
                    
                    lSampleValue = Abs(.fSamples(lSampleIndex * 2) * (lBarHeight \ 2 - 3))
                    
                    If lSampleValue = 0 Then
                        lSampleValue = 1
                    End If
                    
                    lPixPosY = m_tTracks(0).tTrackArea.Top + (lBarHeight \ 4) - (lSampleValue \ 2)
                    MoveToEx hdc, lBarPosX + lPixPos, lPixPosY, ByVal NULL_PTR
                    LineTo hdc, lBarPosX + lPixPos, lPixPosY + lSampleValue
                    
                    lSampleValue = Abs(.fSamples(lSampleIndex * 2 + 1) * (lBarHeight \ 2 - 3))
                    
                    If lSampleValue = 0 Then
                        lSampleValue = 1
                    End If
                    
                    lPixPosY = m_tTracks(0).tTrackArea.Top + (lBarHeight \ 4) * 3 - (lSampleValue \ 2)
                    MoveToEx hdc, lBarPosX + lPixPos, lPixPosY, ByVal NULL_PTR
                    LineTo hdc, lBarPosX + lPixPos, lPixPosY + lSampleValue
                    
                End If
                
                lSampleIndex = lSampleIndex + lSampleStep
                
            Next
        
        End With
        
continue:

    Next
    
End Sub

Private Sub RedrawPanel( _
            ByVal hdc As Handle, _
            ByVal lIndex As Long)
    Dim tLblRect    As RECT
    Dim sName       As String
    
    SetDCBrushColor hdc, &H606060
    SetDCPenColor hdc, &HE0E0E0
    
    With m_tTracks(lIndex)
        
        If lIndex = 2 Then
            
            sName = .sName & vbNewLine & "(" & g_tSong.tEvents(m_lCurrentEvent).sName & ")"
            
            If Not g_tSong.tEvents(m_lCurrentEvent).bCanBeAutomated Then
                sName = .sName & vbNewLine & "[Not automated]"
            End If
            
        Else
            sName = .sName
        End If
        
        Rectangle hdc, .tPanelArea.Left, .tPanelArea.Top, .tPanelArea.Right, .tPanelArea.Bottom + 1
        tLblRect = .tPanelArea
        DrawText hdc, sName, Len(sName), tLblRect, DT_CENTER Or DT_CALCRECT
        OffsetRect tLblRect, (.tPanelArea.Right - .tPanelArea.Left - (tLblRect.Right - tLblRect.Left)) \ 2, _
                             (.tPanelArea.Bottom - .tPanelArea.Top - (tLblRect.Bottom - tLblRect.Top)) \ 2
        DrawText hdc, sName, Len(sName), tLblRect, DT_CENTER
    
    End With
    
End Sub


Private Function PixToBars( _
                 ByVal lX As Long) As Double
                     
    If lX < m_tTimelineArea.Left Then
        PixToBars = 0
    ElseIf lX >= m_tTimelineArea.Right Then
        PixToBars = NUM_OF_BARS - 1
    Else
        PixToBars = (lX - m_tTimelineArea.Left) / m_fBarWidth
    End If
                     
End Function

Private Function BarsToPix( _
                 ByVal dPos As Double) As Long
                     
    If dPos > NUM_OF_BARS Then
        BarsToPix = m_tTimelineArea.Right
    ElseIf dPos <= 0 Then
        BarsToPix = m_tTimelineArea.Left
    Else
        BarsToPix = Int(dPos * m_fBarWidth) + m_tTimelineArea.Left
    End If
                     
End Function

Private Sub Form_Load()
    
    m_hBufDC = CreateCompatibleDC(Me.hdc)
    m_hBufBmp = CreateCompatibleBitmap(Me.hdc, 1, Screen.Height / Screen.TwipsPerPixelY)
    SaveDC m_hBufDC
    SelectObject m_hBufDC, m_hBufBmp
    
    Set m_cCurPencil = LoadResPicture(101, vbResCursor)
    Set m_cCurEraser = LoadResPicture(102, vbResCursor)
    
    m_tTracks(0).sName = "Audio track"
    m_tTracks(1).sName = "MIDI track"
    m_tTracks(2).sName = "Events"
    
    m_bLoaded = True
    
End Sub

Private Sub Form_MouseDown( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    Dim tPt     As Point
    Dim lIndex  As Long
    
    tPt.x = fX: tPt.y = fY
    
    If iButton = vbLeftButton Then
        
        If PtInRect(m_tTimelineArea, tPt.x, tPt.y) <> 0 Then
        
            PlaybackPos = PixToBars(tPt.x)
            UpdatePlayback
        
        ElseIf PtInRect(m_tTracks(0).tTrackArea, tPt.x, tPt.y) <> 0 Then
            
            If g_tSong.tAudioFile.lNumOfSamples = 0 Then
                Exit Sub
            End If
            
            lIndex = Int(PixToBars(tPt.x))
            
            If Not g_tSong.bAudioTrack(lIndex) Then
                
                g_tSong.bAudioTrack(lIndex) = True
                RedrawTrackItem 0, lIndex
                
            End If
            
            Set Me.MouseIcon = m_cCurPencil
            Me.MousePointer = vbCustom
            
        ElseIf PtInRect(m_tTracks(1).tTrackArea, tPt.x, tPt.y) <> 0 Then
                
            lIndex = Int(PixToBars(tPt.x))
            
            If Not g_tSong.bMidiTrack(lIndex) Then
                
                g_tSong.bMidiTrack(lIndex) = True
                RedrawTrackItem 1, lIndex
                
            End If
            
            Set Me.MouseIcon = m_cCurPencil
            Me.MousePointer = vbCustom
            
        ElseIf PtInRect(m_tTracks(2).tTrackArea, tPt.x, tPt.y) <> 0 Then
            
            If g_tSong.tEvents(m_lCurrentEvent).bCanBeAutomated Then
                
                DrawEventLine tPt.x, tPt.y, tPt.x, tPt.y
                RedrawAutomationRange tPt.x, tPt.x + 1
                m_tLastDrawPos.x = tPt.x
                m_tLastDrawPos.y = tPt.y
                m_bDrawMode = True
                
            End If
            
        End If
        
    ElseIf iButton = vbRightButton Then
    
        If PtInRect(m_tTracks(0).tTrackArea, tPt.x, tPt.y) <> 0 Then
            
            lIndex = GetAudioEventStartIndexInMap(PixToBars(tPt.x))
            
            If lIndex >= 0 Then
                
                g_tSong.bAudioTrack(lIndex) = False
                RedrawTrackItem 0, lIndex
                
            End If
            
            Set Me.MouseIcon = m_cCurEraser
            Me.MousePointer = vbCustom
            
        ElseIf PtInRect(m_tTracks(1).tTrackArea, tPt.x, tPt.y) <> 0 Then
                
            lIndex = Int(PixToBars(tPt.x))
            
            If g_tSong.bMidiTrack(lIndex) Then
                
                g_tSong.bMidiTrack(lIndex) = False
                RedrawTrackItem 1, lIndex
                
            End If
            
            Set Me.MouseIcon = m_cCurEraser
            Me.MousePointer = vbCustom
            
        ElseIf PtInRect(m_tTracks(2).tTrackArea, tPt.x, tPt.y) <> 0 Then
                
            If g_tSong.tEvents(m_lCurrentEvent).bCanBeAutomated Then
            
                EraseEventLine tPt.x, tPt.x
                RedrawAutomationRange tPt.x, tPt.x + 1
                m_tLastDrawPos.x = tPt.x
                m_tLastDrawPos.y = tPt.y
                m_bDrawMode = True
                
                Set Me.MouseIcon = m_cCurEraser
                Me.MousePointer = vbCustom

            End If
            
        End If
        
    End If
    
End Sub

Private Sub Form_MouseMove( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    Dim tPt     As Point
    
    tPt.x = fX: tPt.y = fY
    
    If m_bDrawMode Then
        
        If iButton = vbLeftButton Then
        
            DrawEventLine m_tLastDrawPos.x, m_tLastDrawPos.y, tPt.x, tPt.y
            RedrawAutomationRange m_tLastDrawPos.x, tPt.x
            m_tLastDrawPos.x = tPt.x
            m_tLastDrawPos.y = tPt.y
        
        ElseIf iButton = vbRightButton Then
            EraseEventLine m_tLastDrawPos.x, tPt.x
            RedrawAutomationRange m_tLastDrawPos.x, tPt.x
        Else
            Exit Sub
        End If
  
    ElseIf PtInRect(m_tTracks(0).tTrackArea, tPt.x, tPt.y) <> 0 Or _
       PtInRect(m_tTracks(1).tTrackArea, tPt.x, tPt.y) <> 0 Or _
       PtInRect(m_tTracks(2).tTrackArea, tPt.x, tPt.y) <> 0 Then

        If iButton = vbRightButton Then
            Set Me.MouseIcon = m_cCurEraser
        Else
            Set Me.MouseIcon = m_cCurPencil
        End If
        
        Me.MousePointer = vbCustom
    
    Else
        Me.MousePointer = vbDefault
    End If
    
End Sub

Private Sub Form_MouseUp( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    Dim tPt As Point
    
    tPt.x = fX: tPt.y = fY
    
    m_bDrawMode = False
    
    If iButton = vbRightButton Then
        If PtInRect(m_tTracks(2).tPanelArea, tPt.x, tPt.y) Then
            Me.PopupMenu frmMain.mnuEventItemContext, vbPopupMenuLeftAlign, , , frmMain.mnuEventItem(m_lCurrentEvent)
        ElseIf PtInRect(m_tTracks(2).tTrackArea, tPt.x, tPt.y) Then
            Me.PopupMenu frmMain.mnuEventEditContext, vbPopupMenuLeftAlign
        End If
    End If
    
End Sub

Private Sub DrawEventLine( _
            ByVal lX1 As Long, _
            ByVal lY1 As Long, _
            ByVal lX2 As Long, _
            ByVal lY2 As Long)
    Dim fValue1 As Single
    Dim fValue2 As Single
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim lHeight As Long
    Dim lTemp   As Long
    Dim fTemp   As Single
    Dim lIndex  As Long
    Dim lCount  As Long
    Dim fTheta  As Single
    
    If lX1 < m_tTracks(2).tTrackArea.Left Then
        lX1 = m_tTracks(2).tTrackArea.Left
    ElseIf lX1 > m_tTracks(2).tTrackArea.Right Then
        lX1 = m_tTracks(2).tTrackArea.Right
    End If
    
    If lY1 < m_tTracks(2).tTrackArea.Top Then
        lY1 = m_tTracks(2).tTrackArea.Top
    ElseIf lY1 > m_tTracks(2).tTrackArea.Bottom Then
        lY1 = m_tTracks(2).tTrackArea.Bottom
    End If
    
    If lX2 < m_tTracks(2).tTrackArea.Left Then
        lX2 = m_tTracks(2).tTrackArea.Left
    ElseIf lX2 > m_tTracks(2).tTrackArea.Right Then
        lX2 = m_tTracks(2).tTrackArea.Right
    End If
    
    If lY2 < m_tTracks(2).tTrackArea.Top Then
        lY2 = m_tTracks(2).tTrackArea.Top
    ElseIf lY2 > m_tTracks(2).tTrackArea.Bottom Then
        lY2 = m_tTracks(2).tTrackArea.Bottom
    End If
                
    lIndex1 = PixToBars(lX1) * 4 * EVENTS_QUANTIZATION
    lIndex2 = PixToBars(lX2) * 4 * EVENTS_QUANTIZATION
          
    lHeight = m_tTracks(2).tTrackArea.Bottom - m_tTracks(2).tTrackArea.Top
          
    fValue1 = (m_tTracks(2).tTrackArea.Bottom - lY1) / lHeight
    fValue2 = (m_tTracks(2).tTrackArea.Bottom - lY2) / lHeight
    
    If lIndex1 > lIndex2 Then
    
        lTemp = lIndex1
        lIndex1 = lIndex2
        lIndex2 = lTemp
        fTemp = fValue1
        fValue1 = fValue2
        fValue2 = fTemp
        
    End If
    
    With g_tSong.tEvents(m_lCurrentEvent)
    
        If Not .bInit Then
            ReDim .dEvents(NUM_OF_EVENT_TICKS - 1)
            .bInit = True
        End If
        
        lCount = lIndex2 - lIndex1 + 1
        
        For lIndex = 0 To lCount - 1
            
            fTheta = lIndex / lCount
            .dEvents(lIndex1 + lIndex) = fValue2 * fTheta + fValue1 * (1 - fTheta)
            
        Next

    End With
    
End Sub

Private Sub EraseEventLine( _
            ByVal lX1 As Long, _
            ByVal lX2 As Long)
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim lHeight As Long
    Dim lTemp   As Long
    Dim lIndex  As Long
    Dim lCount  As Long
    Dim fValue  As Single
    Dim bRev    As Boolean
    
    If lX1 < m_tTracks(2).tTrackArea.Left Then
        lX1 = m_tTracks(2).tTrackArea.Left
    ElseIf lX1 > m_tTracks(2).tTrackArea.Right Then
        lX1 = m_tTracks(2).tTrackArea.Right
    End If
    
    If lX2 < m_tTracks(2).tTrackArea.Left Then
        lX2 = m_tTracks(2).tTrackArea.Left
    ElseIf lX2 > m_tTracks(2).tTrackArea.Right Then
        lX2 = m_tTracks(2).tTrackArea.Right
    End If
    
    lIndex1 = PixToBars(lX1) * 4 * EVENTS_QUANTIZATION
    lIndex2 = PixToBars(lX2) * 4 * EVENTS_QUANTIZATION
          
    If lIndex1 > lIndex2 Then
    
        lTemp = lIndex1
        lIndex1 = lIndex2
        lIndex2 = lTemp
        bRev = True
        
    End If
    
    With g_tSong.tEvents(m_lCurrentEvent)
    
        If Not .bInit Then
            ReDim .dEvents(NUM_OF_EVENT_TICKS - 1)
            .bInit = True
        End If
        
        lCount = lIndex2 - lIndex1 + 1

        fValue = .dEvents(lIndex1)

        For lIndex = 1 To lCount - 1
            .dEvents(lIndex1 + lIndex) = fValue
        Next

    End With
    
End Sub

Private Sub Form_Resize()
    Dim lClientWidth    As Long
    Dim lClientHeight   As Long
    Dim lTrackHeight    As Long
    Dim lTimelineHeight As Long
    Dim lPanelWidth     As Long
    Dim lTrackIndex     As Long
    Dim lEventHeight    As Long
    Dim lIndex          As Long
    Dim tPos            As Point
    
    lClientWidth = Me.ScaleWidth
    lClientHeight = Me.ScaleHeight
    
    lTimelineHeight = Me.TextHeight("0") * 2
    
    If lClientHeight < lTimelineHeight * 5 Then
        lClientHeight = lTimelineHeight * 5
    End If
    
    lPanelWidth = 128
    
    If lClientWidth < lPanelWidth * 2 Then
        lClientWidth = lPanelWidth * 2
    End If
    
    lEventHeight = ((lClientHeight - m_tTimelineArea.Bottom) \ (UBound(m_tTracks) + 1)) * 2
    lTrackHeight = (lClientHeight - m_tTimelineArea.Bottom - lEventHeight) \ UBound(m_tTracks)
    
    SetRect m_tTimelineArea, lPanelWidth, 0, lClientWidth, lTimelineHeight
    
    tPos.x = lPanelWidth:   tPos.y = m_tTimelineArea.Bottom
    
    For lIndex = 0 To UBound(m_tTracks) - 1
        
        SetRect m_tTracks(lIndex).tTrackArea, tPos.x, tPos.y, m_tTimelineArea.Right, tPos.y + lTrackHeight
        SetRect m_tTracks(lIndex).tPanelArea, 0, tPos.y, lPanelWidth, tPos.y + lTrackHeight
        tPos.y = tPos.y + lTrackHeight
        
    Next
    
    SetRect m_tTracks(UBound(m_tTracks)).tTrackArea, tPos.x, tPos.y, m_tTimelineArea.Right, lClientHeight
    SetRect m_tTracks(UBound(m_tTracks)).tPanelArea, 0, tPos.y, lPanelWidth, lClientHeight
    
    m_fBarWidth = (m_tTimelineArea.Right - m_tTimelineArea.Left) / NUM_OF_BARS
    
    Redraw
    
End Sub

Private Sub Form_Unload( _
            ByRef Cancel As Integer)
    
    If m_hBufDC Then
        RestoreDC m_hBufDC, -1
        DeleteDC m_hBufDC
    End If
    
    If m_hBufBmp Then
        DeleteObject m_hBufBmp
    End If
    
    m_bLoaded = False
    
End Sub
