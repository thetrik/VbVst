VERSION 5.00
Begin VB.UserControl ctlPianoRoll 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   ClipControls    =   0   'False
   FillColor       =   &H00404040&
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
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
End
Attribute VB_Name = "ctlPianoRoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' //
' // ctlPianoRoll.ctl - monophonic piano-roll control
' // by The trick, 2022
' //

Option Explicit

Public Event PatternChanged()

Private Enum eMouseTrackMode
    MTM_NONE
    MTM_SCROLL
    MTM_CREATE_KEY
    MTM_MOVE_KEY
    MTM_LEFT_SIZE
    MTM_RIGHT_SIZE
    MTM_REMOVING
End Enum

Private Enum eSizeGripPos
    SGP_NONE
    SGP_LEFT
    SGP_RIGHT
End Enum

Private Type tRectArea
    lLeft   As Long
    lTop    As Long
    lWidth  As Long
    lHeight As Long
End Type

Private Type tScroll
    lMax        As Long
    lValue      As Long
    bEnabled    As Boolean
    tArea       As tRectArea
    tTrack      As tRectArea
End Type

Private Type tEditKey
    lValue      As Long
    dInitialPos As Single
    lInitialVal As Long
    dPos        As Single
    dLength     As Single
    lSelKey     As Long
End Type

Private Const SCROLL_WIDTH      As Long = 16
Private Const KEY_HEIGHT        As Long = 14
Private Const KEY_WIDTH         As Long = 64
Private Const KEY_MIN_LENGTH    As Single = 0.1
Private Const SIZE_GRIP_AREA    As Long = 3
Private Const TIME_BAR_SIZE     As Long = 10

Private m_hMemDC        As Handle
Private m_hMemBmp       As Handle
Private m_tClientArea   As tRectArea
Private m_tScroll       As tScroll
Private m_eMouseMode    As eMouseTrackMode
Private m_tMouseOffest  As POINT
Private m_lBeats        As Long
Private m_lDivisions    As Long
Private m_bLocked       As Boolean
Private m_tPattern      As tPattern
Private m_bSnap         As Boolean
Private m_tActiveKey    As tEditKey
Private m_fPlayPos      As Single

Friend Property Get Pattern() As tPattern
    Pattern = m_tPattern
End Property

Friend Property Let Pattern( _
                    ByRef tPattern As tPattern)
                    
    m_tPattern = tPattern
    m_lBeats = tPattern.lLengthPerBeats
    Refresh
    
End Property

Public Property Let PlaybackPos( _
                    ByVal fValue As Single)
    Dim lCurSlot    As Long
    
    lCurSlot = m_lDivisions * m_fPlayPos / (m_lBeats * 4)
    
    m_fPlayPos = fValue
    
    If lCurSlot <> m_lDivisions * m_fPlayPos / (m_lBeats * 4) Then
        RedrawPlayback
    End If

End Property
Public Property Get PlaybackPos() As Single
    PlaybackPos = m_fPlayPos
End Property

Public Property Let Snap( _
                    ByVal bValue As Boolean)
    m_bSnap = bValue
End Property
Public Property Get Snap() As Boolean
    Snap = m_bSnap
End Property

Public Property Let Locked( _
                    ByVal bValue As Boolean)
    m_bLocked = bValue
End Property
Public Property Get Locked() As Boolean
    Locked = m_bLocked
End Property

Public Property Let Beats( _
                    ByVal lValue As Long)
                    
    If m_lBeats = lValue Then
        Exit Property
    ElseIf lValue < 1 Then
        lValue = 1
    ElseIf lValue > 8 Then
        lValue = 8
    End If
    
    m_lBeats = lValue
    RaiseEvent PatternChanged
    Refresh
    
End Property
Public Property Get Beats() As Long
    Beats = m_lBeats
End Property

Public Property Let Divisions( _
                    ByVal lValue As Long)
                    
    If m_lDivisions = lValue Then
        Exit Property
    ElseIf lValue < 1 Then
        lValue = 1
    ElseIf lValue > 8 Then
        lValue = 8
    End If
    
    m_lDivisions = lValue
    Refresh
    
End Property
Public Property Get Divisions() As Long
    Divisions = m_lDivisions
End Property

Public Sub Refresh()
    
    SaveDC UserControl.hdc

    DrawKeysGrid 0, 119
    DrawKeys
    RedrawPlayback
    RedrawScroll
    
    SelectObject UserControl.hdc, GetStockObject(NULL_BRUSH)
    SelectObject UserControl.hdc, GetStockObject(DC_PEN)
    
    SetDCPenColor UserControl.hdc, &HE0E0E0
    
    Rectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    RestoreDC UserControl.hdc, -1
    
End Sub

Private Sub RedrawPlayback()
    Dim hdc         As Handle
    Dim lGridCount  As Long
    Dim fGridWidth  As Single
    Dim lGridIndex  As Single
    Dim lPosX       As Long
    Dim lPlayIndex  As Long
    Dim tRC         As RECT
    
    hdc = UserControl.hdc
    
    SaveDC hdc
    
    SelectObject hdc, GetStockObject(DC_BRUSH)
    SelectObject hdc, GetStockObject(DC_PEN)
    
    SetDCBrushColor hdc, &H404040
    SetDCPenColor hdc, &HE0E0E0
    
    SetRect tRC, 0, m_tClientArea.lTop + m_tClientArea.lHeight, m_tClientArea.lLeft + m_tClientArea.lWidth + 1, UserControl.ScaleHeight
    Rectangle hdc, tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
    
    lGridCount = m_lBeats * m_lDivisions
    fGridWidth = m_tClientArea.lWidth / lGridCount
    lPlayIndex = lGridCount * (m_fPlayPos / (4 * m_lBeats))
    
    SetDCBrushColor hdc, &HF0F0F0
    
    For lGridIndex = 0 To lGridCount - 1
        
        lPosX = lGridIndex * fGridWidth + m_tClientArea.lLeft
        MoveToEx hdc, lPosX, m_tClientArea.lTop + m_tClientArea.lHeight, ByVal 0&
        LineTo hdc, lPosX, UserControl.ScaleHeight
        
        If lGridIndex = lPlayIndex Then
            PatBlt hdc, lPosX + 2, m_tClientArea.lTop + m_tClientArea.lHeight + 2, CLng(fGridWidth - 3), TIME_BAR_SIZE - 4, vbPatCopy
        End If
        
    Next
    
    RestoreDC hdc, -1
    
    InvalidateRect UserControl.hWnd, tRC, 0
    
End Sub

Private Sub DrawKeys( _
            Optional ByVal lStartRow As Long = 0, _
            Optional ByVal lEndRow As Long = 119)
    Dim hdc     As Handle
    Dim lIndex  As Long
    Dim tArea   As tRectArea
    Dim bDraw   As Boolean
    
    hdc = UserControl.hdc
    
    SaveDC hdc
    
    SelectObject hdc, GetStockObject(DC_PEN)
    SelectObject hdc, GetStockObject(DC_BRUSH)
    
    SetDCBrushColor hdc, &HF09090
    SetDCPenColor hdc, &H303030
    
    For lIndex = 0 To m_tPattern.lNumOfKeys - 1
        
        With m_tPattern.tKeys(lIndex)
            
            If .lValue >= lStartRow And .lValue <= lEndRow Then
                
                tArea.lTop = NoteToPosY(.lValue) + 1
                tArea.lHeight = KEY_HEIGHT - 1
                
                bDraw = False
                
                If tArea.lTop < m_tClientArea.lTop Then
                    If tArea.lHeight + tArea.lTop > m_tClientArea.lTop Then
                        bDraw = True
                    End If
                ElseIf tArea.lTop < m_tClientArea.lTop + m_tClientArea.lHeight Then
                    bDraw = True
                End If
                    
                If bDraw Then
                
                    tArea.lLeft = QuarterPosToPix(.dPos)
                    tArea.lWidth = QuarterLengthToPix(.dLength)
                    
                    If tArea.lWidth < 3 Then
                        tArea.lWidth = 3
                    End If
                    
                    Rectangle hdc, tArea.lLeft, tArea.lTop, tArea.lLeft + tArea.lWidth, tArea.lTop + tArea.lHeight

                End If
                
            End If
            
        End With
        
    Next
    
    RestoreDC hdc, -1
    
End Sub

Private Sub DrawKeysGrid( _
            ByVal lStartKey As Long, _
            ByVal lEndKey As Long)
    Dim lTopVisible     As Long
    Dim lFirstVisible   As Long
    Dim lTotalRows      As Long
    Dim lPosY           As Long
    Dim lPosX           As Long
    Dim lHeight         As Long
    Dim lBackColor      As Long
    Dim sOctave         As String
    Dim lGridIndex      As Long
    Dim lGridCount      As Long
    Dim fGridWidth      As Single
    Dim hdc             As Handle
    Dim tRC             As RECT
    
    If m_tScroll.bEnabled Then
        lTopVisible = 119 - (m_tScroll.lValue \ KEY_HEIGHT)
        lFirstVisible = lTopVisible - (m_tClientArea.lHeight + (m_tScroll.lValue Mod KEY_HEIGHT)) \ KEY_HEIGHT
    Else
        lTopVisible = 119
        lFirstVisible = 0
    End If
    
    If lStartKey < lFirstVisible Then
        lStartKey = lFirstVisible
    End If
    
    If lEndKey > lTopVisible Then
        lEndKey = lTopVisible
    End If

    lTotalRows = lEndKey - lStartKey + 1

    If lTotalRows <= 0 Then
        Exit Sub
    End If
         
    lPosY = (119 - lStartKey) * KEY_HEIGHT - m_tScroll.lValue + m_tClientArea.lTop
    lHeight = lTotalRows * KEY_HEIGHT
    SetRect tRC, 0, lPosY - lHeight + KEY_HEIGHT, m_tClientArea.lLeft + m_tClientArea.lWidth, lPosY + KEY_HEIGHT
    
    hdc = UserControl.hdc
    
    SaveDC hdc
    SelectObject hdc, GetStockObject(DC_PEN)
    SelectObject hdc, GetStockObject(DC_BRUSH)
    
    SetDCBrushColor hdc, &HF0F0F0
    PatBlt hdc, tRC.Left, tRC.Top, m_tClientArea.lLeft, tRC.Bottom - tRC.Top, vbPatCopy

    Do Until lStartKey > lEndKey
        
        lBackColor = UserControl.BackColor
        
        Select Case lStartKey Mod 12
        Case 4, 11
        
            SetDCPenColor hdc, 0
            MoveToEx hdc, 0, lPosY, ByVal 0&
            LineTo hdc, m_tClientArea.lLeft, lPosY
            
        Case 1, 3, 6, 8, 10
        
            SetDCBrushColor hdc, &H404040
            SetDCPenColor hdc, 0
            MoveToEx hdc, 0, lPosY + KEY_HEIGHT \ 2, ByVal 0&
            LineTo hdc, m_tClientArea.lLeft, lPosY + KEY_HEIGHT \ 2
            PatBlt hdc, 0, lPosY, (m_tClientArea.lLeft \ 3) * 2, KEY_HEIGHT, vbPatCopy
            lBackColor = &H303030
            
        End Select
        
        If (lStartKey Mod 12) = 0 Then
            
            sOctave = "C" & CStr(lStartKey \ 12)
            SetTextColor hdc, &H404040
            TextOut hdc, 0, lPosY, sOctave, Len(sOctave)
        
        End If
        
        SetDCBrushColor hdc, lBackColor
        PatBlt hdc, m_tClientArea.lLeft, lPosY, m_tClientArea.lWidth, KEY_HEIGHT, vbPatCopy
        MoveToEx hdc, m_tClientArea.lLeft, lPosY, ByVal 0&
        SetDCPenColor hdc, &H707070
        LineTo hdc, m_tClientArea.lLeft + m_tClientArea.lWidth, lPosY
        
        lStartKey = lStartKey + 1
        lPosY = lPosY - KEY_HEIGHT
        
    Loop
    
    lGridCount = m_lBeats * m_lDivisions
    
    fGridWidth = m_tClientArea.lWidth / lGridCount
    
    For lGridIndex = 0 To lGridCount - 1
        
        lPosX = lGridIndex * fGridWidth + m_tClientArea.lLeft
        
        If (lGridIndex Mod m_lDivisions) = 0 Then
            lBackColor = &H707070
        Else
            lBackColor = &H505050
        End If
        
        SetDCPenColor hdc, lBackColor
        MoveToEx hdc, lPosX, m_tClientArea.lTop, ByVal 0&
        LineTo hdc, lPosX, m_tClientArea.lTop + m_tClientArea.lHeight
        
    Next
    
    InvalidateRect UserControl.hWnd, tRC, 0
    
    RestoreDC hdc, -1
    
End Sub

Private Sub InvalidateRow( _
            ByVal lIndex As Long)
    Dim tRC     As RECT
    Dim lPosY   As Long
    
    lPosY = NoteToPosY(lIndex)
    
    SetRect tRC, m_tClientArea.lLeft, lPosY, m_tClientArea.lLeft + m_tClientArea.lWidth, lPosY + KEY_HEIGHT
    
    InvalidateRect UserControl.hWnd, tRC, 0
    
End Sub

Private Sub RedrawActiveKey()
    Dim tArea   As tRectArea
    Dim hdc     As Handle
    Dim tRC     As RECT
    
    tArea.lLeft = QuarterPosToPix(m_tActiveKey.dPos)
    tArea.lTop = NoteToPosY(m_tActiveKey.lValue)
    tArea.lHeight = KEY_HEIGHT
    tArea.lWidth = QuarterLengthToPix(m_tActiveKey.dLength)
    
    If tArea.lWidth = 0 Then
        tArea.lWidth = 2
    End If
    
    If tArea.lLeft + tArea.lWidth > m_tClientArea.lLeft + m_tClientArea.lWidth Then
        tArea.lWidth = m_tClientArea.lLeft + m_tClientArea.lWidth - tArea.lLeft
    End If
    
    SetRect tRC, m_tClientArea.lLeft, tArea.lTop, m_tClientArea.lLeft + m_tClientArea.lWidth, tArea.lTop + tArea.lHeight
    
    RedrawWindow UserControl.hWnd, tRC, 0, RDW_INVALIDATE Or RDW_UPDATENOW
    
    hdc = GetDC(UserControl.hWnd)
    
    If hdc Then
        AlphaBlend hdc, tArea.lLeft, tArea.lTop, tArea.lWidth, tArea.lHeight, m_hMemDC, 0, 0, 1, 1, &H800000
        ReleaseDC UserControl.hWnd, hdc
    End If
    
End Sub

Private Sub DrawKey( _
            ByVal lIndex As Long, _
            Optional ByVal bActive As Boolean)
    Dim hdc     As Handle
    Dim tArea   As tRectArea
    Dim bDraw   As Boolean
    
    hdc = UserControl.hdc
    
    SaveDC hdc
    
    SelectObject hdc, GetStockObject(DC_PEN)
    SelectObject hdc, GetStockObject(DC_BRUSH)
    
    If bActive Then
        SetDCBrushColor hdc, &HC06080
        SetDCPenColor hdc, &H303030
    Else
        SetDCBrushColor hdc, &HF09090
        SetDCPenColor hdc, &H303030
    End If
    
    With m_tPattern.tKeys(lIndex)

        tArea.lTop = NoteToPosY(.lValue) + 1
        tArea.lHeight = KEY_HEIGHT - 1
        
        If tArea.lTop < m_tClientArea.lTop Then
            If tArea.lHeight + tArea.lTop > m_tClientArea.lTop Then
                bDraw = True
            End If
        ElseIf tArea.lTop < m_tClientArea.lTop + m_tClientArea.lHeight Then
            bDraw = True
        End If
            
        If bDraw Then
        
            tArea.lLeft = QuarterPosToPix(.dPos)
            tArea.lWidth = QuarterLengthToPix(.dLength)
            
            If tArea.lWidth < 3 Then
                tArea.lWidth = 3
            End If
                    
            Rectangle hdc, tArea.lLeft, tArea.lTop, tArea.lLeft + tArea.lWidth, tArea.lTop + tArea.lHeight
            
        End If

    End With
    
    RestoreDC hdc, -1
    
End Sub

Private Sub RedrawScroll()
    Dim hdc As Handle
    Dim tRC As RECT
    
    hdc = UserControl.hdc
    
    SaveDC hdc

    With m_tScroll
    
        SetRect tRC, .tArea.lLeft, .tArea.lTop, .tArea.lLeft + .tArea.lWidth, .tArea.lTop + .tArea.lHeight
    
        SelectObject hdc, GetStockObject(DC_PEN)
        SelectObject hdc, GetStockObject(DC_BRUSH)
        
        SetDCBrushColor hdc, &H404040
        SetDCPenColor hdc, &HE0E0E0
        PatBlt hdc, .tArea.lLeft, .tArea.lTop, .tArea.lWidth, .tArea.lHeight, vbPatCopy
        Rectangle hdc, tRC.Left, tRC.Top, tRC.Right, tRC.Bottom

        If .bEnabled Then
            SetDCBrushColor hdc, &HE0E0E0
            PatBlt hdc, .tTrack.lLeft, .tTrack.lTop, .tTrack.lWidth, .tTrack.lHeight, vbPatCopy
        End If
    
    End With
    
    InvalidateRect UserControl.hWnd, tRC, 0
    
    RestoreDC hdc, -1
    
End Sub

Private Sub UserControl_Initialize()
    Dim tBI     As BITMAPINFO
    Dim pBits   As Long
    
    m_lBeats = 4
    m_lDivisions = 4
    m_bSnap = True

    m_hMemDC = CreateCompatibleDC(UserControl.hdc)
    
    With tBI.bmiHeader
        .biSize = Len(tBI.bmiHeader)
        .biBitCount = 32
        .biHeight = 1
        .biWidth = 1
        .biPlanes = 1
    End With
    
    m_hMemBmp = CreateDIBSection(UserControl.hdc, tBI, 0, pBits, 0, 0)
    If m_hMemBmp Then
        PutMem4 ByVal pBits, &HFFA080
    End If
    
    SaveDC m_hMemDC
    SelectObject m_hMemDC, m_hMemBmp
    
End Sub

Private Sub UserControl_Show()
    ScrollClientArea m_tScroll.lMax \ 2
End Sub

Private Sub UserControl_Terminate()
    
    RestoreDC m_hMemDC, -1
    DeleteObject m_hMemBmp
    DeleteDC m_hMemDC
    
End Sub

Private Sub OnScroll( _
            ByVal lDelta As Long)
    Dim tRCScroll   As RECT
    Dim tRCUpdate   As RECT
    Dim lStartRow   As Long
    Dim lEndRow     As Long
    
    SaveDC UserControl.hdc
    
    SetRect tRCScroll, 1, m_tClientArea.lTop + 1, m_tClientArea.lLeft + m_tClientArea.lWidth, _
            m_tClientArea.lTop + m_tClientArea.lHeight - 1
    
    ScrollDC UserControl.hdc, 0, -lDelta, tRCScroll, tRCScroll, 0, tRCUpdate
    
    lEndRow = PosYToNote(tRCUpdate.Top)
    lStartRow = PosYToNote(tRCUpdate.Bottom)
    
    IntersectClipRect UserControl.hdc, tRCUpdate.Left, tRCUpdate.Top, tRCUpdate.Right, tRCUpdate.Bottom

    DrawKeysGrid lStartRow, lEndRow
    DrawKeys lStartRow, lEndRow
    
    RestoreDC hdc, -1
    
    InvalidateRect UserControl.hWnd, ByVal 0&, 0
    
End Sub

Private Sub ScrollClientArea( _
            ByVal lDelta As Long)
    Dim lNewValue   As Long
    
    With m_tScroll
        
        lNewValue = .lValue + lDelta
        
        If lNewValue > .lMax Then
            lNewValue = .lMax
        ElseIf lNewValue < 0 Then
            lNewValue = 0
        End If
        
        If lNewValue = .lValue Then
            Exit Sub
        Else
            lDelta = lNewValue - .lValue
        End If
        
        .lValue = lNewValue
        .tTrack.lTop = (.tArea.lHeight - .tTrack.lHeight - 4) * (.lValue / .lMax) + 2

        RedrawScroll
        OnScroll lDelta
        
    End With
    
End Sub

Private Sub UserControl_MouseDown( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    Dim tPt     As POINT
    Dim lValue  As Long
    Dim lKey    As Long
    
    tPt.x = fX
    tPt.y = fY
    
    If iButton = vbLeftButton Then
        
        If m_tScroll.bEnabled Then
            
            With m_tScroll
                
                If PtIsInArea(tPt.x, tPt.y, .tArea) Then
                    
                    m_eMouseMode = MTM_SCROLL
                    
                    If PtIsInArea(tPt.x, tPt.y, .tTrack) Then
                        ' // Pointer on track
                        m_tMouseOffest.x = tPt.x - .tTrack.lLeft
                        m_tMouseOffest.y = tPt.y - .tTrack.lTop
                    Else
                    
                        ' // Move track
                        lValue = (tPt.y - .tTrack.lHeight \ 2) / (.tArea.lHeight - .tTrack.lHeight) * .lMax
                        
                        If lValue < 0 Then
                            lValue = 0
                        ElseIf lValue > .lMax Then
                            lValue = .lMax
                        End If
                        
                        If lValue <> .lValue Then
                        
                            .lValue = lValue
                            .tTrack.lTop = (.tArea.lHeight - 4 - .tTrack.lHeight) * (.lValue / .lMax) + 2
                            Refresh
                            
                        End If
                        
                        m_tMouseOffest.x = tPt.x - .tTrack.lLeft
                        m_tMouseOffest.y = tPt.y - .tTrack.lTop
                        
                    End If
                    
                    Exit Sub
                    
                End If
                
            End With
            
        End If
        
        If Not m_bLocked Then
            
            ' // Cursor within client area.
            If PtIsInArea(tPt.x, tPt.y, m_tClientArea) Then
                
                ' // Check
                lKey = GetKeyFromPos(tPt.x, tPt.y)
                
                If lKey = -1 Then
                    
                    ' // Create new
                    m_eMouseMode = MTM_CREATE_KEY
                    
                    m_tActiveKey.dLength = 1
                    m_tActiveKey.dPos = PosXToQuarterPos(tPt.x)
                    m_tActiveKey.lValue = PosYToNote(tPt.y)
                    
                    SnapActiveKey SGP_LEFT
                    
                    m_tActiveKey.dInitialPos = m_tActiveKey.dPos
                    
                    If m_bSnap Then
                        m_tMouseOffest.x = QuarterPosToPix(m_tActiveKey.dPos)
                    Else
                        m_tMouseOffest.x = tPt.x
                    End If
                    
                    m_tMouseOffest.y = tPt.y
                    
                    UserControl.MousePointer = vbSizeWE
                    
                Else
                    
                    m_tActiveKey.lSelKey = lKey
                    m_tActiveKey.dPos = m_tPattern.tKeys(lKey).dPos
                    m_tActiveKey.lValue = m_tPattern.tKeys(lKey).lValue
                    m_tActiveKey.dLength = m_tPattern.tKeys(lKey).dLength
                    m_tMouseOffest.y = tPt.y
                    
                    Select Case CheckPosOnSizeGrip(lKey, tPt.x)
                    Case SGP_LEFT
                    
                        m_eMouseMode = MTM_LEFT_SIZE
                        m_tActiveKey.dInitialPos = m_tActiveKey.dPos + m_tActiveKey.dLength
                        m_tMouseOffest.x = QuarterPosToPix(m_tActiveKey.dPos + m_tActiveKey.dLength)
                        UserControl.MousePointer = vbSizeWE
                        
                    Case SGP_RIGHT
                    
                        m_eMouseMode = MTM_RIGHT_SIZE
                        m_tActiveKey.dInitialPos = m_tActiveKey.dPos
                        m_tMouseOffest.x = QuarterPosToPix(m_tActiveKey.dPos)
                        UserControl.MousePointer = vbSizeWE
                        
                    Case SGP_NONE
                    
                        m_eMouseMode = MTM_MOVE_KEY
                        m_tActiveKey.dInitialPos = m_tActiveKey.dPos
                        m_tActiveKey.lInitialVal = m_tActiveKey.lValue
                        m_tMouseOffest.x = tPt.x
                        UserControl.MousePointer = vbSizeAll
                        DrawKey m_tActiveKey.lSelKey, True
                        
                    End Select
                    
                End If
                
                RedrawActiveKey
                
            End If
        End If
        
    ElseIf iButton = vbRightButton Then
        
        If m_bLocked Then
            Exit Sub
        End If
        
        If m_eMouseMode <> MTM_NONE Then
            
            If m_eMouseMode = MTM_MOVE_KEY Then
                DrawKey m_tActiveKey.lSelKey, False
                InvalidateRow m_tActiveKey.lInitialVal
            End If
            
            ' // Cancel operation
            m_eMouseMode = MTM_NONE
            InvalidateRow m_tActiveKey.lValue
            UserControl.MousePointer = vbDefault
            
        Else
            ' // Remove key
            
            ' // Cursor within client area.
            If PtIsInArea(tPt.x, tPt.y, m_tClientArea) Then
                
                ' // Check
                lKey = GetKeyFromPos(tPt.x, tPt.y)
                
                If lKey = -1 Then
                    Exit Sub
                End If
                
                m_tActiveKey.lSelKey = lKey
                DrawKey lKey, True
                InvalidateRow m_tPattern.tKeys(lKey).lValue
                UserControl.MousePointer = vbNoDrop
                m_eMouseMode = MTM_REMOVING
                
            End If

        End If
    
    End If
      
End Sub

Private Sub UserControl_MouseMove( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    Dim tPt     As POINT
    Dim tArea   As tRectArea
    Dim lValue  As Long
    Dim lDelta  As Long
    Dim lIndex  As Long
    Dim lPosY   As Long
    
    tPt.x = fX
    tPt.y = fY
    
    If iButton = vbLeftButton Then
    
        If m_eMouseMode = MTM_SCROLL And m_tScroll.bEnabled Then
            
            With m_tScroll
                 
                tPt.x = tPt.x - m_tMouseOffest.x
                tPt.y = tPt.y - m_tMouseOffest.y
                
                lValue = (tPt.y) / (.tArea.lHeight - .tTrack.lHeight) * .lMax
                
                If lValue < 0 Then
                    lValue = 0
                ElseIf lValue > .lMax Then
                    lValue = .lMax
                End If
                
                If lValue <> .lValue Then
                    
                    lDelta = lValue - .lValue
                    .lValue = lValue
                    .tTrack.lTop = (.tArea.lHeight - 4 - .tTrack.lHeight) * (.lValue / .lMax) + 2
                    
                    RedrawScroll
                    OnScroll lDelta
                    
                End If
                        
            End With
        
        ElseIf m_eMouseMode = MTM_CREATE_KEY Or m_eMouseMode = MTM_LEFT_SIZE Or m_eMouseMode = MTM_RIGHT_SIZE Then
            
            lDelta = tPt.x - m_tMouseOffest.x
            
            If lDelta = 0 Then
                Exit Sub
            End If

            With m_tActiveKey
                
                If lDelta > 0 Then
                    
                    .dPos = .dInitialPos
                    .dLength = PosXToQuarterPos(tPt.x) - .dPos
                    SnapActiveKey SGP_RIGHT
                    
                ElseIf lDelta < 0 Then
                
                    .dPos = PosXToQuarterPos(tPt.x)
                    SnapActiveKey SGP_LEFT
                    .dLength = .dInitialPos - .dPos

                End If

            End With
            
            RedrawActiveKey
            
        ElseIf m_eMouseMode = MTM_MOVE_KEY Then
        
            lDelta = tPt.x - m_tMouseOffest.x
            
            If lDelta = 0 Then
                Exit Sub
            End If
            
            With m_tActiveKey
            
                .dPos = PixLenToQuarterLen(tPt.x - m_tMouseOffest.x) + .dInitialPos
                
                If .dPos < 0 Then
                    .dPos = 0
                End If
                
                SnapActiveKey SGP_LEFT

                lDelta = (tPt.y - m_tMouseOffest.y) \ KEY_HEIGHT
                
                If lDelta Then
                
                    InvalidateRow .lValue
                    .lValue = .lInitialVal - lDelta
                    
                    If .lValue > 119 Then
                        .lValue = 119
                    ElseIf .lValue < 0 Then
                        .lValue = 0
                    End If

                    lPosY = NoteToPosY(.lValue)

                    If lPosY < m_tClientArea.lTop Then
                    
                        lDelta = m_tClientArea.lTop + lPosY
                        m_tMouseOffest.y = m_tMouseOffest.y - lDelta
                        ScrollClientArea lDelta
                        
                    ElseIf lPosY + KEY_HEIGHT > m_tClientArea.lTop + m_tClientArea.lHeight Then
                        
                        lDelta = lPosY + KEY_HEIGHT - (m_tClientArea.lTop + m_tClientArea.lHeight)
                        m_tMouseOffest.y = m_tMouseOffest.y - lDelta
                        ScrollClientArea lDelta
                        
                    End If
                    
                End If
                
            End With
            
            RedrawActiveKey
            
        End If
    
    ElseIf iButton = vbRightButton Then
    
        If m_eMouseMode = MTM_REMOVING Then
            If m_tActiveKey.lSelKey = GetKeyFromPos(tPt.x, tPt.y) Then
                UserControl.MousePointer = vbNoDrop
            Else
                UserControl.MousePointer = vbDefault
            End If
        End If
    
    ElseIf iButton = 0 Then
        
        If m_bLocked Then
            Exit Sub
        End If
        
        ' // Track size
        For lIndex = 0 To m_tPattern.lNumOfKeys - 1
            
            tArea.lTop = NoteToPosY(m_tPattern.tKeys(lIndex).lValue)
            tArea.lHeight = KEY_HEIGHT
            
            If tPt.y >= tArea.lTop And tPt.y < tArea.lTop + tArea.lHeight Then
                
                tArea.lLeft = QuarterPosToPix(m_tPattern.tKeys(lIndex).dPos)
                tArea.lWidth = QuarterLengthToPix(m_tPattern.tKeys(lIndex).dLength)
                
                If tPt.x >= tArea.lLeft And tPt.x < tArea.lLeft + tArea.lWidth Then
                                        
                    ' // Over key-note
                    If tPt.x - tArea.lLeft < SIZE_GRIP_AREA Or _
                        tArea.lLeft + tArea.lWidth - tPt.x < SIZE_GRIP_AREA Then
                        UserControl.MousePointer = vbSizeWE
                    Else
                        UserControl.MousePointer = vbSizeAll
                    End If
                                  
                    Exit For
                    
                End If
                
            End If
            
        Next
        
        If lIndex = m_tPattern.lNumOfKeys Then
            UserControl.MousePointer = vbDefault
        End If
        
    End If
    
End Sub

Private Sub UserControl_MouseUp( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    Dim tPt As POINT

    tPt.x = fX
    tPt.y = fY
    
    If iButton = vbLeftButton Then
        
        Select Case m_eMouseMode
        Case MTM_SCROLL
            m_eMouseMode = MTM_NONE
            Exit Sub
        Case MTM_CREATE_KEY, MTM_LEFT_SIZE, MTM_MOVE_KEY, MTM_RIGHT_SIZE

            If m_eMouseMode <> MTM_CREATE_KEY Then
                RemoveKeyFromPattern m_tActiveKey.lSelKey
            End If
            
            If m_tActiveKey.dLength > KEY_MIN_LENGTH Then
                PutActiveKeyToPattern
                Refresh
            Else
                If m_eMouseMode <> MTM_CREATE_KEY Then
                    Refresh
                Else
                    InvalidateRow m_tActiveKey.lValue
                End If
            End If
            
            m_eMouseMode = MTM_NONE
            UserControl.MousePointer = vbDefault
            
        End Select
        
    ElseIf iButton = vbRightButton Then
    
        If m_eMouseMode = MTM_REMOVING Then
        
            If m_tActiveKey.lSelKey = GetKeyFromPos(tPt.x, tPt.y) Then
                RemoveKeyFromPattern m_tActiveKey.lSelKey
            Else
                DrawKey m_tActiveKey.lSelKey, False
            End If
            
            Refresh
            m_eMouseMode = MTM_NONE
            UserControl.MousePointer = vbDefault
            
        End If
        
    End If
    
End Sub

Private Sub RemoveKeyFromPattern( _
            ByVal lIndex As Long)
        
    If lIndex < m_tPattern.lNumOfKeys - 1 Then
        memcpy m_tPattern.tKeys(lIndex), m_tPattern.tKeys(lIndex + 1), _
                (m_tPattern.lNumOfKeys - lIndex - 1) * LenB(m_tPattern.tKeys(0))
    End If
    
    m_tPattern.lNumOfKeys = m_tPattern.lNumOfKeys - 1
    
    RaiseEvent PatternChanged
    
End Sub

Private Sub PutActiveKeyToPattern()
    Dim lIndex      As Long
    Dim bAddNewKey  As Boolean
    Dim dNewKeyPos  As Double
    Dim dNewKeyLen  As Double
    Dim lNewKeyVal  As Long
    Dim lStartDel   As Long
    Dim dActEnd     As Double
    Dim lDelCount   As Long
    Dim lPutIndex   As Long
    Dim lInsCount   As Long
    
    ' // Validate
    If m_tActiveKey.dPos < 0 Then
        m_tActiveKey.dLength = m_tActiveKey.dLength + m_tActiveKey.dPos
        m_tActiveKey.dPos = 0
    End If
    
    If m_tActiveKey.dPos + m_tActiveKey.dLength > m_lBeats * 4 Then
        m_tActiveKey.dLength = m_lBeats * 4 - m_tActiveKey.dPos
    End If
    
    If m_tActiveKey.dLength < KEY_MIN_LENGTH Then
        Exit Sub
    End If
    
    If m_tActiveKey.lValue > 119 Then
        m_tActiveKey.lValue = 119
    ElseIf m_tActiveKey.lValue < 0 Then
        m_tActiveKey.lValue = 0
    End If
    
    With m_tPattern
            
        dActEnd = m_tActiveKey.dPos + m_tActiveKey.dLength
        lStartDel = .lNumOfKeys
        lPutIndex = .lNumOfKeys
        
        For lIndex = 0 To .lNumOfKeys - 1
            With .tKeys(lIndex)
                If .dPos < m_tActiveKey.dPos And .dPos + .dLength > m_tActiveKey.dPos Then

                    ' // ....++++....
                    ' // ..oooo......

                    ' // Check if new key need
                    If .dPos + .dLength > dActEnd Then
                        
                        ' // ....++++....
                        ' // ..oooooooo..
                        
                        dNewKeyLen = .dPos + .dLength - dActEnd
                        
                        If dNewKeyLen > KEY_MIN_LENGTH Then
                        
                            bAddNewKey = True
                            dNewKeyPos = dActEnd
                            lNewKeyVal = .lValue
                            
                        End If
                        
                    End If
                
                    ' // Right trim
                    .dLength = m_tActiveKey.dPos - .dPos
                    
                    If .dLength < KEY_MIN_LENGTH Then
                        lStartDel = lIndex: lDelCount = 1:  lPutIndex = lIndex
                    Else
                        lPutIndex = lIndex + 1
                    End If

                ElseIf .dPos >= m_tActiveKey.dPos And .dPos + .dLength <= dActEnd Then
                        
                    ' // ..++++++++..
                    ' // ....oooo....
                    
                    ' // Delete
                    If lStartDel = m_tPattern.lNumOfKeys Then
                    
                        lStartDel = lIndex: lDelCount = 1
                        
                        If lPutIndex = m_tPattern.lNumOfKeys Then
                            lPutIndex = lIndex
                        End If
                        
                    Else
                        lDelCount = lDelCount + 1
                    End If

                ElseIf .dPos >= m_tActiveKey.dPos And .dPos < dActEnd Then
                    
                    ' // ..+++++....
                    ' // ....ooooo..
                    
                    ' // left trim
                    .dLength = .dLength - (dActEnd - .dPos)
                    .dPos = dActEnd
                    
                    If .dLength < KEY_MIN_LENGTH Then
                        If lStartDel = m_tPattern.lNumOfKeys Then
                            lStartDel = lIndex: lDelCount = 1
                        Else
                            lDelCount = lDelCount + 1
                        End If
                    End If
                    
                    If lPutIndex = m_tPattern.lNumOfKeys Then
                        lPutIndex = lIndex
                    End If
                    
                    Exit For
                
                ElseIf .dPos > m_tActiveKey.dPos Then
                
                    ' // ..oooo......
                    ' // .......+++..
                
                    If lPutIndex = m_tPattern.lNumOfKeys Then
                        lPutIndex = lIndex
                    End If
                    
                    Exit For
                    
                End If
            End With
        Next
        
        If bAddNewKey Then
            lInsCount = 2
        Else
            lInsCount = 1
        End If
        
        ' // Insert new
        If .lNumOfKeys Then
            If (lPutIndex + lInsCount) > UBound(.tKeys) Then
                ReDim Preserve .tKeys((lPutIndex + lInsCount) * 2 - 1)
            End If
        Else
            ReDim .tKeys(15)
        End If
        
        If lDelCount > lInsCount Then
            memcpy .tKeys(lStartDel + lInsCount), .tKeys(lStartDel + lDelCount), (.lNumOfKeys - (lStartDel + lDelCount)) * LenB(.tKeys(0))
        ElseIf lInsCount > lDelCount Then
            memcpy .tKeys(lPutIndex + lInsCount), .tKeys(lPutIndex), (.lNumOfKeys - lPutIndex) * LenB(.tKeys(0))
        End If
        
        With .tKeys(lPutIndex)
            
            .dPos = m_tActiveKey.dPos
            .dLength = m_tActiveKey.dLength
            .lValue = m_tActiveKey.lValue
            
        End With

        If bAddNewKey Then
            With .tKeys(lPutIndex + 1)
            
                .lValue = lNewKeyVal
                .dPos = dNewKeyPos
                .dLength = dNewKeyLen
                
            End With
        End If
        
        .lNumOfKeys = .lNumOfKeys - lDelCount + lInsCount
        
    End With
    
    RaiseEvent PatternChanged
    
End Sub

Private Sub SnapActiveKey( _
            ByVal eMode As eSizeGripPos)
    Dim lCells  As Long
    Dim lCell   As Long
    Dim dStart  As Double
    Dim dEnd    As Double
    
    If m_bSnap Then
        
        lCells = m_lBeats * m_lDivisions
        dStart = m_tActiveKey.dPos
        dEnd = dStart + m_tActiveKey.dLength
        
        If eMode = SGP_LEFT Then
        
            lCell = Int((dStart / (m_lBeats * 4)) * lCells)
            m_tActiveKey.dPos = lCell / lCells * (m_lBeats * 4)
            m_tActiveKey.dLength = dEnd - dStart
            
        ElseIf eMode = SGP_RIGHT Then
            
            lCell = Int((dEnd / (m_lBeats * 4)) * lCells)
            m_tActiveKey.dLength = lCell / lCells * (m_lBeats * 4) - dStart
        
        End If
        
    End If
    
End Sub

Private Function CheckPosOnSizeGrip( _
                 ByVal lKey As Long, _
                 ByVal lX As Long) As eSizeGripPos
    Dim lKeyX   As Long
    Dim lKeyW   As Long
    
    lKeyX = Int(m_tPattern.tKeys(lKey).dPos * m_tClientArea.lWidth) / (4 * m_lBeats) + m_tClientArea.lLeft
    lKeyW = Int(m_tPattern.tKeys(lKey).dLength * m_tClientArea.lWidth) / (4 * m_lBeats)
    
    If lX >= lKeyX And lX < lKeyX + SIZE_GRIP_AREA Then
        CheckPosOnSizeGrip = SGP_LEFT
    ElseIf lX >= lKeyX + lKeyW - SIZE_GRIP_AREA And lX < lKeyX + lKeyW Then
        CheckPosOnSizeGrip = SGP_RIGHT
    End If
    
End Function

Private Function GetKeyFromPos( _
                 ByVal lX As Long, _
                 ByVal lY As Long) As Long
    Dim lIndex  As Long
    Dim lNote   As Long
    Dim dPos    As Double
    
    lNote = PosYToNote(lY)
    dPos = PosXToQuarterPos(lX)
    
    For lIndex = 0 To m_tPattern.lNumOfKeys - 1
        If m_tPattern.tKeys(lIndex).lValue = lNote Then
            If dPos >= m_tPattern.tKeys(lIndex).dPos And _
                dPos < m_tPattern.tKeys(lIndex).dPos + m_tPattern.tKeys(lIndex).dLength Then
                GetKeyFromPos = lIndex
                Exit Function
            End If
        End If
    Next
                          
    GetKeyFromPos = -1
    
End Function

Private Function NoteToPosY( _
                 ByVal lValue As Long) As Long
    
    If m_tScroll.bEnabled Then
        NoteToPosY = (119 - lValue) * KEY_HEIGHT - m_tScroll.lValue
    Else
        NoteToPosY = (119 - lValue) * KEY_HEIGHT
    End If
    
End Function

Private Function QuarterLengthToPix( _
                 ByVal dQ As Double) As Long
    QuarterLengthToPix = m_tClientArea.lWidth * (dQ / (4 * m_lBeats))
End Function

Private Function QuarterPosToPix( _
                 ByVal dQ As Double) As Long
    
    QuarterPosToPix = m_tClientArea.lWidth * (dQ / (4 * m_lBeats)) + m_tClientArea.lLeft
              
    If QuarterPosToPix < m_tClientArea.lLeft Then
        QuarterPosToPix = m_tClientArea.lLeft
    ElseIf QuarterPosToPix >= m_tClientArea.lLeft + m_tClientArea.lWidth Then
        QuarterPosToPix = m_tClientArea.lLeft + m_tClientArea.lWidth - 1
    End If
    
End Function

Private Function PosYToNote( _
                 ByVal lY As Long) As Long
                 
    If m_tScroll.bEnabled Then
        PosYToNote = 119 - ((lY - m_tClientArea.lTop + m_tScroll.lValue) \ KEY_HEIGHT)
    Else
        PosYToNote = 119 - (lY - m_tClientArea.lTop \ KEY_HEIGHT)
    End If
    
    If PosYToNote < 0 Then
        PosYToNote = 0
    ElseIf PosYToNote > 119 Then
        PosYToNote = 119
    End If
    
End Function

Private Function PixLenToQuarterLen( _
                 ByVal lWidth As Long) As Double
    PixLenToQuarterLen = lWidth / m_tClientArea.lWidth * (4 * m_lBeats)
End Function

Private Function PosXToQuarterPos( _
                 ByVal lX As Long) As Double
                 
    lX = lX - m_tClientArea.lLeft
    
    If lX < 0 Then
        lX = 0
    ElseIf lX > m_tClientArea.lWidth Then
        lX = m_tClientArea.lWidth
    End If
    
    PosXToQuarterPos = lX / m_tClientArea.lWidth * (4 * m_lBeats)
    
End Function

Private Function PtIsInArea( _
                 ByVal lX As Long, _
                 ByVal lY As Long, _
                 ByRef tArea As tRectArea) As Boolean
    PtIsInArea = lX >= tArea.lLeft And lX < tArea.lLeft + tArea.lWidth And lY >= tArea.lTop And lY < tArea.lTop + tArea.lHeight
End Function

Private Sub UserControl_Resize()
    Dim lFullHeight As Long
    Dim lScrollSize As Long
    
    lFullHeight = KEY_HEIGHT * 120
    
    With m_tClientArea
    
        .lLeft = KEY_WIDTH
        .lTop = 0
        .lWidth = UserControl.ScaleWidth - KEY_WIDTH - SCROLL_WIDTH
        .lHeight = UserControl.ScaleHeight - TIME_BAR_SIZE
    
    End With
    
    With m_tScroll
    
        .lMax = lFullHeight - m_tClientArea.lHeight
        
        If .lValue > .lMax Then
            .lValue = .lMax
        End If
        
        lScrollSize = m_tClientArea.lHeight * m_tClientArea.lHeight \ lFullHeight
        
        If lScrollSize > m_tClientArea.lHeight Or .lMax <= 0 Then
            .bEnabled = False
            m_tClientArea.lWidth = m_tClientArea.lWidth + SCROLL_WIDTH
        Else
        
            .bEnabled = True
            
            .tArea.lLeft = m_tClientArea.lLeft + m_tClientArea.lWidth
            .tArea.lTop = m_tClientArea.lTop
            .tArea.lWidth = SCROLL_WIDTH
            .tArea.lHeight = m_tClientArea.lHeight
            
            .tTrack.lLeft = .tArea.lLeft + 2
            .tTrack.lTop = (.tArea.lHeight - 4 - lScrollSize) * (.lValue / .lMax) + 2
            .tTrack.lWidth = SCROLL_WIDTH - 4
            .tTrack.lHeight = lScrollSize - 4
            
        End If
        
    End With
    
    Refresh
    
End Sub

