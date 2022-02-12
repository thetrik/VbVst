VERSION 5.00
Begin VB.UserControl ctlKnob 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
End
Attribute VB_Name = "ctlKnob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' //
' // ctlKnob.ctl - simple knob control
' // by The trick, 2022
' //

Option Explicit

Public Event Changed( _
             ByVal lNewValue As Long)

Private m_lMin          As Long
Private m_lMax          As Long
Private m_lValue        As Long
Private m_lSmallChange  As Long
Private m_lLargeChange  As Long
Private m_lZeroValue    As Long
Private m_lSavedValue   As Long
Private m_lKnobSize     As Long
Private m_lTextSize     As Long
Private m_bShowValue    As Boolean
Private m_bTrackMouse   As Boolean
Private m_tMouseOffset  As POINT
Private m_hGraphics     As Handle
Private m_hGPToken      As Handle
Private m_hGPPen        As Handle
Private m_hGPBrush      As Handle

Public Property Get ShowValue() As Boolean
    ShowValue = m_bShowValue
End Property
Public Property Let ShowValue( _
                    ByVal bValue As Boolean)
    m_bShowValue = bValue
    Refresh
    PropertyChanged "ShowValue"
End Property

Public Property Get LargeChange() As Long
    LargeChange = m_lLargeChange
End Property
Public Property Let LargeChange( _
                    ByVal lValue As Long)
    m_lLargeChange = lValue
    PropertyChanged "LargeChange"
End Property

Public Property Get SmallChange() As Long
    SmallChange = m_lSmallChange
End Property
Public Property Let SmallChange( _
                    ByVal lValue As Long)
    m_lSmallChange = lValue
    PropertyChanged "SmallChange"
End Property

Public Property Get ZeroValue() As Long
    ZeroValue = m_lZeroValue
End Property
Public Property Let ZeroValue( _
                    ByVal lValue As Long)
                    
    If lValue > m_lMax Then
        lValue = m_lMax
    ElseIf lValue < m_lMin Then
        lValue = m_lMin
    End If

    If lValue = m_lMin Then
        Exit Property
    End If
    
    m_lZeroValue = lValue
    Refresh
    PropertyChanged "ZeroValue"
    
End Property

Public Property Get Min() As Long
    Min = m_lMin
End Property
Public Property Let Min( _
                    ByVal lValue As Long)
                    
    If lValue > m_lMax Then
        m_lMax = lValue
    End If
    
    If m_lValue < lValue Then
        m_lValue = lValue
        m_lSavedValue = lValue
    End If
    
    If m_lZeroValue < lValue Then
        m_lZeroValue = lValue
    End If
    
    If lValue = m_lMin Then
        Exit Property
    End If
    
    m_lMin = lValue
    Refresh
    PropertyChanged "Min"
    
End Property

Public Property Get Max() As Long
    Max = m_lMax
End Property
Public Property Let Max( _
                    ByVal lValue As Long)
                    
    If lValue < m_lMin Then
        m_lMin = lValue
    End If
    
    If m_lValue > lValue Then
        m_lValue = lValue
        m_lSavedValue = lValue
    End If
    
    If m_lZeroValue > lValue Then
        m_lZeroValue = lValue
    End If
    
    If lValue = m_lMax Then
        Exit Property
    End If
    
    m_lMax = lValue
    Refresh
    PropertyChanged "Max"
    
End Property

Public Property Get Value() As Long
    Value = m_lValue
End Property
Public Property Let Value( _
                    ByVal lValue As Long)
                    
    If lValue > m_lMax Then
        lValue = m_lMax
    ElseIf lValue < m_lMin Then
        lValue = m_lMin
    End If
    
    If lValue = m_lValue Then
        Exit Property
    End If
    
    m_lValue = lValue
    m_lSavedValue = lValue
    
    Refresh
    PropertyChanged "Value"
    
End Property

Public Sub Refresh()
    Dim fValue      As Single
    Dim fAngle      As Single
    Dim fPosX       As Single
    Dim fPosY       As Single
    Dim fPtArea     As Single
    Dim fHalfKnob   As Single
    Dim fHalfText   As Single
    Dim lIndex      As Long
    Dim lCount      As Long
    Dim fDelta      As Single
    Dim sCaption    As String
    Dim tRC         As RECT
    Dim fArcStart   As Single
    Dim fArcSweep   As Single
    
    fHalfKnob = m_lKnobSize / 2
    fHalfText = m_lTextSize / 2
    fValue = (m_lValue - m_lMin) / (m_lMax - m_lMin)
    
    GdipGraphicsClear m_hGraphics, &HFF404040
    GdipSetSolidFillColor m_hGPBrush, &HFFE0E0E0
    GdipFillEllipse m_hGraphics, m_hGPBrush, m_lTextSize, m_lTextSize, m_lKnobSize, m_lKnobSize
    
    GdipSetPenColor m_hGPPen, &HFFE0E0E0
    GdipSetPenWidth m_hGPPen, 1
    GdipDrawArc m_hGraphics, m_hGPPen, fHalfText, fHalfText, m_lKnobSize + m_lTextSize, m_lKnobSize + m_lTextSize, -225, 270

    lCount = m_lMax - m_lMin + 1
    
    If lCount > 11 Then
        lCount = 11
    End If
    
    If lCount > 1 Then
        
        fDelta = 270 / (lCount - 1)
        
        For lIndex = 0 To lCount - 1
            
            GdipRotateWorldTransform m_hGraphics, lIndex * fDelta - 225, MatrixOrderAppend
            GdipTranslateWorldTransform m_hGraphics, m_lTextSize + fHalfKnob, m_lTextSize + fHalfKnob, MatrixOrderAppend
            GdipDrawLine m_hGraphics, m_hGPPen, fHalfKnob, 0, fHalfKnob + fHalfText, 0
            GdipResetWorldTransform m_hGraphics
            
        Next
        
    End If
    
    fArcStart = 270 * (m_lZeroValue - m_lMin) / (m_lMax - m_lMin) - 225
    fArcSweep = 270 * fValue - fArcStart - 225
    
    GdipSetPenColor m_hGPPen, &HFF9090FF
    GdipSetPenWidth m_hGPPen, 4
    GdipDrawArc m_hGraphics, m_hGPPen, fHalfText, fHalfText, m_lKnobSize + m_lTextSize, _
                m_lKnobSize + m_lTextSize, fArcStart, fArcSweep
    
    GdipSetSolidFillColor m_hGPBrush, &HFF4040FF
    
    fAngle = 4.71238898038469 * fValue - 3.92699081698724
    fPtArea = (m_lKnobSize - 10) / 2
    
    fPosX = Cos(fAngle) * fPtArea + m_lTextSize + fHalfKnob
    fPosY = Sin(fAngle) * fPtArea + m_lTextSize + fHalfKnob
    
    GdipFillEllipse m_hGraphics, m_hGPBrush, fPosX - 2, fPosY - 2, 5, 5
    
    If m_bShowValue Then
        
        sCaption = CStr(m_lValue)
        
        SetRect tRC, m_lTextSize, m_lTextSize + m_lKnobSize, m_lTextSize + m_lKnobSize, UserControl.ScaleHeight
    
        DrawText UserControl.hdc, sCaption, Len(sCaption), tRC, DT_CENTER Or DT_END_ELLIPSIS
        
    End If
    
    InvalidateRect UserControl.hWnd, ByVal 0&, 1
    
End Sub

Private Sub UserControl_InitProperties()
    
    m_lMax = 100
    m_lSmallChange = 1
    m_lLargeChange = 10
    m_bShowValue = True
    
End Sub

Private Sub UserControl_MouseDown( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)

    If iButton = vbLeftButton Then

        m_tMouseOffset.x = fX
        m_tMouseOffset.y = fY
        
        m_lSavedValue = m_lValue
        m_bTrackMouse = True
        
    Else
        If m_lValue <> m_lSavedValue Then
        
            m_lValue = m_lSavedValue
            m_bTrackMouse = False
            RaiseEvent Changed(m_lValue)
            Refresh
            
        End If
    End If
    
End Sub

Private Sub UserControl_MouseMove( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    Dim lNewVal As Long
    Dim lScrH   As Long
    Dim tPt     As POINT
    
    If m_bTrackMouse Then
        
        tPt.x = fX: tPt.y = fY
        
        If iShift And vbCtrlMask Then
            lNewVal = m_lValue + ((m_tMouseOffset.y - tPt.y) \ 10) * m_lSmallChange
        Else
            lNewVal = m_lValue + ((m_tMouseOffset.y - tPt.y) \ 10) * m_lLargeChange
        End If
        
        If lNewVal < m_lMin Then
            lNewVal = m_lMin
        ElseIf lNewVal > m_lMax Then
            lNewVal = m_lMax
        End If
        
        If lNewVal <> m_lValue Then
        
            m_lValue = lNewVal
            RaiseEvent Changed(lNewVal)
            Refresh
            
            m_tMouseOffset.x = tPt.x
            m_tMouseOffset.y = tPt.y
        
        End If
        
        ClientToScreen UserControl.hWnd, tPt
        lScrH = GetSystemMetrics(SM_CYSCREEN)
        
        ' // Wrap mouse
        If tPt.y = 0 Then
            
            tPt.y = lScrH - 2
            SetCursorPos tPt.x, tPt.y
            ScreenToClient UserControl.hWnd, tPt
            m_tMouseOffset.x = tPt.x
            m_tMouseOffset.y = tPt.y
            
        ElseIf tPt.y = lScrH - 1 Then
        
            tPt.y = 1
            SetCursorPos tPt.x, tPt.y
            ScreenToClient UserControl.hWnd, tPt
            m_tMouseOffset.x = tPt.x
            m_tMouseOffset.y = tPt.y
            
        End If
        
    End If
    
End Sub

Private Sub UserControl_MouseUp( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    
    If m_bTrackMouse Then
        m_lSavedValue = m_lValue
        m_bTrackMouse = False
    End If
    
End Sub

Private Sub UserControl_Initialize()
    Dim tGPInput    As GdiplusStartupInput

    tGPInput.GdiplusVersion = 1
    
    If GdiplusStartup(m_hGPToken, tGPInput) Then
        GoTo CleanUp
    End If
    
    If GdipCreatePen1(&HFFE0E0E0, 4, UnitPixel, m_hGPPen) Then
        GoTo CleanUp
    End If
    
    If GdipCreateSolidFill(&HFFE0E0E0, m_hGPBrush) Then
        GoTo CleanUp
    End If
    
    Exit Sub
    
CleanUp:
    
    If m_hGPBrush Then
        GdipDeleteBrush m_hGPBrush
    End If
    
    If m_hGPPen Then
        GdipDeletePen m_hGPPen
    End If
    
    If m_hGPToken Then
        GdiplusShutdown m_hGPToken
    End If
    
End Sub

Private Sub UserControl_ReadProperties( _
            ByRef cPropBag As PropertyBag)
    
    ' // No error checking
    m_lZeroValue = cPropBag.ReadProperty("ZeroValue", 0)
    m_lLargeChange = cPropBag.ReadProperty("LargeChange", 10)
    m_lSmallChange = cPropBag.ReadProperty("SmallChange", 0)
    m_lMax = cPropBag.ReadProperty("Max", 100)
    m_lMin = cPropBag.ReadProperty("Min", 0)
    m_lValue = cPropBag.ReadProperty("Value", 0)
    m_bShowValue = cPropBag.ReadProperty("ShowValue", True)
    
End Sub

Private Sub UserControl_WriteProperties( _
            ByRef cPropBag As PropertyBag)
    
    cPropBag.WriteProperty "ZeroValue", m_lZeroValue, 0
    cPropBag.WriteProperty "LargeChange", m_lLargeChange, 0
    cPropBag.WriteProperty "SmallChange", m_lSmallChange, 0
    cPropBag.WriteProperty "Max", m_lMax, 0
    cPropBag.WriteProperty "Min", m_lMin, 0
    cPropBag.WriteProperty "Value", m_lValue, 0
    cPropBag.WriteProperty "ShowValue", m_bShowValue, True
    
End Sub

Private Sub UserControl_Terminate()
    
    If m_hGraphics Then
        GdipDeleteGraphics m_hGraphics
    End If
    
    If m_hGPBrush Then
        GdipDeleteBrush m_hGPBrush
    End If
    
    If m_hGPPen Then
        GdipDeletePen m_hGPPen
    End If
    
    If m_hGPToken Then
        GdiplusShutdown m_hGPToken
    End If
    
End Sub

Private Sub UserControl_Resize()

    If m_hGPToken Then
    
        If m_hGraphics Then
            GdipDeleteGraphics m_hGraphics
        End If
        
        If GdipCreateFromHDC(UserControl.hdc, m_hGraphics) Then
            m_hGraphics = 0
        End If
        
        GdipSetSmoothingMode m_hGraphics, SmoothingModeAntiAlias
        
    End If

    m_lTextSize = UserControl.TextHeight("0")
    
    If UserControl.ScaleWidth > UserControl.ScaleHeight - m_lTextSize Then
        m_lKnobSize = UserControl.ScaleHeight - m_lTextSize * 2
    Else
        m_lKnobSize = UserControl.ScaleWidth - m_lTextSize * 2
    End If

    If m_lKnobSize < 10 Then
        m_lKnobSize = 10
    End If
    
    Refresh
    
End Sub
