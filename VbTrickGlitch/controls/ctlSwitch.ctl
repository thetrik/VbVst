VERSION 5.00
Begin VB.UserControl ctlSwitch 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   2580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
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
   ScaleHeight     =   172
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   325
End
Attribute VB_Name = "ctlSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' //
' // ctlSwitch.ctl - simple switch control
' // by The trick, 2022
' //

Option Explicit

Public Event OnClick( _
             ByVal bNewValue As Boolean)

Private m_sCaption  As String
Private m_bValue    As Boolean
Private m_tSWArea   As RECT

Public Property Get Value() As Boolean
    Value = m_bValue
End Property

Public Property Let Value( _
                    ByVal bValue As Boolean)
                        
    If bValue = m_bValue Then
        Exit Property
    End If
    
    m_bValue = bValue
    Refresh
    PropertyChanged "Value"
    
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_sCaption
End Property

Public Property Let Caption( _
                    ByVal sValue As String)
                        
    If sValue = m_sCaption Then
        Exit Property
    End If
    
    m_sCaption = sValue
    UserControl_Resize
    PropertyChanged "Caption"
    
End Property

Public Sub Refresh()
    Dim hdc     As HANDLE
    Dim tArea   As RECT
    Dim lMargin As Long
    
    hdc = UserControl.hdc
    SaveDC hdc
    
    SelectObject hdc, GetStockObject(DC_PEN)
    SelectObject hdc, GetStockObject(DC_BRUSH)
    
    SetDCPenColor hdc, &HE0E0E0
    SetDCBrushColor hdc, &H404040
    
    PatBlt hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, vbPatCopy
    
    If m_bValue Then
        SetDCBrushColor hdc, &HF09090
    End If
    
    RoundRect hdc, m_tSWArea.Left, m_tSWArea.Top, m_tSWArea.Right, m_tSWArea.Bottom, 6, 6
    
    If m_bValue Then
        SetRect tArea, m_tSWArea.Left + (m_tSWArea.Right - m_tSWArea.Left) \ 2, m_tSWArea.Top + 2, _
                       m_tSWArea.Right - 2, m_tSWArea.Bottom - 2
    Else
        SetRect tArea, m_tSWArea.Left + 2, m_tSWArea.Top + 2, m_tSWArea.Left + _
                      (m_tSWArea.Right - m_tSWArea.Left) \ 2, m_tSWArea.Bottom - 2
    End If

    SetDCBrushColor hdc, &HE0E0E0
    
    RoundRect hdc, tArea.Left, tArea.Top, tArea.Right, tArea.Bottom, 3, 3
    
    lMargin = UserControl.TextWidth(" ")
    
    SetRect tArea, 0, 0, m_tSWArea.Left - lMargin, UserControl.ScaleHeight
    
    DrawText hdc, m_sCaption, Len(m_sCaption), tArea, DT_SINGLELINE Or DT_VCENTER Or DT_END_ELLIPSIS Or DT_RIGHT
    
    RestoreDC hdc, -1
    InvalidateRect UserControl.hWnd, ByVal 0&, 0
    
End Sub

Private Sub UserControl_MouseDown( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)

    If iButton = vbLeftButton Then

        If PtInRect(m_tSWArea, fX, fY) Then
        
            m_bValue = Not m_bValue
            RaiseEvent OnClick(m_bValue)
            Refresh
            
        End If
        
    End If
    
End Sub

Private Sub UserControl_ReadProperties( _
            ByRef cPropBag As PropertyBag)
    m_bValue = cPropBag.ReadProperty("Value", False)
    m_sCaption = cPropBag.ReadProperty("Caption", vbNullString)
End Sub

Private Sub UserControl_WriteProperties( _
            ByRef cPropBag As PropertyBag)
    cPropBag.WriteProperty "Value", m_bValue, False
    cPropBag.WriteProperty "Caption", m_sCaption, vbNullString
End Sub

Private Sub UserControl_Resize()
    Dim lCapWidth   As Long
    Dim lMargin     As Long
    Dim lPosX       As Long
    Dim lHeight     As Long
    Dim lWidth      As Long
    
    lHeight = UserControl.ScaleHeight
    lWidth = lHeight * 2
    
    lCapWidth = UserControl.TextWidth(m_sCaption)
    lMargin = UserControl.TextWidth(" ")
    
    If lCapWidth + lWidth + lMargin > UserControl.ScaleWidth Then
        lCapWidth = UserControl.ScaleWidth - lWidth - lMargin
    End If
    
    SetRect m_tSWArea, lCapWidth + lMargin, 0, lCapWidth + lWidth + lMargin, UserControl.ScaleHeight
    
    Refresh
    
End Sub
