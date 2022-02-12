VERSION 5.00
Begin VB.UserControl ctlButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
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
Attribute VB_Name = "ctlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // ctlButton.ctl - very simple button
' // By The trick, 2022
' //

Option Explicit

Public Event Click()

Private m_bPressed  As Boolean
Private m_sCaption  As String

Private WithEvents m_cFont  As StdFont
Attribute m_cFont.VB_VarHelpID = -1

Public Property Get Font() As StdFont
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property
Public Property Set Font( _
                    ByVal cValue As StdFont)
                    
    Set UserControl.Font = cValue
    Set m_cFont = cValue
    Redraw
    
End Property

Public Property Let Caption( _
                    ByRef sValue As String)
    m_sCaption = sValue
    Redraw
    PropertyChanged "Caption"
End Property
Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = m_sCaption
End Property

Public Sub Redraw()
    Dim tRC     As RECT
    Dim tRcTxt  As RECT
    Dim hdc     As Handle
    
    hdc = UserControl.hdc
    
    SaveDC hdc
    SelectObject hdc, GetStockObject(DC_BRUSH)
    
    SetRect tRC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    SetDCBrushColor hdc, &H404040
    
    PatBlt hdc, tRC.Left, tRC.Top, tRC.Right - tRC.Left, tRC.Bottom - tRC.Top, vbPatCopy
    
    If m_bPressed Then
        SetDCBrushColor hdc, &HF09090
        InflateRect tRC, -2, -2
    Else
        SetDCBrushColor hdc, &H606060
    End If
    
    Rectangle hdc, tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
    
    tRcTxt = tRC
    
    DrawText hdc, m_sCaption, Len(m_sCaption), tRcTxt, DT_CALCRECT
    
    OffsetRect tRcTxt, ((tRC.Right - tRC.Left) - (tRcTxt.Right - tRcTxt.Left)) \ 2, _
                       ((tRC.Bottom - tRC.Top) - (tRcTxt.Bottom - tRcTxt.Top)) \ 2
    InflateRect tRC, -2, -2
    IntersectRect tRcTxt, tRcTxt, tRC

    DrawText hdc, m_sCaption, Len(m_sCaption), tRcTxt, DT_END_ELLIPSIS
    
    RestoreDC hdc, -1
    
    InvalidateRect UserControl.hWnd, ByVal NULL_PTR, 0
    
End Sub

Private Sub m_cFont_FontChanged( _
            ByVal sPropertyName As String)
    Redraw
End Sub

Private Sub UserControl_AccessKeyPress( _
            ByRef iKeyAscii As Integer)
    If iKeyAscii = 13 Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_MouseDown( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
    
    If iButton = vbLeftButton Then
        m_bPressed = True
        Redraw
    End If
    
End Sub

Private Sub UserControl_MouseUp( _
            ByRef iButton As Integer, _
            ByRef iShift As Integer, _
            ByRef fX As Single, _
            ByRef fY As Single)
            
    If iButton = vbLeftButton Then
    
        m_bPressed = False
        Redraw
        
        If fX >= 0 And fY >= 0 And fX < UserControl.ScaleWidth And fY < UserControl.ScaleHeight Then
            RaiseEvent Click
        End If
        
    End If
    
End Sub

Private Sub UserControl_ReadProperties( _
            ByRef cPropBag As PropertyBag)
            
    m_sCaption = cPropBag.ReadProperty("Caption", vbNullString)
    Set UserControl.Font = cPropBag.ReadProperty("Font", Ambient.Font)
    Redraw
    
End Sub

Private Sub UserControl_Resize()
    Redraw
End Sub

Private Sub UserControl_WriteProperties( _
            ByRef cPropBag As PropertyBag)
    
    cPropBag.WriteProperty "Caption", m_sCaption, vbNullString
    cPropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
    
End Sub
