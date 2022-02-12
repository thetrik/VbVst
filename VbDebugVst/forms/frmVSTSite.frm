VERSION 5.00
Begin VB.Form frmVSTSite 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9360
   Icon            =   "frmVSTSite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   Begin VbDebugVst.ctlKnob knbParam 
      Height          =   1035
      Index           =   0
      Left            =   3780
      TabIndex        =   1
      Top             =   300
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1826
      LargeChange     =   10
      SmallChange     =   1
      Max             =   1000
      ShowValue       =   0   'False
   End
   Begin VB.Timer tmrIdle 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   2400
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   60
      ScaleHeight     =   153
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   0
      Top             =   60
      Width           =   3675
   End
   Begin VB.Label lblParamDisplay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   3780
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblParamName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   3780
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmVSTSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmVSTSite.frm - VST container (for editorless plugins)
' // by The trick, 2022
' //

Option Explicit

Private m_cPlugin       As IVBVstEffect_dbg
Private m_bHasEditor    As Boolean

Public Property Set Plugin( _
                    ByVal cValue As IVBVstEffect_dbg)
                    
    Set m_cPlugin = cValue
    
    InitializeSite
    
    If cValue Is Nothing Or m_bHasEditor Then
        Me.Hide
    Else
        Me.Show
    End If
    
End Property

Public Sub ParameterChanged( _
           ByVal lIndex As Long, _
           ByVal fValue As Single)
    Dim sDisplay    As String
    Dim sLabel      As String
    Dim hr          As Long
    
    If m_cPlugin Is Nothing Then
        Exit Sub
    End If
    
    If Not m_bHasEditor Then
    
        knbParam(lIndex).Value = fValue * 1000

        hr = m_cPlugin.ParamDisplay(lIndex, sDisplay)
        If hr < 0 Then
            Log "ParamDisplay failed 0x" & Hex$(hr)
        End If
        
        hr = m_cPlugin.ParamLabel(lIndex, sLabel)
        If hr < 0 Then
            Log "ParamLabel failed 0x" & Hex$(hr)
        End If
        
        lblParamDisplay(lIndex).Caption = sDisplay & " " & sLabel
        
    End If

End Sub

Public Sub ShowPlugin()
    If VstConnected Then
        If Not m_bHasEditor Then
            Me.Show
            Me.ZOrder 0
        Else
            SetForegroundWindow VSTContainerHandle
        End If
    End If
End Sub

Public Sub RequestClose()
    Dim hr  As Long

    If m_bHasEditor Then
        If Not m_cPlugin Is Nothing Then
            hr = m_cPlugin.EditorClose
            Log "Editor has been closed 0x" & Hex$(hr)
        End If
    End If
    
    DestroyPlugin

End Sub

Private Sub InitializeSite()
    Dim tRect       As ERect
    Dim hr          As Long
    Dim bHasEditor  As Boolean
    Dim bResult     As Boolean
    
    If Not m_cPlugin Is Nothing Then
        
        If UseEditor Then
        
            hr = m_cPlugin.HasEditor(bHasEditor)
            If hr < 0 Then
                Log "HasEditor failed 0x" & Hex$(hr)
                Exit Sub
            End If
        
        End If
        
        If bHasEditor Then
        
            hr = m_cPlugin.EditorRect(tRect)
            If hr < 0 Then
                Log "EditorRect failed 0x" & Hex$(hr)
                Exit Sub
            End If
            
            If tRect.wRight <= tRect.wLeft Or tRect.wBottom <= tRect.wTop Then
                Log "EditorRect returns invalid rectangle"
                Exit Sub
            End If
                        
            hr = m_cPlugin.EditorOpen(VSTContainerHandle, bResult)
            If hr < 0 Then
                Log "EditorOpen failed 0x" & Hex$(hr)
                Exit Sub
            End If
            
            If Not bResult Then
                Log "EditorOpen returns false"
                Exit Sub
            End If
            
            SetWindowText VSTContainerHandle, EffectName
            
            m_bHasEditor = True
            tmrIdle.Enabled = True
            
        Else
        
            InitializeWithoutEditor
            tmrIdle.Enabled = False
            m_bHasEditor = False
            
        End If
        
    Else
        tmrIdle.Enabled = False
    End If

End Sub

Private Sub InitializeWithoutEditor()
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim lSingleWidth    As Long
    Dim lSingleHeight   As Long
    Dim lIndex          As Long
    Dim lItemsPerLine   As Long
    Dim lTotalLines     As Long
    Dim lOffsetX        As Long
    Dim lOffsetY        As Long
    Dim lLineIndex      As Long
    Dim fValue          As Single
    Dim sName           As String
    Dim sDisplay        As String
    Dim sLabel          As String
    Dim hr              As Long
    
    Me.Caption = EffectName
    
    For lIndex = knbParam.LBound To knbParam.UBound
    
        If lIndex <> 0 Then
            Unload knbParam(lIndex)
            Unload lblParamName(lIndex)
            Unload lblParamDisplay(lIndex)
        End If
        
    Next
    
    knbParam(0).Visible = False
    lblParamName(0).Visible = False
    lblParamDisplay(0).Visible = False
    
    lSingleWidth = knbParam(0).Width
    lSingleHeight = knbParam(0).Height + lblParamName(0).Height + lblParamDisplay(0).Height
    
    lItemsPerLine = ((Screen.Width / Screen.TwipsPerPixelX) \ 2) \ lSingleWidth
    
    lTotalLines = (NumOfParams \ lItemsPerLine) + 1
    
    If lItemsPerLine > NumOfParams Then
        lItemsPerLine = NumOfParams
    End If
    
    If lItemsPerLine < 2 Then
        picContainer.Move 0, 0, lSingleWidth * 2, lSingleHeight
    Else
        picContainer.Move 0, 0, lItemsPerLine * lSingleWidth, lTotalLines * lSingleHeight
    End If
    
    For lIndex = 0 To NumOfParams - 1
        
        lLineIndex = lIndex \ lItemsPerLine
        
        lOffsetY = lLineIndex * lSingleHeight
        
        If lLineIndex = lTotalLines - 1 Then
            lOffsetX = ((lItemsPerLine - (NumOfParams - lItemsPerLine * (lTotalLines - 1))) * lSingleWidth) \ 2
        Else
            lOffsetX = 0
        End If
        
        If lIndex <> 0 Then
        
            Load knbParam(lIndex)
            Load lblParamName(lIndex)
            Load lblParamDisplay(lIndex)
            
        End If
        
        Set knbParam(lIndex).Container = picContainer
        Set lblParamName(lIndex).Container = picContainer
        Set lblParamDisplay(lIndex).Container = picContainer
        
        lblParamName(lIndex).Move lOffsetX + (lIndex Mod lItemsPerLine) * lSingleWidth, lOffsetY
        knbParam(lIndex).Move lblParamName(lIndex).Left, lOffsetY + lblParamName(lIndex).Height
        lblParamDisplay(lIndex).Move knbParam(lIndex).Left, lOffsetY + knbParam(lIndex).Height + lblParamName(lIndex).Height
        
        hr = m_cPlugin.ParamName(lIndex, sName)
        If hr < 0 Then
            Log "ParamName failed 0x" & Hex$(hr)
        End If
        
        lblParamName(lIndex).Caption = sName
        
        hr = m_cPlugin.ParamDisplay(lIndex, sDisplay)
        If hr < 0 Then
            Log "ParamDisplay failed 0x" & Hex$(hr)
        End If
        
        hr = m_cPlugin.ParamLabel(lIndex, sLabel)
        If hr < 0 Then
            Log "ParamLabel failed 0x" & Hex$(hr)
        End If
        
        lblParamDisplay(lIndex).Caption = sDisplay & " " & sLabel
        
        hr = m_cPlugin.ParamValue_get(lIndex, fValue)
        If hr < 0 Then
            Log "ParamValue_get failed 0x" & Hex$(hr)
        End If
        
        knbParam(lIndex).Value = fValue * 1000
        
        knbParam(lIndex).Visible = True
        lblParamName(lIndex).Visible = True
        lblParamDisplay(lIndex).Visible = True
    
    Next
    
    Me.Width = picContainer.Width * Screen.TwipsPerPixelX + (Me.Width - Me.ScaleWidth * Screen.TwipsPerPixelX)
    Me.Height = picContainer.Height * Screen.TwipsPerPixelY + (Me.Height - Me.ScaleHeight * Screen.TwipsPerPixelY)
    
End Sub

Private Sub Form_QueryUnload( _
            ByRef Cancel As Integer, _
            ByRef UnloadMode As Integer)
    
    If UnloadMode = vbFormControlMenu And Not m_cPlugin Is Nothing Then
        If MsgBox("Are you sure?", vbQuestion Or vbYesNo) = vbNo Then
            Cancel = 1
        End If
    End If
    
End Sub

Private Sub Form_Unload( _
            ByRef Cancel As Integer)
    Dim hr  As Long
    
    If m_bHasEditor Then
        If Not m_cPlugin Is Nothing Then
            hr = m_cPlugin.EditorClose
            Log "Editor has been closed 0x" & Hex$(hr)
        End If
    End If
    
    DestroyPlugin
    
End Sub

Private Sub knbParam_Changed( _
            ByRef iIndex As Integer, _
            ByVal lNewValue As Long)
    Dim sDisplay    As String
    Dim sLabel      As String
    Dim hr          As Long

    hr = m_cPlugin.ParamValue_put(iIndex, knbParam(iIndex).Value / 1000)
    
    If hr < 0 Then
        Log "ParamValue_put failed 0x" & Hex$(hr)
    End If
    
    hr = m_cPlugin.ParamDisplay(iIndex, sDisplay)
    If hr < 0 Then
        Log "ParamDisplay failed 0x" & Hex$(hr)
    End If
    
    hr = m_cPlugin.ParamLabel(iIndex, sLabel)
    If hr < 0 Then
        Log "ParamLabel failed 0x" & Hex$(hr)
    End If
    
    lblParamDisplay(iIndex).Caption = sDisplay & " " & sLabel
    
End Sub

Private Sub tmrIdle_Timer()
    Dim hr          As Long
    Dim lRecords    As Long
    Dim lIndex      As Long
    
    hr = GetRemoteHostState()
    If hr < 0 Then
        Log "GetRemoteHostState failed 0x" & Hex$(hr)
    ElseIf hr = 0 Then
        Unload Me
    ElseIf hr = 2 Then
        ' // Paused
    ElseIf hr = 1 Then
        
        hr = m_cPlugin.EditorIdle(g_tAutomation(), lRecords)
        If hr < 0 Then
            Log "EditorIdle failed 0x" & Hex$(hr)
            Exit Sub
        End If
        
        For lIndex = 0 To lRecords - 1
            Log "New automation event. [param: " & CStr(g_tAutomation(lIndex).lParamIndex) & _
                ", value: " & CStr(g_tAutomation(lIndex).fParamValue) & "]"
        Next
       
     End If
     
End Sub
