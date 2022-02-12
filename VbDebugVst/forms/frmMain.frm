VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H00404040&
   Caption         =   "VbVst debugger by The trick"
   ClientHeight    =   8310
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10845
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPanel 
      Align           =   1  'Align Top
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   723
      TabIndex        =   0
      Top             =   0
      Width           =   10845
      Begin VbDebugVst.ctlButton cmdPlay 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Unicode MS"
            Size            =   14.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VbDebugVst.ctlButton cmdCreate 
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Create plugin"
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
      Begin VbDebugVst.ctlButton cmdShow 
         Height          =   495
         Left            =   2520
         TabIndex        =   3
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Show"
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
      Begin VB.Line linBorder 
         BorderColor     =   &H00E0E0E0&
         X1              =   352
         X2              =   676
         Y1              =   40
         Y2              =   40
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load audio..."
      End
      Begin VB.Menu mnuLoadState 
         Caption         =   "Load s&tate..."
      End
      Begin VB.Menu mnuSaveState 
         Caption         =   "&Save state..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuSong 
         Caption         =   "&Song"
      End
      Begin VB.Menu mnuPattern 
         Caption         =   "&Pattern"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "&Log"
      End
      Begin VB.Menu mnuPlugin 
         Caption         =   "&Plugin"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Begin VB.Menu mnuSetup 
         Caption         =   "&Setup..."
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuEventItemContext 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuEventItem 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuEventEditContext 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuClearEvent 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmMain.frm - main MDI form
' // by The trick, 2022
' //

Option Explicit

Private m_bPlayback As Boolean

Private Sub cmdPlay_Click()

    If Not m_bPlayback Then
        m_bPlayback = StartPlayback
    Else
        m_bPlayback = Not StopPlayback
    End If
    
    If m_bPlayback Then
        cmdPlay.Caption = ChrW$(&H25AE) & ChrW$(&H25AE)
    Else
        cmdPlay.Caption = ChrW$(&H25B6)
    End If
    
End Sub

Private Sub cmdCreate_Click()
    DestroyPlugin
    InitializePlugin
End Sub

Private Sub cmdShow_Click()
    mnuPlugin_Click
End Sub

Private Sub MDIForm_Load()
    cmdPlay.Caption = ChrW$(&H25B6)
End Sub

Private Sub MDIForm_Unload( _
            ByRef Cancel As Integer)
    UninitializeAll
End Sub

Private Sub mnuClearEvent_Click()
    frmSongEdit.ClearTrack
End Sub

Private Sub mnuEventItem_Click( _
            ByRef iIndex As Integer)
    frmSongEdit.SelectEvent iIndex
End Sub

Private Sub mnuLoad_Click()
    Dim sFile  As String
    
    sFile = GetOpenFile(Me.hWnd, "Open WAVE file", "WAVE PCM" & vbNullChar & "*.wav" & vbNullChar)
    If Len(sFile) = 0 Then
        Exit Sub
    End If
    
    If Not LoadAudioFile(sFile) Then
        MsgBox "Unable to load audio", vbCritical
    End If

    frmSongEdit.UpdateAudio

End Sub

Private Sub mnuLog_Click()
    frmLog.Show
    frmLog.ZOrder 0
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub mnuPattern_Click()
    frmPatternEdit.Show
    frmPatternEdit.ZOrder 0
End Sub

Private Sub mnuPlugin_Click()
    frmVSTSite.ShowPlugin
End Sub

' // The first DWORD is UNIQUE id of plugin
Private Sub mnuLoadState_Click()
    Dim sFile   As String
    Dim bData() As Byte
    Dim lSize   As LARGE_INTEGER
    Dim hFile   As Handle
    Dim lID     As Long
    Dim hr      As Long
    
    sFile = GetOpenFile(Me.hWnd, "Save state", "Binary file" & vbNullChar & "*.bin" & vbNullChar)
    If Len(sFile) = 0 Then
        Exit Sub
    End If
    
    hFile = CreateFile(sFile, GENERIC_READ, 0, ByVal NULL_PTR, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        Log "CreateFile failed 0x" & Hex$(GetLastError)
        Exit Sub
    End If
    
    If SetFilePointerEx(hFile, 0, 0, lSize, FILE_END) Then
        If CBool(lSize.HighPart) Or lSize.LowPart > 10000000 Or lSize.LowPart < 4 Then
            Log "Invalid file size"
        ElseIf SetFilePointerEx(hFile, 0, 0, ByVal NULL_PTR, FILE_BEGIN) = 0 Then
            Log "SetFilePointerEx failed 0x" & Hex$(GetLastError)
        ElseIf ReadFile(hFile, lID, 4, 0, ByVal NULL_PTR) = 0 Then
            Log "ReadFile failed 0x" & Hex$(GetLastError)
        ElseIf lID <> UniqueID Then
            Log "Invalid UniqueID"
        ElseIf lSize.LowPart > 4 Then
        
            ReDim bData(lSize.LowPart - 5)
            
            If ReadFile(hFile, bData(0), lSize.LowPart - 4, 0, ByVal NULL_PTR) = 0 Then
                Log "ReadFile failed 0x" & Hex$(GetLastError)
            Else
                
                hr = LoadState(bData, lSize.LowPart - 4)
                
                If hr < 0 Then
                    Log "LoadState failed 0x" & Hex$(hr)
                End If
                
            End If
            
        End If
    Else
        Log "SetFilePointerEx failed 0x" & Hex$(GetLastError)
    End If
    
    CloseHandle hFile
    
End Sub

' // The first DWORD is UNIQUE id of plugin
Private Sub mnuSaveState_Click()
    Dim sFile   As String
    Dim bOut()  As Byte
    Dim lSize   As Long
    Dim hFile   As Handle
    Dim hr      As Long
    
    If Not VstConnected Then
        Exit Sub
    End If
    
    hr = SaveState(bOut, lSize)
    If hr < 0 Then
        Log "SaveState failed 0x" & Hex$(hr)
        Exit Sub
    End If
    
    sFile = GetSaveFile(Me.hWnd, "Save state", "Binary file" & vbNullChar & "*.bin" & vbNullChar, "bin")
    If Len(sFile) = 0 Then
        Exit Sub
    End If
    
    hFile = CreateFile(sFile, GENERIC_READ Or GENERIC_WRITE, 0, ByVal NULL_PTR, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        Log "CreateFile failed 0x" & Hex$(GetLastError)
        Exit Sub
    End If
    
    If lSize > 0 Then
        If WriteFile(hFile, UniqueID, 4, 0, ByVal NULL_PTR) = 0 Then
            Log "WriteFile failed 0x" & Hex$(GetLastError)
        ElseIf WriteFile(hFile, bOut(0), lSize, 0, ByVal NULL_PTR) = 0 Then
            Log "WriteFile failed 0x" & Hex$(GetLastError)
        End If
    End If
    
    CloseHandle hFile
    
End Sub

Private Sub mnuSetup_Click()
    frmDebugOptions.Show vbModal
End Sub

Private Sub mnuSong_Click()
    frmSongEdit.Show
    frmSongEdit.ZOrder 0
End Sub

Private Sub picPanel_Resize()

    linBorder.X1 = 0
    linBorder.X2 = picPanel.ScaleWidth
    linBorder.Y1 = picPanel.ScaleHeight - 1
    linBorder.Y2 = picPanel.ScaleHeight - 1
    
End Sub
