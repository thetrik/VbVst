VERSION 5.00
Begin VB.Form frmDebugOptions 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug options"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3930
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDebugOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   262
   StartUpPosition =   3  'Windows Default
   Begin VbDebugVst.ctlSwitch chkShowEditor 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   661
      Caption         =   "Use editor"
   End
   Begin VbDebugVst.ctlButton cmdSave 
      Default         =   -1  'True
      Height          =   495
      Left            =   900
      TabIndex        =   3
      Top             =   1380
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   873
      Caption         =   "Save"
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
   Begin VB.TextBox txtProgID 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3675
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00E0E0E0&
      Height          =   405
      Left            =   105
      Top             =   345
      Width           =   3705
   End
   Begin VB.Label lblProgID 
      BackColor       =   &H00404040&
      Caption         =   "ProgID:"
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3675
   End
End
Attribute VB_Name = "frmDebugOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmDebugOptions.frm - debugging options
' // by The trick, 2022
' //

Option Explicit

Private Sub cmdSave_Click()
    Dim sProgId As String
    
    sProgId = txtProgID.Text
    
    If Len(sProgId) = 0 Or Len(sProgId) > 39 Then
        MsgBox "Invalid ProgID", vbCritical
        Exit Sub
    End If
    
    Me.Hide
    
    UseEditor = chkShowEditor.Value
    ProgID = sProgId
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    txtProgID.Text = ProgID
    chkShowEditor.Value = UseEditor
End Sub
