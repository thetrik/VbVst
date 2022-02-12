VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Log"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6645
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Log"
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   6645
   Begin VB.TextBox txtLog 
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
      ForeColor       =   &H00E0E0E0&
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // frmLog.frm - logging
' // by The trick, 2022
' //

Option Explicit

Private m_lLogLen   As Long

Public Sub PutLog( _
           ByRef sMsg As String)
               
    txtLog.SelStart = m_lLogLen
    txtLog.SelText = sMsg & vbNewLine
    m_lLogLen = m_lLogLen + Len(sMsg) + 2
                   
    If Me.Visible = False Then
        Me.Show
    End If
                   
End Sub

Private Sub Form_QueryUnload( _
            ByRef iCancel As Integer, _
            ByRef iUnloadMode As Integer)
    
    If iUnloadMode = vbFormControlMenu Then
        Me.Hide
        iCancel = True
    End If
    
End Sub

Private Sub Form_Resize()
    If Me.ScaleWidth > 100 And Me.ScaleHeight > 100 Then
        txtLog.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End If
End Sub
