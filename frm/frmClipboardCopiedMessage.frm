VERSION 5.00
Begin VB.Form frmClipboardCopiedMessage 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   984
   ClientLeft      =   4632
   ClientTop       =   2136
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClipboardCopiedMessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   984
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   144
      Top             =   180
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Text copied"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "frmClipboardCopiedMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblMessage.Caption = " " & GetLocalizedString(efnGUIStr_frmClipboardCopiedMessage_lblMessage_Caption) & " "
End Sub

Private Sub tmrHide_Timer()
    tmrHide.Enabled = False
    If IsFormLoaded(Me) Then
        On Error Resume Next
        ShowWindow Me.hWnd, SW_HIDE
'        Me.Hide
    End If
End Sub

Public Sub ShowMessage()
    Dim iM As POINTAPI
    
    GetCursorPos iM
    lblMessage.Move 0, 0
    Me.Width = lblMessage.Width
    Me.Height = lblMessage.Height
    Me.Move iM.x * Screen.TwipsPerPixelX - Me.Width / 2, iM.y * Screen.TwipsPerPixelY + 300
    If (Me.Left + Me.Width > Screen.Width) Then
        Me.Left = Screen.Width - Me.Width
    End If
    If Me.Left < 0 Then
        Me.Left = 0
    End If
    If (Me.Top + Me.Height) > ScreenUsableHeight Then
        Me.Top = ScreenUsableHeight - Me.Height
    End If
    If Me.Top < 0 Then
        Me.Top = 0
    End If
    On Error Resume Next
    ShowNoActivate Me, , , True
    tmrHide.Enabled = True
    DoEvents
    If Not Me.Visible Then
        Me.Show 1
    End If
    
    Do Until tmrHide.Enabled = False
        DoEvents
    Loop
    If IsFormLoaded(Me) Then
        Unload Me
    End If
    On Error GoTo 0
End Sub
