VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.0#0"; "vbExtra2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3996
   ClientLeft      =   1740
   ClientTop       =   2640
   ClientWidth     =   4164
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3996
   ScaleWidth      =   4164
   Begin vbExtra.TrayIcon TrayIcon1 
      Left            =   3168
      Top             =   540
      _ExtentX        =   720
      _ExtentY        =   720
   End
   Begin VB.Menu mnuTayPopup 
      Caption         =   "mnuTayPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore window"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close program"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mWindowState As Long
Private mBalloonTipShown As Boolean

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        TrayIcon1.Create App.Title, Me.Icon
        Me.Visible = False
        If Not mBalloonTipShown Then
            TrayIcon1.BalloonTip "The program is here", vxBTSInfo, App.Title
            mBalloonTipShown = True
        End If
    Else
        mWindowState = Me.WindowState
    End If
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuRestore_Click()
    Me.WindowState = mWindowState
    Me.Visible = True
    TrayIcon1.Remove
End Sub

Private Sub TrayIcon1_TrayClick(Button As vbExtra.vbExTrayIconMouseEventConstants)
    If Button = vxMERightButtonClick Then
        TrayIcon1.PopupMenu mnuTayPopup
    End If
End Sub

Private Sub TrayIcon1_DblClick()
    mnuRestore_Click
End Sub
