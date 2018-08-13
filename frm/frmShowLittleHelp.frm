VERSION 5.00
Begin VB.Form frmShowLittleHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3384
   ClientLeft      =   4740
   ClientTop       =   7428
   ClientWidth     =   4596
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShowLittleHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3384
   ScaleWidth      =   4596
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLH1 
      BackColor       =   &H00FFFFF9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5355
   End
   Begin VB.Label lblLH 
      BackColor       =   &H00FFFFF9&
      Caption         =   "Help text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   570
      TabIndex        =   1
      Top             =   2910
      Width           =   1725
   End
End
Attribute VB_Name = "frmShowLittleHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mToolTipEx As ToolTipEx
Attribute mToolTipEx.VB_VarHelpID = -1
Private mClosedByCode As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not mClosedByCode Then
        If Not mToolTipEx Is Nothing Then
            mToolTipEx.RaiseEventBeforeClose
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lblLH.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Function ShowIt(nText As String) As ToolTipEx
    Dim iMaxWidth As Long
    Dim iRect As RECT
    Dim iArea As Long
    Dim iCP As POINTAPI
    Dim iEP As DRAWTEXTPARAMS
    
    Set Me.Font = lblLH.Font
    
    iEP.cbSize = Len(iEP)
    iRect.Right = 200
    
    iMaxWidth = Screen.Width / 3 / Screen.TwipsPerPixelX
    DrawText Me.hDC, nText, -1, iRect, DT_CALCRECT Or DT_WORDBREAK
    Call DrawTextEx(Me.hDC, StrPtr(nText), Len(nText), iRect, DT_EXPANDTABS Or DT_TABSTOP Or DT_NOPREFIX Or DT_WORDBREAK Or DT_EDITCONTROL Or DT_CALCRECT, iEP)
    
    If iRect.Bottom > iRect.Right / 3 * 2 Then
        iArea = iRect.Right * iRect.Bottom
        iRect.Right = Sqr(iArea / (2 * 3)) * 3
        Call DrawTextEx(Me.hDC, StrPtr(nText), Len(nText), iRect, DT_EXPANDTABS Or DT_TABSTOP Or DT_NOPREFIX Or DT_WORDBREAK Or DT_EDITCONTROL Or DT_CALCRECT, iEP)
    End If
    iRect.Bottom = iRect.Bottom
    GetCursorPos iCP
    iCP.X = iCP.X - 15
    iCP.Y = iCP.Y + 30
    
    Me.Move iCP.X * Screen.TwipsPerPixelX, iCP.Y * Screen.TwipsPerPixelY, iRect.Right * Screen.TwipsPerPixelX + (Me.Width - Me.ScaleWidth) + 100, iRect.Bottom * Screen.TwipsPerPixelY + (Me.Height - Me.ScaleHeight) + 200
    
    If Me.Top + Me.Height > ScreenUsableHeight Then
        Me.Top = (iCP.Y - 50) * Screen.TwipsPerPixelY - Me.Height
    End If
    If Me.Top < 0 Then Me.Top = 0
    If Me.Left < 0 Then Me.Left = 0
    If (Me.Left + Me.Width) > Screen.Width Then
        Me.Left = Screen.Width - Me.Width
    End If
    lblLH.Caption = nText
    
    ShowNoActivate Me, , False, True
    Set mToolTipEx = New ToolTipEx
    Set ShowIt = mToolTipEx
End Function

Private Sub Form_Unload(Cancel As Integer)
    If Not mToolTipEx Is Nothing Then
        mToolTipEx.RaiseEventClosed
    End If
End Sub

Private Sub mToolTipEx_BeforeClose()
    mClosedByCode = True
    Unload Me
End Sub

