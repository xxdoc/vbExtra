VERSION 5.00
Begin VB.Form frmSettingGridDataProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1692
   ClientLeft      =   4656
   ClientTop       =   4080
   ClientWidth     =   6852
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettingGridDataProgress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1692
   ScaleWidth      =   6852
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "# Cancel"
      Height          =   435
      Left            =   4920
      TabIndex        =   0
      Top             =   1050
      Width           =   1515
   End
   Begin vbExtra.ButtonEx cmdCancel_2 
      Height          =   432
      Left            =   1920
      TabIndex        =   2
      Top             =   1044
      Width           =   432
      _ExtentX        =   402
      _ExtentY        =   402
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin vbExtra.ctlProgressBar pgb1 
      Height          =   195
      Left            =   420
      Top             =   540
      Width           =   6015
      _ExtentX        =   10605
      _ExtentY        =   339
   End
   Begin VB.Label lblMessage 
      Caption         =   "# Generating preview..."
      Height          =   255
      Left            =   420
      TabIndex        =   1
      Top             =   180
      Width           =   2895
   End
End
Attribute VB_Name = "frmSettingGridDataProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Canceled As Boolean

Private Sub cmdCancel_2_Click()
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    Canceled = True
End Sub

Private Sub Form_Load()
    CenterForm Me

    If gButtonsStyle <> -1 Then
        cmdCancel_2.Move cmdCancel.Left, cmdCancel.Top, cmdCancel.Width, cmdCancel.Height
        cmdCancel_2.Caption = cmdCancel.Caption
        cmdCancel.Visible = False
        cmdCancel_2.Default = cmdCancel.Default
        cmdCancel_2.Cancel = cmdCancel.Cancel
        cmdCancel_2.Visible = True
        cmdCancel_2.TabIndex = cmdCancel.TabIndex
        cmdCancel_2.ButtonStyle = gButtonsStyle
    End If
    
    lblMessage.Caption = GetLocalizedString(efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Start) & "..."
    cmdCancel.Caption = GetLocalizedString(efnGUIStr_General_CancelButton_Caption)
End Sub

Public Property Let CurrentPage(nValue As Long)
    lblMessage.Caption = GetLocalizedString(efnGUIStr_frmSettingGridDataProgress_lblMessage_Caption_Progress) & " " & nValue & "..."
End Property
