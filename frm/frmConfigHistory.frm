VERSION 5.00
Begin VB.Form frmConfigHistory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "# History configuration"
   ClientHeight    =   2688
   ClientLeft      =   6108
   ClientTop       =   4320
   ClientWidth     =   4908
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigHistory.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2688
   ScaleWidth      =   4908
   ShowInTaskbar   =   0   'False
   Begin vbExtra.ButtonEx cmdEraseContext_2 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2070
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   614
      _ExtentY        =   656
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "# Close"
      Default         =   -1  'True
      Height          =   435
      Left            =   3000
      TabIndex        =   4
      Top             =   2100
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3690
      TabIndex        =   3
      Top             =   2910
      Width           =   1245
   End
   Begin VB.CheckBox chkRememberHistory 
      Caption         =   "# Remember the history across sessions"
      Height          =   225
      Left            =   750
      TabIndex        =   0
      Top             =   300
      Width           =   3485
   End
   Begin VB.CommandButton cmdEraseAll 
      Caption         =   "# Erase all"
      Height          =   435
      Left            =   330
      TabIndex        =   2
      Top             =   1320
      Width           =   4185
   End
   Begin VB.CommandButton cmdEraseContext 
      Caption         =   "# Erase history for this context"
      Height          =   435
      Left            =   330
      TabIndex        =   1
      Top             =   780
      Width           =   4185
   End
   Begin vbExtra.ButtonEx btnAyudaHistorial 
      Height          =   225
      Left            =   4275
      TabIndex        =   5
      Top             =   300
      Width           =   225
      _ExtentX        =   402
      _ExtentY        =   402
      ButtonStyle     =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16354918
      Caption         =   "?"
      ForeColor       =   16777215
      PictureAlign    =   4
   End
   Begin vbExtra.ButtonEx cmdEraseAll_2 
      Height          =   375
      Left            =   630
      TabIndex        =   7
      Top             =   2070
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   614
      _ExtentY        =   656
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      UseMaskCOlor    =   -1  'True
   End
   Begin vbExtra.ButtonEx cmdClose_2 
      Height          =   375
      Left            =   1020
      TabIndex        =   8
      Top             =   2070
      Visible         =   0   'False
      Width           =   345
      _ExtentX        =   614
      _ExtentY        =   656
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      UseMaskCOlor    =   -1  'True
   End
End
Attribute VB_Name = "frmConfigHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mContext As String
Private mHistoryErased As Boolean

Private Sub btnAyudaHistorial_Click()
    ShowToolTipEx GetLocalizedString(efnGUIStr_frmConfigHistory_HelpMessage), GetLocalizedString(efnGUIStr_frmConfigHistory_HelpMessageTitle), , True, vxTTIconInfo
End Sub

Private Sub cmdClose_2_Click()
    cmdClose_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEraseAll_2_Click()
    cmdEraseAll_Click
End Sub

Private Sub cmdEraseAll_Click()
    On Error Resume Next
    DeleteSetting AppNameForRegistry, "History"
    On Error GoTo 0
 '   mHistoryErased = True
    EraseAllHistories
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEraseContext_2_Click()
    cmdEraseContext_Click
End Sub

Private Sub cmdEraseContext_Click()
    Dim iCol
    Dim iSetting As Variant
    
    iCol = GetAllSettings(AppNameForRegistry, "History")
    For Each iSetting In iCol
        If Base64Decode(CStr(iSetting)) Like Context & "*" Then
            DeleteSetting AppNameForRegistry, "History", iSetting
        End If
    Next
    On Error Resume Next
    DeleteSetting AppNameForRegistry, "History", Base64Encode(Context)
    On Error GoTo 0
    mHistoryErased = True
End Sub

Private Sub Form_Load()
    Dim iLng As Long
    
    PersistForm Me, Forms
    iLng = Val(GetSetting(AppNameForRegistry, "History", "Record", "1"))
    If (iLng < 0) Or (iLng > 1) Then
        iLng = 1
    End If
    chkRememberHistory.Value = iLng
    LoadGUICaptions
    AssignAccelerators Me
    
    If gButtonsStyle <> -1 Then
        
        cmdEraseContext_2.Move cmdEraseContext.Left, cmdEraseContext.Top, cmdEraseContext.Width, cmdEraseContext.Height
        cmdEraseContext_2.Caption = cmdEraseContext.Caption
        cmdEraseContext.Visible = False
        cmdEraseContext_2.Visible = True
        cmdEraseContext_2.TabIndex = cmdEraseContext.TabIndex
        cmdEraseContext_2.ButtonStyle = gButtonsStyle
        
        cmdEraseAll_2.Move cmdEraseAll.Left, cmdEraseAll.Top, cmdEraseAll.Width, cmdEraseAll.Height
        cmdEraseAll_2.Caption = cmdEraseAll.Caption
        cmdEraseAll.Visible = False
        cmdEraseAll_2.Visible = True
        cmdEraseAll_2.TabIndex = cmdEraseAll.TabIndex
        cmdEraseAll_2.ButtonStyle = gButtonsStyle
        
        cmdClose_2.Move cmdClose.Left, cmdClose.Top, cmdClose.Width, cmdClose.Height
        cmdClose_2.Caption = cmdClose.Caption
        cmdClose.Visible = False
        cmdClose_2.Default = cmdClose.Default
        cmdClose_2.Cancel = cmdClose.Cancel
        cmdClose_2.Visible = True
        cmdClose_2.TabIndex = cmdClose.TabIndex
        cmdClose_2.ButtonStyle = gButtonsStyle
        
    End If
End Sub

Public Property Let Context(nValor As String)
    mContext = nValor
End Property

Private Property Get Context() As String
    If Trim$(mContext) <> "" Then
        Context = mContext
    Else
        Context = "General"
    End If
End Property

Public Property Get HistoryErased() As Boolean
    HistoryErased = mHistoryErased
End Property

Private Sub Form_Unload(Cancel As Integer)
    If Not (chkRememberHistory.Value = 1) Then
        On Error Resume Next
        DeleteSetting AppNameForRegistry, "History"
        On Error GoTo 0
'        mHistoryErased = True
    End If
    SaveSetting AppNameForRegistry, "History", "Record", chkRememberHistory.Value
End Sub

Private Sub LoadGUICaptions()
    Me.Caption = GetLocalizedString(efnGUIStr_frmConfigHistory_Caption)
    chkRememberHistory.Caption = GetLocalizedString(efnGUIStr_frmConfigHistory_chkRememberHistory_Caption)
    cmdEraseContext.Caption = GetLocalizedString(efnGUIStr_frmConfigHistory_cmdEraseContext_Caption)
    cmdEraseAll.Caption = GetLocalizedString(efnGUIStr_frmConfigHistory_cmdEraseAll_Caption)
    cmdClose.Caption = GetLocalizedString(efnGUIStr_General_CloseButton_Caption)
End Sub
