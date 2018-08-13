VERSION 5.00
Begin VB.Form frmSelectColumns 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "# Configure visible columns"
   ClientHeight    =   4308
   ClientLeft      =   3996
   ClientTop       =   3060
   ClientWidth     =   5796
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectColumns.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4308
   ScaleWidth      =   5796
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "# Close"
      Height          =   435
      Left            =   3744
      TabIndex        =   3
      Top             =   3708
      Width           =   1515
   End
   Begin vbExtra.ButtonEx cmdClose_2 
      Default         =   -1  'True
      Height          =   432
      Left            =   1752
      TabIndex        =   0
      Top             =   3672
      Visible         =   0   'False
      Width           =   432
      _ExtentX        =   762
      _ExtentY        =   762
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
   Begin VB.ListBox lstColumns 
      Height          =   2556
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label lblTitle 
      Caption         =   "# Select the columns to display:"
      Height          =   315
      Left            =   270
      TabIndex        =   1
      Top             =   150
      Width           =   5235
   End
End
Attribute VB_Name = "frmSelectColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mColumnsHidden() As Boolean
Private mLoading As Boolean

Private Sub cmdClose_2_Click()
    cmdClose_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    PersistForm Me, Forms

    If gButtonsStyle <> -1 Then
        cmdClose_2.Move cmdClose.Left, cmdClose.Top, cmdClose.Width, cmdClose.Height
        cmdClose_2.Caption = cmdClose.Caption
        cmdClose.Visible = False
        cmdClose_2.Default = cmdClose.Default
        cmdClose_2.Cancel = cmdClose.Cancel
        cmdClose_2.Visible = True
        cmdClose_2.TabIndex = cmdClose.TabIndex
        cmdClose_2.ButtonStyle = gButtonsStyle
    End If
    LoadGUICaptions
    AssignAccelerators Me, True
End Sub

Private Sub lstColumns_ItemCheck(Item As Integer)
    If Not mLoading Then
        If Not EnsureOneVisible Then
            MsgBox GetLocalizedString(efnGUIStr_frmSelectColumns_OneVisible_Message), vbExclamation, ClientProductName
            lstColumns.Selected(Item) = True
        End If
    End If
    mColumnsHidden(lstColumns.ItemData(Item)) = Not lstColumns.Selected(Item)
End Sub

Private Function EnsureOneVisible() As Boolean
    Dim c As Long
    
    EnsureOneVisible = True
    For c = 0 To lstColumns.ListCount - 1
        If lstColumns.Selected(c) Then Exit Function
    Next c
    EnsureOneVisible = False
End Function

Public Sub SetData(nColHeaders() As String, nColsHiddenByClientProgram() As Boolean, nColsHiddenByComponent() As Boolean)
    Dim c As Long
    
    mLoading = True
    ReDim mColumnsHidden(UBound(nColHeaders))
    For c = 0 To UBound(nColHeaders)
        If Not nColsHiddenByClientProgram(c) Then
            lstColumns.AddItem nColHeaders(c)
            lstColumns.ItemData(lstColumns.NewIndex) = c
            If Not nColsHiddenByComponent(c) Then
                lstColumns.Selected(lstColumns.NewIndex) = True
            Else
                mColumnsHidden(c) = True
            End If
        End If
    Next c
    mLoading = False
End Sub

Public Property Get ColumnsHidden(Index As Long) As Boolean
    ColumnsHidden = mColumnsHidden(Index)
End Property

Private Sub LoadGUICaptions()
    Me.Caption = GetLocalizedString(efnGUIStr_frmSelectColumns_Caption)
    cmdClose.Caption = GetLocalizedString(efnGUIStr_General_CloseButton_Caption)
    lblTitle.Caption = GetLocalizedString(efnGUIStr_frmSelectColumns_lblTitle_Caption)
End Sub
