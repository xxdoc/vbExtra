VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.0#0"; "vbExtra2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   4872
   ClientLeft      =   7224
   ClientTop       =   1452
   ClientWidth     =   5712
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4872
   ScaleWidth      =   5712
   Begin vbExtra.MouseWheelEnabler MouseWheelEnabler1 
      Left            =   5220
      Top             =   4176
      _ExtentX        =   720
      _ExtentY        =   720
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   408
      Left            =   3852
      TabIndex        =   0
      Top             =   4320
      Width           =   1308
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3936
      Left            =   144
      TabIndex        =   1
      Top             =   144
      Width           =   5448
      _ExtentX        =   9610
      _ExtentY        =   6943
      _Version        =   393216
      Rows            =   100
      Cols            =   5
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label Label1 
      Caption         =   "Note: if you use a FlexFn control it also enables the scroll automatically (without needing a MouseWheelEnabler control)"
      ForeColor       =   &H00FF0000&
      Height          =   552
      Left            =   144
      TabIndex        =   2
      Top             =   4140
      Width           =   3684
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim c As Long
    Dim r As Long
    
    For c = 0 To MSHFlexGrid1.Cols - 1
        For r = 0 To MSHFlexGrid1.Rows - 1
            MSHFlexGrid1.TextMatrix(r, c) = Rnd
        Next r
    Next c
End Sub
