VERSION 5.00
Begin VB.UserControl ctlBuildHelp 
   BackColor       =   &H0000FFFF&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "ctlGetParentObject.ctx":0000
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "from property page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   144
      TabIndex        =   1
      Top             =   408
      Width           =   1452
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Load text files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   108
      TabIndex        =   0
      Top             =   120
      Width           =   1512
   End
End
Attribute VB_Name = "ctlBuildHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Resize()
    UserControl.Size ScaleX(lblText.Left * 2 + lblText.Width, UserControl.ScaleMode, vbTwips), ScaleY(lblText.Top * 2 + lblText.Height * 1.5, UserControl.ScaleMode, vbTwips)
End Sub

Public Function GetParent() As Object
    On Error Resume Next
    Set GetParent = UserControl.Parent
End Function
