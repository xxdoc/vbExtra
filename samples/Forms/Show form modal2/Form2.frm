VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4020
   ClientLeft      =   4008
   ClientTop       =   2484
   ClientWidth     =   4368
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   4368
   Begin VB.CommandButton Command2 
      Caption         =   "Show Form4 non modally"
      Height          =   552
      Left            =   756
      TabIndex        =   1
      Top             =   1548
      Width           =   1704
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Form3 modally"
      Height          =   552
      Left            =   756
      TabIndex        =   0
      Top             =   468
      Width           =   1704
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ShowModal Form3
End Sub

Private Sub Command2_Click()
    Form4.Show
End Sub
