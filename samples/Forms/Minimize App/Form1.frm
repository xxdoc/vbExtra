VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4704
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   4836
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
   ScaleHeight     =   4704
   ScaleWidth      =   4836
   Begin VB.CommandButton Command1 
      Caption         =   "Show Form2"
      Height          =   624
      Left            =   1404
      TabIndex        =   0
      Top             =   1080
      Width           =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form2.Show 1
End Sub
