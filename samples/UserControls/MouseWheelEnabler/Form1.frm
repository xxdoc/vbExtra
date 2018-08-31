VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3816
   ClientLeft      =   1548
   ClientTop       =   1440
   ClientWidth     =   4908
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3816
   ScaleWidth      =   4908
   Begin VB.CommandButton Command4 
      Caption         =   "Sample 4: scroll other controls"
      Height          =   408
      Left            =   612
      TabIndex        =   5
      Top             =   2484
      Width           =   3612
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sample 3: scroll a VScrollBar"
      Height          =   408
      Left            =   612
      TabIndex        =   4
      Top             =   1944
      Width           =   3612
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sample 2: scroll a flexgrid"
      Height          =   408
      Left            =   612
      TabIndex        =   3
      Top             =   1404
      Width           =   3612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample 1: scroll a RichTextBox without focus"
      Height          =   408
      Left            =   612
      TabIndex        =   2
      Top             =   864
      Width           =   3612
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   408
      Left            =   3204
      TabIndex        =   0
      Top             =   3240
      Width           =   1308
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The MouseWheelEnabler control allows to scroll controls with the mouse wheel"
      ForeColor       =   &H00FF0000&
      Height          =   552
      Left            =   612
      TabIndex        =   1
      Top             =   216
      Width           =   3576
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Command3_Click()
    Form4.Show
End Sub

Private Sub Command4_Click()
    Form5.Show
End Sub
