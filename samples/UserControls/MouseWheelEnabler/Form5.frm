VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#1.6#0"; "vbExtra1.ocx"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form5"
   ClientHeight    =   4752
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   4956
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4752
   ScaleWidth      =   4956
   Begin vbExtra.MouseWheelEnabler MouseWheelEnabler1 
      Left            =   4500
      Top             =   4176
      _ExtentX        =   720
      _ExtentY        =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   408
      Left            =   108
      TabIndex        =   0
      Top             =   72
      Width           =   1848
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   408
      Left            =   3132
      TabIndex        =   4
      Top             =   4176
      Width           =   1308
   End
   Begin VB.TextBox Text1 
      Height          =   3396
      Left            =   108
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form5.frx":0000
      Top             =   612
      Width           =   4728
   End
   Begin VB.Label Label2 
      Caption         =   "It requires some code, an API call  (check the code)"
      ForeColor       =   &H00FF0000&
      Height          =   408
      Left            =   144
      TabIndex        =   3
      Top             =   4176
      Width           =   2856
   End
   Begin VB.Label Label1 
      Caption         =   "Look, Text1 can be scrolled even when the focus is on the CommandButton"
      ForeColor       =   &H00FF0000&
      Height          =   444
      Left            =   2052
      TabIndex        =   1
      Top             =   108
      Width           =   2748
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim c As Long
    
    For c = 1 To 100
        Text1.Text = Text1.Text & c & vbCrLf
    Next c
End Sub

Private Sub MouseWheelEnabler1_Message(ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean)
    SendMessageLong Text1.hWnd, iMsg, wParam, lParam
    Handled = True
End Sub
