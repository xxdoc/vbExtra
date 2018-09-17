VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.0#0"; "vbExtra2.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3744
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   4872
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
   ScaleHeight     =   3744
   ScaleWidth      =   4872
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   372
      Left            =   2952
      TabIndex        =   2
      Top             =   3204
      Width           =   1308
   End
   Begin VB.TextBox Text1 
      Height          =   2460
      Left            =   108
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   576
      Width           =   4620
   End
   Begin vbExtra.FontSizeChanger FontSizeChanger1 
      Height          =   372
      Left            =   324
      TabIndex        =   0
      Top             =   108
      Width           =   996
      _ExtentX        =   1757
      _ExtentY        =   656
      BoundControlName=   "Text1"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub
