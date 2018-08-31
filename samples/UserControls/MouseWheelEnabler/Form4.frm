VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#1.6#0"; "vbExtra1.ocx"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form4"
   ClientHeight    =   4764
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   4932
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4764
   ScaleWidth      =   4932
   Begin vbExtra.MouseWheelEnabler MouseWheelEnabler1 
      Left            =   3276
      Top             =   648
      _ExtentX        =   720
      _ExtentY        =   720
   End
   Begin VB.TextBox Text1 
      Height          =   552
      Left            =   432
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   576
      Width           =   1812
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4296
      LargeChange     =   5
      Left            =   4392
      Max             =   40
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   216
      Width           =   336
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub VScroll1_Scroll()
    Text1.Text = VScroll1.Value
End Sub
