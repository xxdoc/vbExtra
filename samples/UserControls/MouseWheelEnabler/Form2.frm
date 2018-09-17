VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.0#0"; "vbExtra2.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6168
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   5760
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
   ScaleHeight     =   6168
   ScaleWidth      =   5760
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   408
      Left            =   3744
      TabIndex        =   0
      Top             =   5616
      Width           =   1308
   End
   Begin VB.TextBox Text2 
      Height          =   1272
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form2.frx":0000
      Top             =   3564
      Width           =   5412
   End
   Begin VB.TextBox Text1 
      Height          =   408
      Left            =   2844
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   252
      Width           =   1920
   End
   Begin vbExtra.MouseWheelEnabler MouseWheelEnabler1 
      Left            =   5040
      Top             =   360
      _ExtentX        =   720
      _ExtentY        =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   408
      Left            =   432
      TabIndex        =   1
      Top             =   252
      Width           =   2208
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1920
      Left            =   180
      TabIndex        =   3
      Top             =   900
      Width           =   5448
      _ExtentX        =   9610
      _ExtentY        =   3387
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form2.frx":0006
   End
   Begin VB.Label Label2 
      Caption         =   "Only when the focus is on other scrollable control it doesn't scroll the RichTextBox (click on Text2 to set the focus to it)"
      ForeColor       =   &H00FF0000&
      Height          =   552
      Left            =   180
      TabIndex        =   6
      Top             =   4968
      Width           =   5412
   End
   Begin VB.Label Label1 
      Caption         =   "Look, the focus is in the CommandButton and the mouse wheel still scrolls the RichTextBox"
      ForeColor       =   &H00FF0000&
      Height          =   552
      Left            =   180
      TabIndex        =   4
      Top             =   2952
      Width           =   5412
   End
End
Attribute VB_Name = "Form2"
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
    
    For c = 1 To 100
        RichTextBox1.Text = RichTextBox1.Text & vbCrLf & c
        Text2.Text = Text2.Text & vbCrLf & c
    Next c
End Sub



