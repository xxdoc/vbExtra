VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.0#0"; "vbExtra2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3552
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   4536
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
   ScaleHeight     =   3552
   ScaleWidth      =   4536
   Begin VB.CommandButton Command1 
      Caption         =   "Change form's BackColor"
      Height          =   444
      Left            =   252
      TabIndex        =   0
      Top             =   2916
      Width           =   2352
   End
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   228
      Left            =   4308
      Top             =   3324
      Width           =   228
      _ExtentX        =   402
      _ExtentY        =   402
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Simply place a SizeGrip control on a sizable form"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1344
      Left            =   252
      TabIndex        =   1
      Top             =   252
      Width           =   4044
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim iDlg As New CommonDialogExObject
    
    iDlg.ShowColor
    If Not iDlg.Canceled Then
        Me.BackColor = iDlg.Color
    End If
End Sub
