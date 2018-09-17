VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.0#0"; "vbExtra2.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4824
   ClientLeft      =   2268
   ClientTop       =   2016
   ClientWidth     =   8088
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
   ScaleHeight     =   4824
   ScaleWidth      =   8088
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   408
      Left            =   6048
      TabIndex        =   11
      Top             =   4068
      Width           =   1452
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   408
      Left            =   756
      TabIndex        =   9
      Top             =   4068
      Width           =   1452
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   408
      Left            =   2520
      TabIndex        =   8
      Top             =   4068
      Width           =   1452
   End
   Begin VB.TextBox txtLastName 
      Height          =   300
      Left            =   1476
      TabIndex        =   3
      Top             =   2592
      Width           =   2500
   End
   Begin vbExtra.DateEnter denBirthDate 
      Height          =   288
      Left            =   1476
      TabIndex        =   5
      Top             =   3024
      Width           =   2496
      _ExtentX        =   4403
      _ExtentY        =   508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "__/__/____"
   End
   Begin VB.TextBox txtFirstName 
      Height          =   300
      Left            =   1476
      TabIndex        =   1
      Top             =   2160
      Width           =   2500
   End
   Begin VB.ListBox List1 
      Height          =   1776
      Left            =   72
      TabIndex        =   10
      Top             =   108
      Width           =   7896
   End
   Begin vbExtra.DateEnter denJoinDate 
      Height          =   288
      Left            =   1476
      TabIndex        =   7
      Top             =   3456
      Width           =   2496
      _ExtentX        =   4403
      _ExtentY        =   508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "__/__/____"
   End
   Begin VB.Label Label6 
      Caption         =   "It is a combination of a Masked Edit control and a DTPicker control (but without the dependency on these libraries)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   732
      Left            =   4176
      TabIndex        =   13
      Top             =   2700
      Width           =   3504
   End
   Begin VB.Label Label5 
      Caption         =   "The purpose of the DateEnter control is to speed up the data entry of dates"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   516
      Left            =   4176
      TabIndex        =   12
      Top             =   2160
      Width           =   3504
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Join date:"
      Height          =   228
      Left            =   108
      TabIndex        =   6
      Top             =   3456
      Width           =   1236
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "First name:"
      Height          =   228
      Left            =   324
      TabIndex        =   2
      Top             =   2628
      Width           =   1020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Date of birth:"
      Height          =   228
      Left            =   108
      TabIndex        =   4
      Top             =   3024
      Width           =   1236
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "First name:"
      Height          =   228
      Left            =   324
      TabIndex        =   0
      Top             =   2196
      Width           =   1020
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If Not ValidateData Then Exit Sub
    List1.AddItem UCase(Left(txtFirstName.Text, 1)) & Mid(txtFirstName.Text, 2) & " " & UCase(Left(txtLastName.Text, 1)) & Mid(txtLastName.Text, 2) & vbTab & "Birth date: " & FormatDateTime(denBirthDate.Value, vbShortDate) & vbTab & "Join date: " & FormatDateTime(denJoinDate.Value, vbShortDate)
    cmdCancel_Click
End Sub

Private Sub cmdCancel_Click()
    txtFirstName.Text = ""
    txtLastName.Text = ""
    denBirthDate.Value = Null
    denJoinDate.Value = Null
    txtFirstName.SetFocus
End Sub

Private Function ValidateData() As Boolean
    If Trim$(txtFirstName.Text) = "" Then
        MsgBox "Please enter the first name", vbExclamation
        txtFirstName.SetFocus
        Exit Function
    End If
    If Trim$(txtLastName.Text) = "" Then
        MsgBox "Please enter the last name", vbExclamation
        txtLastName.SetFocus
        Exit Function
    End If
    If IsNull(denBirthDate.Value) Then
        MsgBox "Please enter the birth date", vbExclamation
        denBirthDate.SetFocus
        Exit Function
    End If
    If IsNull(denJoinDate.Value) Then
        MsgBox "Please enter the join date", vbExclamation
        denJoinDate.SetFocus
        Exit Function
    End If
    ValidateData = True
End Function

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub denBirthDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        denJoinDate.SetFocus
    End If
End Sub

Private Sub denJoinDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        denJoinDate.Validate
        cmdAdd_Click
    End If
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        txtLastName.SetFocus
    End If
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        denBirthDate.SetFocus
    End If
End Sub
