VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRemoveBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Const RGN_OR = 2
Private Const SM_CXEDGE = 45
Private Const SM_CYEDGE = 46
Private Const SM_CXVSCROLL = 2
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6

Private mRegionCombo As Long
Private mHwndControl As Long

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mLabelName As String
Private mBorderLabel As Label

Private mBorderWidth As Long
Private mBorderHeight As Long
Private mComboArrowWidth As Long
Private mFlatBorderWidth As Long
Private mFlatBorderHeight As Long
Private mVerticalScrollbarWidth As Long

Public Sub SetControl(nControl As Control, Optional nBorder As Boolean)
    Dim iRect1 As RECT
    Dim iRect2 As RECT
    Dim iTmpRgn As Long
    Dim iRgnOld As Long
    
    mHwndControl = nControl.hWnd
    
    iRgnOld = mRegionCombo
    
    GetWindowRect mHwndControl, iRect1
    
    iRect1.Bottom = iRect1.Bottom - iRect1.Top
    iRect1.Top = 0
    iRect1.Right = iRect1.Right - iRect1.Left
    iRect1.Left = 0
    
    InflateRect iRect1, -mBorderWidth, -mBorderHeight
    CopyRect iRect2, iRect1
    
    iRect1.Left = iRect1.Right - mComboArrowWidth
    iRect2.Right = iRect1.Right - mComboArrowWidth
    InflateRect iRect2, -mFlatBorderWidth, -mFlatBorderHeight
    
    mRegionCombo = CreateRectRgnIndirect(iRect1)
    iTmpRgn = CreateRectRgnIndirect(iRect2)
    CombineRgn mRegionCombo, mRegionCombo, iTmpRgn, RGN_OR
    DeleteObject iTmpRgn
    SetWindowRgn mHwndControl, mRegionCombo, True
    
    If iRgnOld <> 0 Then
        DeleteObject iRgnOld
    End If
    
    If nBorder Then
        Set mForm = GetParentForm(nControl)
        mLabelName = "lblX_" & nControl.Name
        Set mBorderLabel = mForm.Controls.Add("VB.Label", mLabelName)
        On Error Resume Next
        Set mBorderLabel.Container = nControl.Container
        mBorderLabel.BackColor = vbRed ' nControl.BackColor
        On Error GoTo 0
        mBorderLabel.Appearance = 0 ' flat
        mBorderLabel.BorderStyle = 1
        mBorderLabel.Move nControl.Left, nControl.Top, nControl.Width - (2 + mVerticalScrollbarWidth) * Screen.TwipsPerPixelX, nControl.Height
        mBorderLabel.Visible = True
    End If
End Sub

Private Sub Class_Initialize()
    mBorderWidth = GetSystemMetrics(SM_CXEDGE)
    mBorderHeight = GetSystemMetrics(SM_CYEDGE)
    mComboArrowWidth = GetSystemMetrics(SM_CXVSCROLL)
    mFlatBorderWidth = GetSystemMetrics(SM_CXBORDER)
    mFlatBorderHeight = GetSystemMetrics(SM_CYBORDER)
    mVerticalScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL)
End Sub

Private Sub Class_Terminate()
    DeleteRegion
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        DeleteRegion
        If mLabelName <> "" Then
            On Error Resume Next
            mForm.Controls.Remove mLabelName
            mLabelName = ""
        End If
    End If
    Set mForm = Nothing
End Sub

Private Sub DeleteRegion()
    If mRegionCombo <> 0 Then
        If mHwndControl <> 0 Then
            If IsWindowVisible(mHwndControl) Then
                SetWindowRgn mHwndControl, 0&, True
            End If
        End If
        DeleteObject mRegionCombo
        mRegionCombo = 0
    End If
End Sub

Private Function GetParentForm(nControl As Control) As Object
    Set GetParentForm = nControl.Parent
    Do Until TypeOf GetParentForm Is Form
        Set GetParentForm = GetParentForm.Parent
    Loop
End Function

