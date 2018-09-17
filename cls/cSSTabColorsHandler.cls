VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSSTabColorsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Const WM_CTLCOLORSTATIC As Long = &H138

Private Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Private mSSTabHwnd As Long
Private mParentFormHwnd As Long
Private mBackColorBrush As Long
Private mHandlersCollection As cObjectHandlersCollection
Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1
Private mSSTab As Object

Public Sub SetSSTab(nSSTab As Object, Optional nHandlersCollection As Object)
    
    On Error Resume Next
    mSSTabHwnd = nSSTab.hWnd
    Set mSSTab = nSSTab
    On Error GoTo 0
    If mSSTabHwnd = 0 Then Exit Sub
    
    If LCase$(TypeName(nSSTab)) <> "sstab" Then Exit Sub
    
    If Not nHandlersCollection Is Nothing Then
        Set mHandlersCollection = nHandlersCollection
    End If
    
    AttachMessage Me, mSSTabHwnd, WM_PAINT
    AttachMessage Me, mSSTabHwnd, WM_CTLCOLORSTATIC
    AttachMessage Me, mSSTabHwnd, WM_DESTROY
    
    mParentFormHwnd = GetParentFormHwnd(mSSTabHwnd)
    If mParentFormHwnd <> 0 Then
        AttachMessage Me, mParentFormHwnd, WM_DESTROY
    End If
    
    mBackColorBrush = CreateSolidBrush(nSSTab.BackColor)
    
    RedrawWindow mSSTabHwnd, ByVal 0&, 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_ALLCHILDREN
    
    On Error Resume Next
    If TypeOf nSSTab.Parent Is Form Then
        Set mForm = nSSTab.Parent
    End If
End Sub

Private Sub Class_Terminate()
    Finish
End Sub

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Dim aRect       As RECT
    Dim updateRect  As RECT
    Dim destDC      As Long
    Dim tempDC      As Long
    Dim tempBmp     As Long
    Dim origDC      As Long
    Dim origBmp     As Long
    Dim maskDC      As Long
    Dim maskBmp     As Long
    Dim memDC       As Long
    Dim memBmp      As Long
    
    Dim wid         As Long
    Dim hgt         As Long
    
    Dim origBrush As Long
    Dim origColor As Long
    Static sLastBackColor As Long
    
    Select Case iMsg
        Case WM_PAINT
            GetUpdateRect hWnd, updateRect, False
            destDC = GetDC(hWnd)
            
            If Not mSSTab Is Nothing Then
                On Error Resume Next
                If CLng(mSSTab.BackColor) <> sLastBackColor Then
                    If mBackColorBrush <> 0 Then
                        DeleteObject mBackColorBrush
                    End If
                    mBackColorBrush = CreateSolidBrush(mSSTab.BackColor)
                    sLastBackColor = mSSTab.BackColor
                End If
                On Error GoTo 0
            End If
            
            ' get the SSTab's window dimensions
            GetWindowRect hWnd, aRect
            wid = aRect.Right - aRect.Left
            hgt = aRect.Bottom - aRect.Top
            
            ' create our other temporary device contexts.
            maskDC = CreateCompatibleDC(destDC)
            maskBmp = CreateBitmap(wid, hgt, 1, 1, ByVal 0&)
            memDC = CreateCompatibleDC(destDC)
            memBmp = CreateCompatibleBitmap(destDC, wid, hgt)
            tempDC = CreateCompatibleDC(destDC)
            tempBmp = CreateCompatibleBitmap(destDC, wid, hgt)
            origDC = CreateCompatibleDC(destDC)
            origBmp = CreateCompatibleBitmap(destDC, wid, hgt)
            
            ' delete the temporary 1x1 bitmap and put our (wid)x(hgt) ones in
            DeleteObject SelectObject(maskDC, maskBmp)
            DeleteObject SelectObject(memDC, memBmp)
            DeleteObject SelectObject(tempDC, tempBmp)
            DeleteObject SelectObject(origDC, origBmp)
            
            ' Call the control's original handler... paints the control on our back buffer
            ISubclass_WindowProc = CallOldWindowProc(hWnd, iMsg, origDC, lParam)
            bConsume = True
            'CallWindowProc oldSSTabProc, hWnd, wMsg, origDC, lParam
    
            ' This helps our mask to correctly calculate the B & w pixels of
            '  our mask. Only really works in Win98 and greater... and even then
            '  it is sometimes flakey... may need to loop thru x & y and use
            '  GetPixel/SetPixel to create mask if it is not generated properly.
            origColor = SetBkColor(destDC, GetSysColor(15))
            SetBkColor origDC, GetSysColor(15&)
            ' create a B&w pixel mask - background color is white, everything else
            '  is black.
            BitBlt maskDC, 0, 0, wid, hgt, origDC, 0, 0, vbSrcCopy
            
            ' select the pattern brush into the DC and pattern blit
            origBrush = SelectObject(tempDC, mBackColorBrush)
            PatBlt tempDC, 0, 0, wid, hgt, vbPatCopy
            SelectObject tempDC, origBrush
            
            ' clean up our original image of the control so only the non background
            '  color parts are showing... make everything else white.
            BitBlt memDC, 0, 0, wid, hgt, maskDC, 0, 0, vbSrcCopy
            BitBlt memDC, 0, 0, wid, hgt, origDC, 0, 0, vbSrcPaint
    
            'punch the hole for our control image
            BitBlt tempDC, 0, 0, wid, hgt, maskDC, 0, 0, vbMergePaint
            'put the control images back in
            BitBlt tempDC, 0, 0, wid, hgt, memDC, 0, 0, vbSrcAnd
            'copy our new version back to the control
            BitBlt destDC, 0, 0, wid, hgt, tempDC, 0, 0, vbSrcCopy
    
            ' clean up all of our used graphical resources (VERY IMPORTANT!!!)
            DeleteDC tempDC
            DeleteObject tempBmp
            DeleteDC maskDC
            DeleteObject maskBmp
            DeleteDC memDC
            DeleteObject memBmp
            DeleteDC origDC
            DeleteObject origBmp
            
            ' Replace the original background color
            SetBkColor destDC, origColor
            ' Release the SSTab's device context back to the system
            ReleaseDC hWnd, destDC
            
            ValidateRect hWnd, 0
                    
        Case WM_DESTROY
            Finish
        Case WM_CTLCOLORSTATIC
            SetBkMode wParam, 1      ' make the text draw transparent
            ISubclass_WindowProc = mBackColorBrush   ' return the background brush
        Case Else
    End Select
End Function

Private Sub Finish()
    If Not mHandlersCollection Is Nothing Then
        If mSSTabHwnd <> 0 Then
            mHandlersCollection.Remove mSSTabHwnd
        End If
        Set mHandlersCollection = Nothing
    End If
    
    If mSSTabHwnd <> 0 Then
        DetachMessage Me, mSSTabHwnd, WM_PAINT
        DetachMessage Me, mSSTabHwnd, WM_CTLCOLORSTATIC
        DetachMessage Me, mSSTabHwnd, WM_DESTROY
        mSSTabHwnd = 0
    End If
    If mParentFormHwnd <> 0 Then
        DetachMessage Me, mParentFormHwnd, WM_DESTROY
        mParentFormHwnd = 0
    End If
    If mBackColorBrush <> 0 Then
        DeleteObject mBackColorBrush
        mBackColorBrush = 0
    End If
    
    If Not mForm Is Nothing Then
        Set mForm = Nothing
    End If
    
    Set mSSTab = Nothing
End Sub

Private Sub mForm_Unload(Cancel As Integer)
    If Cancel = 0 Then
        Finish
    End If
End Sub
