VERSION 5.00
Begin VB.UserControl ButtonEx 
   AutoRedraw      =   -1  'True
   ClientHeight    =   492
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1332
   DefaultCancel   =   -1  'True
   LockControls    =   -1  'True
   PropertyPages   =   "ctlButtonEx.ctx":0000
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   111
   ToolboxBitmap   =   "ctlButtonEx.ctx":0035
End
Attribute VB_Name = "ButtonEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISubclass

'***************************************************************************
'*  Title:      JC button
'*  Function:   An ownerdrawn multistyle button
'*  Author:     Juned Chhipa
'*  Created:    November 2008
'*  Contact me: juned.chhipa@yahoo.com
'*
'*  Copyright © 2008-2009 Juned Chhipa. All rights reserved.
'***************************************************************************
'* This control can be used as an alternative to Command Button. It is
'* a lightweight button control which will emulate new command buttons.
'* Compile to get more faster results
'*
'* This control uses self-subclassing routines of Paul Caton. [Not any more]
'* Feel free to use this control. Please read Licence.txt
'* Please send comments/suggestions/bug reports to juned.chhipa@yahoo.com
'****************************************************************************
'*
'* - CREDITS:
'* - Paul Caton  :-  Self-Subclass Routines
'* - Noel Dacara :-  For helping me (Also, his dcbutton helped me a lot)
'* - LaVolpe     :-  Great Inspirations
'* - Jim Jose    :-  To make grayscale (disabled) bitmap/icon
'* - Carles P.V. :-  For fastest gradient routines
'*   If any bugs found, please report  :- juned.chhipa@yahoo.com
'*
'* I have tested this control many times and I have tried my best to make
'* it work as a real command button. But still, I cannot guarantee that
'* that this is FREE OF BUGS. So please let me know if u find any.

'****************************************************************************
'* This software is provided "as-is" without any express/implied warranty.  *
'* In no event shall the author be held liable for any damages arising      *
'* from the use of this software.                                           *
'* If you do not agree with these terms, do not install "JCButton". Use     *
'* of the program implicitly means you have agreed to these terms.          *        *
'                                                                           *
'* Permission is granted to anyone to use this software for any purpose,    *
'* including commercial use, and to alter and redistribute it, provided     *
'* that the following conditions are met:                                   *
'*                                                                          *
'* 1.All redistributions of source code files must retain all copyright     *
'*   notices that are currently in place, and this list of conditions       *
'*   without any modification.                                              *
'*                                                                          *
'* 2.All redistributions in binary form must retain all occurrences of      *
'*   above copyright notice and web site addresses that are currently in    *
'*   place (for example, in the About boxes).                               *
'*                                                                          *
'* 3.Modified versions in source or binary form must be plainly marked as   *
'*   such, and must not be misrepresented as being the original software.   *                         NOTE: this is a modified version
'****************************************************************************

Public Enum vbExButtonStyleConstants
    vxStandard                 'Standard VB
    vxFlat                          'Standard Toolbar
    vxWindowsXP            'Win XP
    vxXPToolbar               'XP Toolbar
    vxVistaAero                'Vista Aero
    vxAOL                         'AOL
    vxInstallShield             'InstallShield?!?~?
    vxOutlook2007           'Office 2007 Outlook
    vxVistaToolbar           'Vista Toolbar
    vxVisualStudio            'Visual Studio 2005
    vxGelButton                'Gel
    vx3DHover                 '3D Hover
    vxFlatHover                'Flat Hover
    vxVector
    vxPlastic                     'Inspired from Candy Button (but drawn in a different style)
    vxInstallShieldToolbar
    vxInstallShieldToolBar2
    vxInstallShield2
    vxVistaAero2
End Enum

Public Enum vbExButtonCaptionAlignConstants
    vxLeftAlign
    vxCenterAlign
    vxRightAlign
End Enum

Public Enum vbExButtonPictureAlignConstants
    vxLeftEdge
    vxLeftOfCaption
    vxRightEdge
    vxRightOfCaption
    vxCenter
    vxTopEdge
    vxTopOfCaption
    vxBottomEdge
    vxBottomOfCaption
End Enum


Private Declare Function ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, pwHue As Long, pwLuminance As Long, pwSaturation As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Point) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ByRef pccolorref As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

'User32 Declares
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long

Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Private bTrack              As Boolean
Private bTrackUser32        As Boolean

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1
    TME_LEAVE = &H2
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOUSELEAVE             As Long = &H2A3
'Private Const WM_MOVING                 As Long = &H216
Private Const WM_NCACTIVATE             As Long = &H86
Private Const WM_ACTIVATE               As Long = &H6

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                As Long
    dwFlags                               As TRACKMOUSEEVENT_FLAGS
    hwndTrack                             As Long
    dwHoverTime                           As Long
End Type

'Kernel32 declares used by the Subclasser
Private Declare Function GetModuleHandleA Lib "Kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "Kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'  End of Subclassing Declares
'==========================================================================================================================================================================================================================================================================================================

Private Enum enumButtonStates
    [eStateNormal]              'Normal State
    [eStateOver]                'Hover State
    [eStateDown]                'Down State
End Enum

Private Enum enumGradientDirectionCts
    [gdHorizontal] = 0
    [gdVertical] = 1
    [gdDownwardDiagonal] = 2
    [gdUpwardDiagonal] = 3
End Enum

'  used for Button colors
Private Type tButtonColors
    tBackColor      As Long
    tDisabledColor  As Long
    tForeColor      As Long
    tGreyText       As Long
End Type

'  used to define various graphics areas
Private Type RECT
    Left As Long
    Top     As Long
    Right As Long
    Bottom  As Long
End Type

Private Type Point
    x       As Long
    y       As Long
End Type

'  RGB Colors structure
'Private Type RGBColor
'    R       As Single
'    G       As Single
'    B       As Single
'End Type

'  for gradient painting and bitmap tiling
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type ICONINFO
    fIcon       As Long
    xHotspot    As Long
    yHotspot    As Long
    hbmMask     As Long
    hbmColor    As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformID        As Long
    szCSDVersion        As String * 128 '* Maintenance string for PSS usage.
End Type

' --constants for unicode support
Private Const VER_PLATFORM_WIN32_NT = 2

' --constants for  Flat Button
Private Const BDR_RAISEDINNER   As Long = &H4

' --constants for Standard VB button
Private Const BDR_SUNKEN95 As Long = &HA
Private Const BDR_RAISED95 As Long = &H5

Private Const BF_Left As Long = &H1
Private Const BF_TOP        As Long = &H2
Private Const BF_Right As Long = &H4
Private Const BF_BOTTOM     As Long = &H8
Private Const BF_RECT       As Long = (BF_Left Or BF_TOP Or BF_Right Or BF_BOTTOM)

' --System Hand Pointer
'Private Const IDC_HAND As Long = 32649

' --Color Constant
Private Const CLR_INVALID       As Long = &HFFFF
Private Const DIB_RGB_COLORS    As Long = 0

' --Formatting Text Consts
'Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_NOPREFIX As Long = &H800&
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_RIGHT As Long = &H2
Private Const DT_CALCRECT As Long = &H400
' --for drawing Icon Constants
Private Const DI_NORMAL As Long = &H3

' --Property Variables:
Private m_Picture           As StdPicture           'Icon of button
Private m_Pic16             As StdPicture
Private m_Pic24             As StdPicture
Private m_Pic20             As StdPicture
Private m_PicToUse      As StdPicture
Private m_DisabledPicture  As StdPicture

Private m_ButtonStyle       As vbExButtonStyleConstants     'Choose your Style
Private m_Buttonstate       As enumButtonStates     'Normal / Over / Down
Private m_bIsDown           As Boolean              'Is button is pressed?
Private m_bMouseInCtl       As Boolean              'Is Mouse in Control
Private m_bHasFocus         As Boolean              'Has focus?
'Private m_bHandPointer      As Boolean              'Use Hand Pointer
Private m_bDefault          As Boolean              'Is Default?
Private m_bCheckBoxMode     As Boolean              'Is checkbox?
Private m_bValue            As Boolean              'Value (Checked/Unchekhed)
Private m_bShowFocus        As Boolean              'Bool to show focus
Private m_bParentActive     As Boolean              'Parent form Active or not
Private m_lParenthWnd       As Long                 'Is parent active?
Private m_WindowsNT         As Long                 'OS Supports Unicode?
Private m_bEnabled          As Boolean              'Enabled/Disabled
Private m_Caption           As String               'String to draw caption
'Private m_TextRect          As RECT                 'Text Position
Private m_CapRect           As RECT                 'For InstallShield style
Private m_CaptionAlign      As vbExButtonCaptionAlignConstants
Private m_PictureAlign      As vbExButtonPictureAlignConstants     'Picture Alignments
Private m_bColors           As tButtonColors        'Button Colors
Private m_bUseMaskColor     As Boolean              'Transparent areas
Private m_bUseMnemonic     As Boolean              'Transparent areas
Private m_lMaskColor        As Long                 'Set Transparent color
Private m_lButtonRgn        As Long                 'Button Region
Private m_bIsSpaceBarDown   As Boolean              'Space bar down boolean
Private m_lDownButton       As Integer              'For click/Dblclick events
Private m_lDShift           As Integer              'A flag for dblClick
Private m_lDX               As Single
Private m_lDY               As Single
Private m_ButtonRect        As RECT                 'Button Position
'Private m_FocusRect         As RECT
Private lh                  As Long                 'ScaleHeight of button
Private lw                  As Long                 'ScaleWidth of button
'Private XPos                As Long                 'X position of picture
'Private YPos                As Long                 'Y Position of Picture
Private mUserControlHwnd As Long
Private m_BackColor As Long
Private m_ForeColor As Long
Private m_BackColorBkg As Long
Private mBackColorR As Long
Private mBackColorG As Long
Private mBackColorB As Long
Private m_BlendDisabledPicWithBackColor As Boolean
Private mRedraw As Boolean
Private mRedrawPending As Boolean
Private mSetPicToUsePending As Boolean

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1

'  Events
Public Event Click()
Public Event DblClick()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)

'  PRIVATE ROUTINES

Private Function PaintGrayScale(ByVal lHDC As Long, ByVal hPicture As Long, ByVal lLeft As Long, ByVal lTop As Long, Optional ByVal lWidth As Long = -1, Optional ByVal lHeight As Long = -1) As Boolean

    '****************************************************************************
    '*  Converts an icon/bitmap to grayscale (used for Disabled buttons)        *
    '*  Author:  Jim Jose                                                       *
    '*  Modified by me for Disabled Bitmaps (for Maskcolor)
    '*  All Credits goes to Jim Jose                                            *
    '****************************************************************************

    Dim BMP        As BITMAP
    Dim BMPiH      As BITMAPINFOHEADER
    Dim lBits()    As Byte 'Packed DIB
    Dim lTrans()   As Byte 'Packed DIB
    Dim TmpDC      As Long
    Dim x          As Long
    Dim xMax       As Long
    Dim TmpCol     As Long
    Dim R1         As Long
    Dim G1         As Long
    Dim B1         As Long
    Dim bIsIcon    As Boolean
    
    'Dim hDCSrc   As Long
    'Dim hOldob   As Long
    'Dim PicSize  As Long
    Dim oPic     As New StdPicture

    Set oPic = m_PicToUse

    '  Get the Image format
    If (GetObjectType(hPicture) = 0) Then
        Dim mIcon As ICONINFO
        bIsIcon = True
        GetIconInfo hPicture, mIcon
        hPicture = mIcon.hbmColor
    End If
    
    '  Get image info
    GetObject hPicture, Len(BMP), BMP

    '  Prepare DIB header and redim. lBits() array
    With BMPiH
        .biSize = Len(BMPiH) '40
        .biPlanes = 1
        .biBitCount = 24
        .biWidth = BMP.bmWidth
        .biHeight = BMP.bmHeight
        .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        'If lWidth = -1 Then lWidth = .biWidth
        If lWidth = -1 Then
            lWidth = .biWidth
        End If
        'If lHeight = -1 Then lHeight = .biHeight
        If lHeight = -1 Then
            lHeight = .biHeight
        End If
    End With

    ReDim lBits(Len(BMPiH) + BMPiH.biSizeImage)   '[Header + Bits]

    '  Create TemDC and Get the image bits
    TmpDC = CreateCompatibleDC(lHDC)
    GetDIBits TmpDC, hPicture, 0, BMP.bmHeight, lBits(0), BMPiH, DIB_RGB_COLORS

    '  Loop through the array... (grayscale - average!!)
    xMax = BMPiH.biSizeImage - 1
    For x = 0 To xMax - 3 Step 3
        R1 = lBits(x)
        G1 = lBits(x + 1)
        B1 = lBits(x + 2)
        TmpCol = (R1 + G1 + B1) \ 3
        lBits(x) = TmpCol
        lBits(x + 1) = TmpCol
        lBits(x + 2) = TmpCol
    Next x

    '  Paint it!
    If bIsIcon Then
        ReDim lTrans(Len(BMPiH) + BMPiH.biSizeImage)
        GetDIBits TmpDC, mIcon.hbmMask, 0, BMP.bmHeight, lTrans(0), BMPiH, DIB_RGB_COLORS   ' Get the mask
        StretchDIBits lHDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lTrans(0), BMPiH, DIB_RGB_COLORS, vbSrcAnd     ' Draw the mask
        PaintGrayScale = StretchDIBits(lHDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, DIB_RGB_COLORS, vbSrcPaint)   'Draw the gray
        DeleteObject mIcon.hbmMask  'Delete the extracted images
        DeleteObject mIcon.hbmColor
    Else
        ReDim lTrans(Len(BMPiH) + BMPiH.biSizeImage)
        GetDIBits TmpDC, mIcon.hbmMask, 0, BMP.bmHeight, lTrans(0), BMPiH, DIB_RGB_COLORS  ' Get the mask
        StretchDIBits lHDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lTrans(0), BMPiH, DIB_RGB_COLORS, vbSrcAnd    ' Draw the mask
        PaintGrayScale = StretchDIBits(lHDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, DIB_RGB_COLORS, vbSrcPaint)
        DeleteObject mIcon.hbmMask  'Delete the extracted images
        DeleteObject mIcon.hbmColor
    End If

    '   Clear memory
    DeleteDC TmpDC

End Function

Private Sub DrawLineApi(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

    '****************************************************************************
    '*  draw lines
    '****************************************************************************

    Dim pt      As Point
    Dim hPen    As Long
    Dim hPenOld As Long

    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(UserControl.hDC, hPen)
    MoveToEx UserControl.hDC, X1, Y1, pt
    LineTo UserControl.hDC, X2, Y2
    SelectObject UserControl.hDC, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld

End Sub

Private Function BlendColors(ByVal lBackColorFrom As Long, ByVal lBackColorTo As Long) As Long

    '***************************************************************************
    '*  Combines (mix) two colors                                              *
    '***************************************************************************

    BlendColors = RGB(((lBackColorFrom And &HFF) + (lBackColorTo And &HFF)) / 2, (((lBackColorFrom \ &H100) And &HFF) + ((lBackColorTo \ &H100) And &HFF)) / 2, (((lBackColorFrom \ &H10000) And &HFF) + ((lBackColorTo \ &H10000) And &HFF)) / 2)

End Function

Private Sub DrawRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long)

    '****************************************************************************
    '*  Draws a rectangle specified by coords and color of the rectangle        *
    '****************************************************************************

    Dim brect As RECT
    Dim hBrush As Long
    Dim ret As Long

    brect.Left = x
    brect.Top = y
    brect.Right = x + Width
    brect.Bottom = y + Height

    hBrush = CreateSolidBrush(Color)

    ret = FrameRect(hDC, brect, hBrush)

    ret = DeleteObject(hBrush)

End Sub

'Private Sub DrawFocusRectangle(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long)
'
''****************************************************************************
''*  Draws a Focus Rectangle inside button if m_bShowFocus property is True  *
''****************************************************************************
'
'Dim brect As RECT
'Dim RetVal As Long
'
'    brect.Left = x
'    brect.Top = y
'    brect.Right = x + Width
'    brect.Bottom = y + Height
'
'    RetVal = DrawFocusRect(hDc, brect)
'
'End Sub

Private Sub DrawGradientEx(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal GradientDirection As enumGradientDirectionCts)

    '****************************************************************************
    '* Draws very fast Gradient in four direction.                              *
    '* Author: Carles P.V (Gradient Master)                                     *
    '* This routine works as a heart for this control.                          *
    '* Thank you so much Carles.                                                *
    '****************************************************************************

    Dim uBIH    As BITMAPINFOHEADER
    Dim lBits() As Long
    Dim lGrad() As Long

    Dim R1      As Long
    Dim G1      As Long
    Dim B1      As Long
    Dim R2      As Long
    Dim G2      As Long
    Dim B2      As Long
    Dim dR      As Long
    Dim dG      As Long
    Dim dB      As Long

    Dim Scan    As Long
    Dim i       As Long
    Dim iEnd    As Long
    Dim iOffset As Long
    Dim J       As Long
    Dim jEnd    As Long
    Dim iGrad   As Long

    '-- A minor check

    'If (Width < 1 Or Height < 1) Then Exit Sub
    If (Width < 1 Or Height < 1) Then
        Exit Sub
    End If

    '-- Decompose colors
    Color1 = Color1 And &HFFFFFF
    R1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    G1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    B1 = Color1 Mod &H100&
    Color2 = Color2 And &HFFFFFF
    R2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    G2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    B2 = Color2 Mod &H100&

    '-- Get color distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1

    '-- Size gradient-colors array
    Select Case GradientDirection
        Case [gdHorizontal]
            ReDim lGrad(0 To Width - 1)
        Case [gdVertical]
            ReDim lGrad(0 To Height - 1)
        Case Else
            ReDim lGrad(0 To Width + Height - 2)
    End Select

    '-- Calculate gradient-colors
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
    Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If

    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width

    '-- Render gradient DIB
    Select Case GradientDirection

        Case [gdHorizontal]

            For J = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next J

        Case [gdVertical]

            For J = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(J)
                Next i
                iOffset = iOffset + Scan
            Next J

        Case [gdDownwardDiagonal]

            iOffset = jEnd * Scan
            For J = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = J
            Next J

        Case [gdUpwardDiagonal]

            iOffset = 0
            For J = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = J
            Next J
    End Select

    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With

    '-- Paint it!
    StretchDIBits UserControl.hDC, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy

End Sub

Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional ByRef hPalette As Long = 0) As Long

    '****************************************************************************
    '*  System color code to long rgb                                           *
    '****************************************************************************

    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function

Private Sub RedrawButton(Optional nFromMouseMove As Boolean)

    '****************************************************************************
    '*  The main routine of this usercontrol. Everything is drawn here.         *
    '****************************************************************************
    Static sLastFrmMM As Boolean

    If nFromMouseMove And sLastFrmMM Then Exit Sub
    sLastFrmMM = nFromMouseMove

    If Not Redraw Then
        mRedrawPending = True
        Exit Sub
    End If
    mRedrawPending = False

    UserControl.Cls                                'Clears usercontrol
    lh = ScaleHeight
    lw = ScaleWidth

    m_bColors.tBackColor = TranslateColor(m_BackColor)
    m_bColors.tForeColor = TranslateColor(m_ForeColor)

    SetRect m_ButtonRect, 0, 0, lw, lh             'Sets the button rectangle

    If (m_bCheckBoxMode) Then                      'If Checkboxmode True
        If Not (m_ButtonStyle = vxStandard Or m_ButtonStyle = vxXPToolbar Or m_ButtonStyle = vxVisualStudio) Then
            If m_bValue Then m_Buttonstate = eStateDown
        End If
    End If

    Select Case m_ButtonStyle

        Case vxStandard
            DrawStandardButton m_Buttonstate
        Case vx3DHover
            DrawStandardButton m_Buttonstate
        Case vxFlat
            DrawStandardButton m_Buttonstate
        Case vxFlatHover
            DrawStandardButton m_Buttonstate
        Case vxWindowsXP
            DrawWinXPButton m_Buttonstate
        Case vxXPToolbar
            DrawXPToolbar m_Buttonstate
        Case vxGelButton
            DrawGelButton m_Buttonstate
        Case vxAOL
            DrawAOLButton m_Buttonstate
        Case vxInstallShield
            DrawInstallShieldButton m_Buttonstate
        Case vxInstallShieldToolbar
            DrawInstallShieldToolBarDAButton m_Buttonstate
        Case vxInstallShieldToolBar2
            DrawInstallShieldToolbar2Button m_Buttonstate
        Case vxInstallShield2
            DrawInstallShieldReverseButton m_Buttonstate
        Case vxVistaAero
            DrawVistaButton m_Buttonstate
        Case vxVistaToolbar
            DrawVistaToolbarStyle m_Buttonstate
        Case vxVistaAero2
            DrawVistaAero2Style m_Buttonstate
        Case vxVisualStudio
            DrawVisualStudio2005 m_Buttonstate
        Case vxOutlook2007
            DrawOutlook2007 m_Buttonstate
        Case vxVector
            DrawVectorButton m_Buttonstate
        Case vxPlastic
            DrawPlasticButton m_Buttonstate
    End Select

    DrawPicWithCaption

End Sub

Private Sub CreateRegion()

    '***************************************************************************
    '*  Create region everytime you redraw a button.                           *
    '*  Because some settings may have changed the button regions              *
    '***************************************************************************
    lh = ScaleHeight
    lw = ScaleWidth

    Select Case m_ButtonStyle
        Case vxWindowsXP, vxVistaToolbar, vxInstallShield, vxXPToolbar, vxVector, vxInstallShieldToolbar, vxInstallShieldToolBar2, vxInstallShield2, vxVistaAero2
            m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 3, 3)
        Case vxGelButton
            m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 4, 4)
        Case vxPlastic
            m_lButtonRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 9, 9)
        Case Else
            m_lButtonRgn = 0 'CreateRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight)
    End Select
    SetWindowRgn UserControl.hWnd, m_lButtonRgn, True       'Set Button Region
    DeleteObject m_lButtonRgn                               'Free memory

End Sub

Private Sub DrawPicWithCaption()

    '****************************************************************************
    '* Draws a Picture in Enabled / Disabled mode along with Caption            *
    '* Also captions are drawn here calculating all rects                       *
    '* Routine to make GrayScale images is the work of Jim Jose.                *
    '****************************************************************************

    Dim PicX     As Long                       'X position of picture
    Dim PicY     As Long                       'Y Position of Picture
    Dim PicSizeW As Long                       'Picture Size
    Dim PicSizeH As Long
    Dim tmpPic   As New StdPicture             'Temp picture

    Dim hdcSrc   As Long
    Dim hOldob   As Long

    Dim lpRect   As RECT                      'RECT to draw caption
    Dim CaptionW As Long                      'Width of Caption
    Dim CaptionH As Long                      'Height of Caption
    Dim CaptionX As Long                      'Left of Caption
    Dim CaptionY As Long                      'Top of Caption
    Dim iMnemonicFlag As Long
    Dim iAlignFlag As Long
    Dim iPictureAlign As Long
    Dim iPicInfo As BITMAP

    iPictureAlign = m_PictureAlign
    If Len(Trim$(m_Caption)) = 0 Then
        If iPictureAlign = vxLeftOfCaption Then ' the default
            iPictureAlign = vxCenter
        End If
    End If
    lw = ScaleWidth                          'Width of Button
    lh = ScaleHeight                         'Height of Button

    '  Get the Caption's height and Width
    CaptionW = lw ' TextWidth(m_Caption)           'Caption's Width
    lpRect.Left = 0
    lpRect.Right = lw
    lpRect.Top = 0
    lpRect.Bottom = 0
    If Not m_bUseMnemonic Then iMnemonicFlag = DT_NOPREFIX
    Select Case m_CaptionAlign
        Case vxLeftAlign
            iAlignFlag = DT_LEFT
        Case vxCenterAlign
            iAlignFlag = DT_CENTER
        Case vxRightAlign
            iAlignFlag = DT_RIGHT
    End Select
    If m_Caption <> "" Then
        DrawTextW hDC, StrPtr(m_Caption), Len(m_Caption), lpRect, DT_CALCRECT Or DT_WORDBREAK Or iAlignFlag Or iMnemonicFlag
        CaptionW = lpRect.Right - lpRect.Left + 10
    End If
    CaptionH = lpRect.Bottom  ' TextHeight(m_Caption)         'Caption's Height

    '  Copy the original picture into a temp var
    If m_bEnabled Then
        Set tmpPic = m_PicToUse
    Else
        If Not m_PicToUse Is Nothing Then
            If m_PicToUse.Type = vbPicTypeBitmap Then
                If m_DisabledPicture Is Nothing Then
                    Set m_DisabledPicture = PictureToGrayScale(m_PicToUse)
                End If
                Set tmpPic = m_DisabledPicture
            Else
                Set tmpPic = m_PicToUse
            End If
        End If
    End If

    GetObjectAPI tmpPic.Handle, Len(iPicInfo), iPicInfo
    PicSizeW = iPicInfo.bmWidth
    PicSizeH = iPicInfo.bmHeight

    Select Case iPictureAlign
        Case vxLeftOfCaption
            PicX = (lw - (PicSizeW + CaptionW)) \ 2
            If PicX < 4 Then PicX = 4
            PicY = (lh - PicSizeH) \ 2
            CaptionX = (lw \ 2 - CaptionW \ 2) + (PicSizeW \ 2) + 3 'Some distance of 3
            If CaptionX < (PicSizeW + 8) Then CaptionX = PicSizeW + 8  'Text shouldn't draw over picture
            CaptionY = (lh \ 2 - CaptionH \ 2)

        Case vxLeftEdge
            PicX = 4
            PicY = (lh - PicSizeH) \ 2
            CaptionX = (lw \ 2) - (CaptionW \ 2) + (PicSizeW \ 2)
            If CaptionX < (PicSizeW + 8) Then CaptionX = PicSizeW + 8  'Text shouldn't draw over picture
            CaptionY = (lh \ 2 - CaptionH \ 2)

        Case vxRightEdge

            PicX = lw - PicSizeW - 4
            PicY = (lh - PicSizeH) \ 2
            CaptionX = (lw - CaptionW - 4) - PicSizeW
            CaptionY = (lh \ 2 - CaptionH \ 2)

        Case vxRightOfCaption

            PicX = (lw - (PicSizeW - CaptionW)) \ 2
            If PicX > (lw - PicSizeW - 4) Then PicX = lw - PicSizeW - 4
            PicY = (lh - PicSizeH) \ 2
            CaptionX = (lw \ 2 - CaptionW \ 2) - (PicSizeW \ 2) - 3
            If CaptionX + CaptionW < CaptionW Then
                CaptionX = (lw - CaptionW - 4) - PicSizeW
            End If
            CaptionY = lh \ 2 - (CaptionH \ 2)

        Case vxCenter
            PicX = (lw - PicSizeW) \ 2
            PicY = (lh - PicSizeH) \ 2
            CaptionX = (lw \ 2) - (CaptionW \ 2)
            CaptionY = (lh \ 2) - CaptionH \ 2

        Case vxBottomEdge
            PicX = (lw - PicSizeW) \ 2
            PicY = (lh - PicSizeH) - 4
            CaptionX = (lw \ 2 - CaptionW \ 2)
            CaptionY = (lh \ 2 - PicSizeH \ 2 - CaptionH \ 2) - 2

        Case vxBottomOfCaption
            PicX = (lw - PicSizeW) \ 2
            PicY = (lh - (PicSizeH - CaptionH)) \ 2
            If PicY > lh - PicSizeH - 4 Then PicY = lh - PicSizeH - 4
            CaptionX = (lw \ 2 - CaptionW \ 2)
            CaptionY = (lh \ 2 - PicSizeH \ 2 - CaptionH \ 2) - 2

        Case vxTopEdge
            PicX = (lw - PicSizeW) \ 2
            PicY = 4
            CaptionX = (lw \ 2 - CaptionW \ 2)
            CaptionY = (lh \ 2 + PicSizeH \ 2 - CaptionH \ 2) + 2

        Case vxTopOfCaption
            PicX = (lw - PicSizeW) \ 2
            PicY = (lh - (PicSizeH + CaptionH)) \ 2
            If PicY < 4 Then PicY = 4
            CaptionX = (lw \ 2 - CaptionW \ 2)
            CaptionY = (lh \ 2 + PicSizeH \ 2 - CaptionH \ 2) + 2

    End Select

    ' --Minor check if picture's size exceeds button size
    If PicX < 1 Then PicX = 1
    If PicY < 1 Then PicY = 1
    If PicX + PicSizeW > ScaleWidth Then PicSizeW = ScaleWidth - 8
    If PicY + PicSizeH > ScaleHeight Then PicSizeH = ScaleHeight - 8

    ' --Calculate caption rects with Caption Alignment
    If m_PicToUse Is Nothing Then
        ' --Calculate caption rects if no picture available
        Select Case m_CaptionAlign
            Case vxLeftAlign
                CaptionX = 4
            Case vxCenterAlign
                CaptionX = (lw \ 2) - (CaptionW \ 2)
            Case vxRightAlign
                CaptionX = (lw - CaptionW - 4)
        End Select
        CaptionY = (lh \ 2) - (CaptionH \ 2)
        PicX = 0
        PicY = 0
    Else
        ' --There is a picture, so calc rects with that too.. (depending on Picture Align)
        Select Case m_CaptionAlign
            Case vxLeftAlign
                If iPictureAlign = vxLeftEdge Then
                    CaptionX = PicSizeW + 8
                ElseIf iPictureAlign = vxLeftOfCaption Then
                    CaptionX = PicX + PicSizeW + 4
                ElseIf iPictureAlign = vxRightEdge Then
                    If CaptionX < 4 Then
                        CaptionX = (lw - CaptionW - 4) - PicSizeW
                    Else
                        CaptionX = 4
                    End If
                ElseIf iPictureAlign = vxRightOfCaption Then
                    CaptionX = 4
                    PicX = CaptionW + 4
                Else
                    CaptionX = 4
                End If
            Case vxRightAlign
                If iPictureAlign = vxRightEdge Then
                    CaptionX = (lw - CaptionW - 4) - PicSizeW
                ElseIf iPictureAlign = vxRightOfCaption Then

                ElseIf iPictureAlign = vxLeftEdge Then
                    CaptionX = (lw - CaptionW - 4)
                    If CaptionX < PicSizeW + 4 Then
                        CaptionX = PicSizeW + 4
                    End If
                ElseIf iPictureAlign = vxLeftOfCaption Then
                    CaptionX = (lw - CaptionW - 4)
                    PicX = CaptionX - PicSizeW - 4
                Else
                    CaptionX = (lw - CaptionW - 4)
                End If
            Case vxCenterAlign
                If iPictureAlign = vxRightEdge Then
                    If CaptionX + CaptionW < CaptionW Then
                        CaptionX = (lw - CaptionW - 4) - PicSizeW
                    Else
                        CaptionX = (lw \ 2) - (CaptionW \ 2)
                    End If
                ElseIf iPictureAlign = vxLeftOfCaption Then
                    CaptionX = (lw - (PicX + PicSizeW) - CaptionW) / 2 + PicX + PicSizeW
                End If
        End Select
    End If

    ' --Uncomment the below lines and see what happens!! Oops
    ' --The caption draws awkwardly with accesskeys!
    '    If UserControl.AccessKeys <> vbNullString Then
    '        CaptionX = CaptionX + 3
    '    End If

    '  Adjust Picture Positions
    Select Case m_ButtonStyle
        Case vxStandard, vxFlat, vxVistaToolbar, vxVistaAero2
            If m_Buttonstate = eStateDown Then
                PicX = PicX + 1
                PicY = PicY + 1
            End If
        Case vxAOL
            If m_Buttonstate = eStateDown Then
                PicX = PicX + 2
                PicY = PicY + 2
            Else
                PicX = PicX - 1
                PicY = PicY - 1
            End If
    End Select

    ' --If picture available, Set text rects with Picture
    If m_Buttonstate = eStateDown Then
        Select Case m_ButtonStyle
            Case vxStandard, vxFlat, vxVistaToolbar, vxVistaAero2
                ' --Caption pos for Standard/Flat buttons on down state
                SetRect lpRect, CaptionX + 1, CaptionY + 1, (CaptionW + CaptionX) + 1, (CaptionH + CaptionY) + 1
            Case vxAOL
                ' --Caption RECT for AOL buttons
                SetRect lpRect, CaptionX + 1, CaptionY + 2, (CaptionW + CaptionX) + 1, (CaptionH + CaptionY) + 1
            Case Else
                ' --for other buttons on down state
                SetRect lpRect, CaptionX, CaptionY, CaptionW + CaptionX, CaptionH + CaptionY
        End Select
    Else
        Select Case m_ButtonStyle
            Case vxAOL
                SetRect lpRect, CaptionX - 2, CaptionY - 2, CaptionW + CaptionX - 2, CaptionH + CaptionY - 2
            Case Else
                SetRect lpRect, CaptionX, CaptionY, CaptionW + CaptionX, CaptionH + CaptionY
                ' --For drawing Focus rect exactly around Caption
                SetRect m_CapRect, CaptionX - 2, CaptionY, CaptionW + CaptionX + 1, CaptionH + CaptionY + 1
        End Select
    End If

    ' --Draw Picture Enabled/Disabled depending of Pic type
    Select Case tmpPic.Type
        Case vbPicTypeIcon

            If m_bEnabled Then
                DrawIconEx UserControl.hDC, PicX, PicY, tmpPic.Handle, PicSizeW, PicSizeH, 0, 0, DI_NORMAL
            Else
                ' --Draw grayed picture (Thanks to Jim Jose)
                PaintGrayScale hDC, tmpPic.Handle, PicX, PicY, PicSizeW, PicSizeH
            End If

        Case vbPicTypeBitmap
            If m_bEnabled Then
                If m_bUseMaskColor Then
                    hdcSrc = CreateCompatibleDC(0)
                    hOldob = SelectObject(hdcSrc, tmpPic.Handle)
                    '                Debug.Print PicSizeW, PicSizeH, Ambient.DisplayName
                    TransparentBlt hDC, PicX, PicY, PicSizeW, PicSizeH, hdcSrc, 0, 0, PicSizeW, PicSizeH, m_lMaskColor
                    SelectObject hdcSrc, hOldob
                    DeleteDC hdcSrc
                Else
                    PaintPicture tmpPic, PicX, PicY, PicSizeW, PicSizeH
                End If
            Else
                If m_bUseMaskColor Then
                    hdcSrc = CreateCompatibleDC(0)
                    hOldob = SelectObject(hdcSrc, tmpPic.Handle)
                    TransparentBlt hDC, PicX, PicY, PicSizeW, PicSizeH, hdcSrc, 0, 0, PicSizeW, PicSizeH, m_lMaskColor
                    SelectObject hdcSrc, hOldob
                    DeleteDC hdcSrc
                Else
                    PaintPicture tmpPic, PicX, PicY, PicSizeW, PicSizeH
                End If
            End If
    End Select

    ' --At last, draw text
    SetTextColor hDC, IIf(m_bEnabled, m_bColors.tForeColor, TranslateColor(vbGrayText))

    If m_CaptionAlign <> vxCenterAlign Then
        InflateRect lpRect, -5, 0
    End If

    If Not m_WindowsNT Then
        ' --Unicode not supported
        DrawText hDC, m_Caption, Len(m_Caption), lpRect, DT_WORDBREAK Or iAlignFlag Or iMnemonicFlag     'Button looks good in SingleLine!
    Else
        ' --Supports Unicode (i.e above Windows NT)
        'Debug.Print lpRect.Left
        DrawTextW hDC, StrPtr(m_Caption), Len(m_Caption), lpRect, DT_WORDBREAK Or iAlignFlag Or iMnemonicFlag
    End If

    ' --Clear memory
    Set tmpPic = Nothing

End Sub

Private Sub SetAccessKey()
    Dim i As Long

    ' if CanGetFocus is set to true, remove the two following lines
    'UserControl.AccessKeys = ""
    'Exit Sub

    UserControl.AccessKeys = ""
    If Len(m_Caption) > 1 Then
        i = InStr(1, m_Caption, "&", vbTextCompare)
        If (i < Len(m_Caption)) And (i > 0) Then
            If Mid$(m_Caption, i + 1, 1) <> "&" Then
                AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
            Else
                i = InStr(i + 2, m_Caption, "&", vbTextCompare)
                If Mid$(m_Caption, i + 1, 1) <> "&" Then
                    AccessKeys = LCase$(Mid$(m_Caption, i + 1, 1))
                End If
            End If
        End If
    End If

End Sub

Private Sub DrawCorners(Color As Long)

    '****************************************************************************
    '* Draws four Corners of the button specified by Color                      *
    '****************************************************************************

    With UserControl
        lh = .ScaleHeight
        lw = .ScaleWidth

        SetPixel .hDC, 1, 1, Color
        SetPixel .hDC, 1, lh - 2, Color
        SetPixel .hDC, lw - 2, 1, Color
        SetPixel .hDC, lw - 2, lh - 2, Color

    End With

End Sub

Private Sub DrawStandardButton(ByVal vState As enumButtonStates)

    '****************************************************************************
    ' Draws  four different styles in one procedure                             *
    ' Makes reading the code difficult, but saves much space!! ;)               *
    '****************************************************************************

    Dim FocusRect   As RECT
    Dim tmpRect     As RECT

    lh = ScaleHeight
    lw = ScaleWidth
    SetRect m_ButtonRect, 0, 0, lw, lh

    If Not m_bEnabled Then
        '     Draws raised edge border
        DrawEdge hDC, m_ButtonRect, BDR_RAISED95, BF_RECT
    End If

    If m_bCheckBoxMode And m_bValue Then
        PaintRect ShiftColor(m_bColors.tBackColor, 0.02), m_ButtonRect
        If m_ButtonStyle <> vxFlatHover Then
            DrawEdge hDC, m_ButtonRect, BDR_SUNKEN95, BF_RECT
            If m_bShowFocus And m_bHasFocus And m_ButtonStyle = vxStandard Then
                DrawRectangle 4, 4, lw - 7, lh - 7, TranslateColor(vbApplicationWorkspace)
            End If
        End If
        Exit Sub
    End If

    Select Case vState
        Case eStateNormal
            '        CreateRegion
            PaintRect m_bColors.tBackColor, m_ButtonRect
            ' --Draws flat raised edge border
            Select Case m_ButtonStyle
                Case vxStandard
                    DrawEdge hDC, m_ButtonRect, BDR_RAISED95, BF_RECT
                Case vxFlat
                    DrawEdge hDC, m_ButtonRect, BDR_RAISEDINNER, BF_RECT
            End Select
        Case eStateOver
            PaintRect m_bColors.tBackColor, m_ButtonRect
            Select Case m_ButtonStyle
                Case vxFlatHover, vxFlat
                    ' --Draws flat raised edge border
                    DrawEdge hDC, m_ButtonRect, BDR_RAISEDINNER, BF_RECT
                Case Else
                    ' --Draws 3d raised edge border
                    DrawEdge hDC, m_ButtonRect, BDR_RAISED95, BF_RECT
            End Select

        Case eStateDown
            PaintRect m_bColors.tBackColor, m_ButtonRect
            Select Case m_ButtonStyle
                Case vxStandard
                    DrawRectangle 1, 1, lw - 2, lh - 2, &H99A8AC
                    DrawRectangle 0, 0, lw, lh, vbBlack
                Case vx3DHover
                    DrawEdge hDC, m_ButtonRect, BDR_SUNKEN95, BF_RECT
                Case vxFlatHover, vxFlat
                    ' --Draws flat pressed edge
                    DrawRectangle 0, 0, lw, lh, vbWhite
                    DrawRectangle 0, 0, lw + 1, lh + 1, TranslateColor(vbGrayText)
            End Select
    End Select

    ' --Button has focus but not downstate Or button is Default
    If m_bHasFocus Or m_bDefault Then
        If m_bShowFocus And Ambient.UserMode Then
            If m_ButtonStyle = vx3DHover Or m_ButtonStyle = vxStandard Then
                SetRect FocusRect, 4, 4, lw - 4, lh - 4
            Else
                SetRect FocusRect, 3, 3, lw - 3, lh - 3
            End If
            If m_bParentActive Then
                DrawFocusRect hDC, FocusRect
            End If
        End If
        If vState <> eStateDown And m_ButtonStyle = vxStandard Then
            SetRect tmpRect, 0, 0, lw - 1, lh - 1
            DrawEdge hDC, tmpRect, BDR_RAISED95, BF_RECT
            DrawRectangle 0, 0, lw - 1, lh - 1, TranslateColor(vbApplicationWorkspace)
            DrawRectangle 0, 0, lw, lh, vbBlack
        End If
    End If

End Sub

Private Sub DrawXPToolbar(ByVal vState As enumButtonStates)

    Dim lpRect As RECT
    Dim bColor As Long

    lh = ScaleHeight
    lw = ScaleWidth
    UserControl.BackColor = Ambient.BackColor
    bColor = m_bColors.tBackColor

    If m_bCheckBoxMode And m_bValue Then
        ' --Check with XP Toolbar!
        If m_bIsDown Then vState = eStateDown
    End If

    If m_bCheckBoxMode And m_bValue And Not m_bIsDown Then
        SetRect lpRect, 0, 0, lw, lh
        PaintRect ShiftColor(bColor, 0.2), lpRect
        '        m_bColors.tForeColor = TranslateColor(vbButtonText)
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.3)
        DrawCorners ShiftColor(bColor, -0.1)
        If m_bMouseInCtl Then
            DrawLineApi lw - 2, 1, lw - 2, lh - 2, ShiftColor(bColor, -0.04) 'Right Line
            DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, -0.07)  'Bottom
            DrawLineApi 1, lh - 3, lw - 2, lh - 3, ShiftColor(bColor, -0.04) 'Bottom
        End If
        Exit Sub
    End If

    Select Case vState
        Case eStateNormal
            '        CreateRegion
            PaintRect bColor, m_ButtonRect
        Case eStateOver
            DrawGradientEx 0, 0, lw, lh - 1, ShiftColor(bColor, 0.03), bColor, gdVertical
            DrawGradientEx 0, 0, lw, 5, ShiftColor(bColor, 0.11), ShiftColor(bColor, 0.04), gdVertical
            DrawLineApi lw - 2, 1, lw - 2, lh - 2, ShiftColor(bColor, -0.06) 'Right Line
            DrawLineApi 0, lh - 5, lw - 3, lh - 5, ShiftColor(bColor, -0.01) 'Bottom
            DrawLineApi 0, lh - 3, lw - 3, lh - 3, ShiftColor(bColor, -0.06) 'Bottom
            DrawLineApi 0, lh - 4, lw - 3, lh - 4, ShiftColor(bColor, -0.04) 'Bottom
            DrawLineApi 1, lh - 1, lw - 1, lh - 1, ShiftColor(bColor, -0.17) 'Bottom
            DrawLineApi 0, 1, 1, lh - 4, ShiftColor(bColor, 0.04)
            DrawRectangle 0, 0, lw, lh - 1, ShiftColor(bColor, -0.15)
            DrawCorners ShiftColor(bColor, -0.1)
        Case eStateDown
            PaintRect ShiftColor(bColor, -0.05), m_ButtonRect               'Paint with Darker color
            DrawLineApi 1, 1, lw - 2, 1, ShiftColor(bColor, -0.12)          'Topmost Line
            DrawLineApi 1, 2, lw - 2, 2, ShiftColor(bColor, -0.08)          'A lighter top line
            DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, -0.01) 'Bottom Line
            DrawLineApi 1, lh - 1, lw - 2, lh - 1, ShiftColor(bColor, -0.02)
            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.3)
            DrawCorners ShiftColor(bColor, -0.1)
    End Select

    If vState = eStateDown Then
        m_bColors.tForeColor = vbWhite
        '    Else
        '        m_bColors.tForeColor = TranslateColor(vbButtonText)
    End If

End Sub

Private Sub DrawWinXPButton(ByVal vState As enumButtonStates)

    '****************************************************************************
    '* Windows XP Button                                                        *
    '* I made this in just 4 hours                                              *
    '* Totally written from Scratch and coded by Me!!                           *
    '****************************************************************************

    Dim lpRect As RECT
    Dim bColor As Long

    lh = ScaleHeight
    lw = ScaleWidth
    bColor = m_bColors.tBackColor
    SetRect m_ButtonRect, 0, 0, lw, lh

    If Not m_bEnabled Then
        '        CreateRegion
        PaintRect ShiftColor(bColor, 0.03), m_ButtonRect
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.1)
        DrawCorners ShiftColor(bColor, 0.2)
        Exit Sub
    End If

    Select Case vState

        Case eStateNormal
            '        CreateRegion
            DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, 0.07), bColor, gdVertical
            DrawGradientEx 0, 0, lw, 5, ShiftColor(bColor, 0.2), ShiftColor(bColor, 0.08), gdVertical
            DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, -0.09) 'BottomMost line
            DrawLineApi 1, lh - 3, lw - 2, lh - 3, ShiftColor(bColor, -0.05) 'Bottom Line
            DrawLineApi 1, lh - 4, lw - 2, lh - 4, ShiftColor(bColor, -0.01) 'Bottom Line
            DrawLineApi lw - 2, 2, lw - 2, lh - 2, ShiftColor(bColor, -0.08) 'Right Line
            DrawLineApi 1, 1, 1, lh - 2, BlendColors(vbWhite, (bColor)) 'Left Line
            DrawLineApi 2, 2, 2, lh - 2, BlendColors(vbWhite, (bColor)) 'Left Line

        Case eStateOver
            DrawGradientEx 0, 0, lw, lh, ShiftColor(bColor, 0.07), bColor, gdVertical
            DrawGradientEx 0, 0, lw, 5, ShiftColor(bColor, 0.2), ShiftColor(bColor, 0.08), gdVertical
            DrawLineApi 1, 2, lw - 2, 2, &H89D8FD           'uppermost inner hover
            DrawLineApi 1, 1, lw - 2, 1, &HCFF0FF           'uppermost outer hover
            DrawLineApi 1, 1, 1, lh - 2, &H49BDF9           'Leftmost Line
            DrawLineApi lw - 2, 2, lw - 2, lh - 2, &H49BDF9 'Rightmost Line
            DrawLineApi 2, 2, 2, lh - 3, &H7AD2FC           'Left Line
            DrawLineApi lw - 3, 3, lw - 3, lh - 3, &H7AD2FC 'Right Line
            DrawLineApi 2, lh - 3, lw - 2, lh - 3, &H30B3F8 'BottomMost Line
            DrawLineApi 2, lh - 2, lw - 2, lh - 2, &H97E5&  'Bottom Line

        Case eStateDown
            PaintRect ShiftColor(bColor, -0.05), m_ButtonRect               'Paint with Darker color
            DrawLineApi 1, 1, lw - 2, 1, ShiftColor(bColor, -0.16)          'Topmost Line
            DrawLineApi 1, 2, lw - 2, 2, ShiftColor(bColor, -0.1)          'A lighter top line
            DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(bColor, 0.07) 'Bottom Line
            DrawLineApi 1, 1, 1, lh - 2, ShiftColor(bColor, -0.16)  'Leftmost Line
            DrawLineApi 2, 2, 2, lh - 2, ShiftColor(bColor, -0.1)   'Left Line
            DrawLineApi lw - 2, 2, lw - 2, lh - 2, ShiftColor(bColor, 0.04) 'Right Line

    End Select

    If m_bParentActive Then
        If (m_bHasFocus Or m_bDefault) And (m_Buttonstate <> eStateDown And m_Buttonstate <> eStateOver) Then
            DrawLineApi 1, 2, lw - 2, 2, &HF6D4BC           'uppermost inner hover
            DrawLineApi 1, 1, lw - 2, 1, &HFFE7CE           'uppermost outer hover
            DrawLineApi 1, 1, 1, lh - 2, &HE6AF8E           'Leftmost Line
            DrawLineApi lw - 2, 2, lw - 2, lh - 2, &HE6AF8E 'Rightmost Line
            DrawLineApi 2, 2, 2, lh - 3, &HF4D1B8           'Left Line
            DrawLineApi lw - 3, 3, lw - 3, lh - 3, &HF4D1B8 'Right Line
            DrawLineApi 2, lh - 3, lw - 2, lh - 3, &HE4AD89 'BottomMost Line
            DrawLineApi 2, lh - 2, lw - 2, lh - 2, &HEE8269 'Bottom Line
        End If
    End If

    On Error Resume Next
    If m_bParentActive Then
        If m_bShowFocus And m_bParentActive And (m_bHasFocus Or m_bDefault) Then  'show focusrect at runtime only
            SetRect lpRect, 2, 2, lw - 2, lh - 2     'I don't like this ugly focusrect!!
            DrawFocusRect hDC, lpRect
        End If
    End If

    DrawRectangle 0, 0, lw, lh, &H743C00
    DrawCorners ShiftColor(&H743C00, 0.3)

End Sub

Private Sub DrawVisualStudio2005(ByVal vState As enumButtonStates)

    'Dim lpRect As RECT
    Dim bColor As Long

    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth

    bColor = m_bColors.tBackColor
    SetRect m_ButtonRect, 0, 0, lw, lh

    If Not m_bEnabled Then
        DrawGradientEx 0, 0, lw, lh, BlendColors(ShiftColor(bColor, 0.26), vbWhite), bColor, gdVertical
    End If

    If m_bCheckBoxMode And m_bValue Then
        PaintRect &HE8E6E1, m_ButtonRect
        DrawRectangle 0, 0, lw, lh, ShiftColor(&H6F4B4B, 0.05)
        If m_Buttonstate = eStateOver Then
            PaintRect &HE2B598, m_ButtonRect
            DrawRectangle 0, 0, lw, lh, &HC56A31
        End If
        Exit Sub
    End If

    Select Case vState

        Case eStateNormal
            DrawGradientEx 0, 0, lw, lh, BlendColors(ShiftColor(bColor, 0.26), vbWhite), bColor, gdVertical
        Case eStateOver
            PaintRect &HEED2C1, m_ButtonRect
            DrawRectangle 0, 0, lw, lh, &HC56A31
        Case eStateDown
            PaintRect &HE2B598, m_ButtonRect
            DrawRectangle 0, 0, lw, lh, &H6F4B4B
    End Select

End Sub

Private Sub DrawAOLButton(ByVal vState As enumButtonStates)

    '****************************************************************************
    '* AOL (American Online) buttons.                                           *
    '****************************************************************************

    Dim lpRect As RECT
    'Dim FocusRect As RECT
    Dim bColor As Long

    bColor = m_bColors.tBackColor

    If Not m_bEnabled Then                   'Draw Disabled button
    End If

    Select Case vState
        Case eStateNormal
            '        CreateRegion
On Error GoTo h:
            UserControl.BackColor = Ambient.BackColor  'Transparent?!?

            ' --Shadows
            DrawRectangle 6, 6, lw - 9, lh - 9, &H808080
            DrawRectangle 5, 5, lw - 7, lh - 7, &HA0A0A0
            DrawRectangle 4, 4, lw - 5, lh - 5, &HC0C0C0

            SetRect lpRect, 0, 0, lw - 5, lh - 5
            PaintRect bColor, lpRect

            DrawRectangle 0, 0, lw - 4, lh - 4, ShiftColor(bColor, 0.3)

        Case eStateOver
            UserControl.BackColor = Ambient.BackColor

            ' --Shadows
            DrawRectangle 6, 6, lw - 9, lh - 9, &H808080
            DrawRectangle 5, 5, lw - 7, lh - 7, &HA0A0A0
            DrawRectangle 4, 4, lw - 5, lh - 5, &HC0C0C0

            SetRect lpRect, 0, 0, lw - 5, lh - 5
            PaintRect bColor, lpRect

            DrawRectangle 0, 0, lw - 4, lh - 4, ShiftColor(bColor, 0.3)

        Case eStateDown
            UserControl.BackColor = Ambient.BackColor

            SetRect lpRect, 3, 3, lw, lh
            PaintRect bColor, lpRect

            DrawRectangle 3, 3, lw - 3, lh - 3, ShiftColor(bColor, 0.3)

    End Select

    If m_bParentActive Then
        If m_bShowFocus And (m_bHasFocus Or m_bDefault) Then
            UserControl.DrawMode = 6        'For exact AOL effect
            If m_Buttonstate = eStateDown Then
                SetRect lpRect, 6, 6, lw - 3, lh - 3
            Else
                SetRect lpRect, 3, 3, lw - 6, lh - 6
            End If
            DrawFocusRect hDC, lpRect
        End If
    End If
h:
    'Client Site not available (Error in Ambient.BackColor) rarely occurs

End Sub

Private Sub DrawInstallShieldButton(ByVal vState As enumButtonStates)

    '****************************************************************************
    '* I saw this style while installing JetAudio in my PC.                     *
    '* I liked it, so I implemented and gave it a name 'InstallShield'          *
    '* hehe .....
    '****************************************************************************

    'Dim FocusRect As RECT
    'Dim lpRect As RECT

    lh = ScaleHeight
    lw = ScaleWidth

    If Not m_bEnabled Then
        vState = eStateNormal                 'Simple draw normal state for Disabled
    End If

    Select Case vState
        Case eStateNormal
            '        CreateRegion
            SetRect m_ButtonRect, 0, 0, lw, lh 'Maybe have changed before!

            ' --Draw upper gradient
            DrawGradientEx 0, 0, lw, lh / 2, vbWhite, m_bColors.tBackColor, gdVertical
            ' --Draw Bottom Gradient
            DrawGradientEx 0, lh / 2, lw, lh, m_bColors.tBackColor, m_bColors.tBackColor, gdVertical
            ' --Draw Inner White Border
            DrawRectangle 1, 1, lw - 2, lh, vbWhite
            ' --Draw Outer Rectangle
            DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.5)
            DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(m_bColors.tBackColor, -0.5)
        Case eStateOver

            ' --Draw upper gradient
            DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(m_bColors.tBackColor, -0.1), vbWhite, gdVertical
            ' --Draw Bottom Gradient
            DrawGradientEx 0, lh / 2, lw, lh / 2, vbWhite, ShiftColor(m_bColors.tBackColor, -0.1), gdVertical
            ' --Draw Inner White Border
            DrawRectangle 1, 1, lw - 2, lh, vbWhite
            ' --Draw Outer Rectangle
            DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.4)
            DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(m_bColors.tBackColor, -0.4)
        Case eStateDown

            ' --draw upper gradient
            DrawGradientEx 0, 0, lw, lh / 2, vbWhite, ShiftColor(m_bColors.tBackColor, -0.1), gdVertical
            ' --Draw Bottom Gradient
            DrawGradientEx 0, lh / 2, lw, lh, ShiftColor(m_bColors.tBackColor, -0.1), ShiftColor(m_bColors.tBackColor, -0.05), gdVertical
            ' --Draw Inner White Border
            DrawRectangle 1, 1, lw - 2, lh, vbWhite
            ' --Draw Outer Rectangle
            DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.23)
            DrawCorners ShiftColor(m_bColors.tBackColor, -0.1)
            DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(m_bColors.tBackColor, -0.4)

    End Select

    DrawCorners ShiftColor(m_bColors.tBackColor, 0.05)

    If m_bParentActive And (m_bHasFocus Or m_bDefault) Then
        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(m_bColors.tBackColor, -0.2)
    End If
    If m_bParentActive And m_bShowFocus And (m_bHasFocus Or m_bDefault) Then
        InflateRect m_ButtonRect, -2, -2
        DrawFocusRect hDC, m_ButtonRect
        '        DrawFocusRect hDC, m_CapRect
    End If

End Sub

Private Sub DrawGelButton(ByVal vState As enumButtonStates)

    '****************************************************************************
    ' Draws a Gelbutton                                                         *
    '****************************************************************************

    Dim lpRect    As RECT                              'RECT to fill regions
    Dim bColor    As Long                              'Original backcolor

    lh = ScaleHeight
    lw = ScaleWidth

    bColor = m_bColors.tBackColor
    Select Case vState

        Case eStateNormal                                'Normal State

            '        CreateRegion

            ' --Fill the button region with background color
            SetRect lpRect, 0, 0, lw, lh
            PaintRect bColor, lpRect

            ' --Make a shining Upper Light
            DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, vbWhite), 0.1), bColor, gdVertical
            DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.05), BlendColors(vbWhite, ShiftColor(bColor, 0.1)), gdVertical

            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.33)

        Case eStateOver
            ' --Fill the button region with background color
            SetRect lpRect, 0, 0, lw, lh
            PaintRect ShiftColor(bColor, 0.05), lpRect

            ' --Make a shining Upper Light
            DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(ShiftColor(bColor, 0.05), vbWhite), 0.15), ShiftColor(bColor, 0.05), gdVertical
            DrawGradientEx 0, 6, lw, lh - 1, bColor, BlendColors(vbWhite, ShiftColor(bColor, 0.15)), gdVertical

            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.28)

        Case eStateDown

            ' --fill the button region with background color
            SetRect lpRect, 0, 0, lw, lh
            PaintRect ShiftColor(bColor, -0.03), lpRect

            ' --Make a shining Upper Light
            DrawGradientEx 0, 0, lw, 5, ShiftColor(BlendColors(bColor, vbWhite), 0.1), bColor, gdVertical
            DrawGradientEx 0, 6, lw, lh - 1, ShiftColor(bColor, -0.08), BlendColors(vbWhite, ShiftColor(bColor, 0.07)), gdVertical

            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.36)

    End Select

    DrawCorners ShiftColor(bColor, -0.5)

End Sub

Private Sub DrawInstallShieldToolBarDAButton(ByVal vState As enumButtonStates)
    Dim lpRect As RECT
    'Dim FocusRect As RECT

    lh = ScaleHeight
    lw = ScaleWidth

    If Not m_bEnabled Then
        ' --Draw Disabled button
        PaintRect m_bColors.tBackColor, m_ButtonRect
        DrawCorners m_bColors.tBackColor
        Exit Sub
    End If

    If vState = eStateNormal Then
        '        CreateRegion
        ' --Set the rect to fill back color
        SetRect lpRect, 0, 0, lw, lh
        ' --Simply fill the button with one color (No gradient effect here!!)
        PaintRect m_bColors.tBackColor, lpRect

    ElseIf vState = eStateOver Then

        ' --Draw upper gradient
        DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(m_bColors.tBackColor, -0.1), vbWhite, gdVertical
        ' --Draw Bottom Gradient
        DrawGradientEx 0, lh / 2, lw, lh / 2, vbWhite, ShiftColor(m_bColors.tBackColor, -0.1), gdVertical
        ' --Draw Inner White Border
        DrawRectangle 1, 1, lw - 2, lh, ShiftColor(m_bColors.tBackColor, 0.6)
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.4)
        DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(m_bColors.tBackColor, -0.4)
    ElseIf vState = eStateDown Then

        ' --draw upper gradient
        DrawGradientEx 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.05), ShiftColor(m_bColors.tBackColor, 0.05), gdDownwardDiagonal
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, 0.4)
        DrawLineApi 0, 0, lw - 2, 0, ShiftColor(m_bColors.tBackColor, -0.6)
        DrawLineApi 0, 0, 0, lh - 2, ShiftColor(m_bColors.tBackColor, -0.6)
        DrawLineApi 1, 1, lw - 2, 1, ShiftColor(m_bColors.tBackColor, -0.2)
        DrawLineApi 1, 1, 1, lh - 2, ShiftColor(m_bColors.tBackColor, -0.2)
        SetPixel hDC, lw - 2, lh - 2, ShiftColor(m_bColors.tBackColor, 0.6)
    End If

End Sub

Private Sub DrawInstallShieldToolbar2Button(ByVal vState As enumButtonStates)
    Dim lpRect As RECT
    'Dim FocusRect As RECT

    lh = ScaleHeight
    lw = ScaleWidth

    If Not m_bEnabled Then
        ' --Draw Disabled button
        PaintRect m_bColors.tBackColor, m_ButtonRect
        DrawCorners m_bColors.tBackColor
        Exit Sub
    End If

    If vState = eStateNormal Then
        '        CreateRegion
        ' --Set the rect to fill back color
        SetRect lpRect, 0, 0, lw, lh
        ' --Simply fill the button with one color (No gradient effect here!!)
        PaintRect m_bColors.tBackColor, lpRect

    ElseIf vState = eStateOver Then

        ' --Draw upper gradient
        DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(m_bColors.tBackColor, 0.1), &HECECEC, gdVertical
        ' --Draw Bottom Gradient
        DrawGradientEx 0, lh / 2, lw, lh / 2, &HECECEC, ShiftColor(m_bColors.tBackColor, -0.05), gdVertical
        ' --Draw Inner White Border
        DrawRectangle 1, 1, lw - 2, lh, ShiftColor(m_bColors.tBackColor, 0.6)
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.2)
        DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(m_bColors.tBackColor, -0.2)
    ElseIf vState = eStateDown Then

        '        ' --draw upper gradient
        '        DrawGradientEx 0, 0, lw, lh / 2, vbWhite, ShiftColor(m_bColors.tBackColor, -0.1), gdVertical
        '        ' --Draw Bottom Gradient
        '        DrawGradientEx 0, lh / 2, lw, lh, ShiftColor(m_bColors.tBackColor, -0.1), ShiftColor(m_bColors.tBackColor, -0.05), gdVertical
        '        ' --Draw Inner White Border
        '        DrawRectangle 1, 1, lw - 2, lh, vbWhite
        '        ' --Draw Outer Rectangle
        '        DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.23)
        '        DrawCorners ShiftColor(m_bColors.tBackColor, -0.1)
        '        DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(m_bColors.tBackColor, -0.4)

        ' --draw upper gradient
        DrawGradientEx 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.05), ShiftColor(m_bColors.tBackColor, 0.05), gdDownwardDiagonal
        ' --Draw Outer Rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, 0.4)
        DrawLineApi 0, 0, lw - 2, 0, ShiftColor(m_bColors.tBackColor, -0.6)
        DrawLineApi 0, 0, 0, lh - 2, ShiftColor(m_bColors.tBackColor, -0.6)
        DrawLineApi 1, 1, lw - 2, 1, ShiftColor(m_bColors.tBackColor, -0.2)
        DrawLineApi 1, 1, 1, lh - 2, ShiftColor(m_bColors.tBackColor, -0.2)
        SetPixel hDC, lw - 2, lh - 2, ShiftColor(m_bColors.tBackColor, 0.6)
    End If

End Sub

Private Sub DrawInstallShieldReverseButton(ByVal vState As enumButtonStates)
    lh = ScaleHeight
    lw = ScaleWidth

    If Not m_bEnabled Then
        vState = eStateNormal                 'Simple draw normal state for Disabled
    End If

    Select Case vState
        Case eStateNormal
            '        CreateRegion
            SetRect m_ButtonRect, 0, 0, lw, lh 'Maybe have changed before!

            ' --Draw upper gradient
            DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(m_bColors.tBackColor, -0.05), ShiftColor(m_bColors.tBackColor, 0.25), gdVertical
            ' --Draw Bottom Gradient
            DrawGradientEx 0, lh / 2, lw, lh / 2, ShiftColor(m_bColors.tBackColor, 0.15), ShiftColor(m_bColors.tBackColor, -0.05), gdVertical
            ' --Draw Inner White Border
            DrawRectangle 1, 1, lw - 2, lh, ShiftColor(m_bColors.tBackColor, 0.6)
            DrawLineApi 1, lh - 2, lw - 2, lh - 2, ShiftColor(m_bColors.tBackColor, 0.3)
            ' --Draw Outer Rectangle
            DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.4)
            DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(m_bColors.tBackColor, -0.4)
            DrawCorners ShiftColor(m_bColors.tBackColor, 0.05)
        Case eStateOver

            ' --Draw upper gradient
            DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(m_bColors.tBackColor, 0.1), ShiftColor(m_bColors.tBackColor, 0.4), gdVertical
            ' --Draw Bottom Gradient
            DrawGradientEx 0, lh / 2, lw, lh / 2, ShiftColor(m_bColors.tBackColor, 0.3), ShiftColor(m_bColors.tBackColor, -0.2), gdVertical
            ' --Draw Inner White Border
            DrawRectangle 1, 1, lw - 2, lh, ShiftColor(m_bColors.tBackColor, 0.6)
            ' --Draw Outer Rectangle
            DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.5)
            DrawLineApi 2, lh - 1, lw - 2, lh - 1, ShiftColor(m_bColors.tBackColor, -0.5)
            DrawCorners ShiftColor(m_bColors.tBackColor, 0.05)
        Case eStateDown

            ' --draw upper gradient
            DrawGradientEx 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, -0.05), ShiftColor(m_bColors.tBackColor, 0.05), gdDownwardDiagonal
            ' --Draw Outer Rectangle
            DrawRectangle 0, 0, lw, lh, ShiftColor(m_bColors.tBackColor, 0.4)
            DrawLineApi 0, 0, lw - 2, 0, ShiftColor(m_bColors.tBackColor, -0.6)
            DrawLineApi 0, 0, 0, lh - 2, ShiftColor(m_bColors.tBackColor, -0.6)
            DrawLineApi 1, 1, lw - 2, 1, ShiftColor(m_bColors.tBackColor, -0.2)
            DrawLineApi 1, 1, 1, lh - 2, ShiftColor(m_bColors.tBackColor, -0.2)
    End Select

    If m_bParentActive And (m_bHasFocus Or m_bDefault) Then
        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(m_bColors.tBackColor, -0.2)
        If vState = eStateDown Then
            SetPixel hDC, lw - 2, lh - 2, ShiftColor(m_bColors.tBackColor, 0.6)
        End If
    End If
    If m_bParentActive And m_bShowFocus And (m_bHasFocus Or m_bDefault) Then
        InflateRect m_ButtonRect, -2, -2
        DrawFocusRect hDC, m_ButtonRect
        '        DrawFocusRect hDC, m_CapRect
    End If

End Sub

Private Sub DrawVistaToolbarStyle(ByVal vState As enumButtonStates)
    Static sVistaColor(4) As Long
    Static sLastBackColor As Long
    
    Dim c As Long
    
    Dim lpRect As RECT
    'Dim FocusRect As RECT
    
    If (sLastBackColor = 0) Or (m_BackColor <> sLastBackColor) Then
        If m_BackColor = vbButtonFace Then
            For c = 0 To 4
                sVistaColor(c) = GetVistaColor(c)
            Next c
        Else
            Dim R1 As Long
            Dim G1 As Long
            Dim B1 As Long
            Dim H1 As Long
            Dim L1 As Long
            Dim S1 As Long
            
            Dim R2 As Long
            Dim G2 As Long
            Dim B2 As Long
            Dim H2 As Long
            Dim L2 As Long
            Dim S2 As Long
            
            Dim H3 As Long
            Dim L3 As Long
            Dim S3 As Long
            
            R1 = GetVistaColor(5) And 255 ' R
            G1 = (GetVistaColor(5) \ 256) And 255 ' G
            B1 = (GetVistaColor(5) \ 65536) And 255 ' B
            
            ColorRGBToHLS RGB(R1, G1, B1), H1, L1, S1
            
            R2 = m_bColors.tBackColor And 255 ' R
            G2 = (m_bColors.tBackColor \ 256) And 255 ' G
            B2 = (m_bColors.tBackColor \ 65536) And 255 ' B
            
            ColorRGBToHLS RGB(R2, G2, B2), H2, L2, S2
            
            H3 = H2 - H1
            
            If H3 > 120 Then
                H3 = H3 - 240
            End If
            If H3 < -120 Then
                H3 = H3 + 240
            End If
            
            L3 = L2 - L1
            S3 = S2 - S1
            
            ' some limits:
            If L3 > 50 Then L3 = 50
            If S3 > 50 Then S3 = 50
            
            For c = 0 To 4
                sVistaColor(c) = AdjustColorWithHLS(GetVistaColor(c), H3, L3, S3)
            Next c
        End If
        sLastBackColor = m_BackColor
    End If
    
    lh = ScaleHeight
    lw = ScaleWidth

    If Not m_bEnabled Then
        ' --Draw Disabled button
        PaintRect m_bColors.tBackColor, m_ButtonRect
        DrawCorners m_bColors.tBackColor
        Exit Sub
    End If

    If vState = eStateNormal Then
        '        CreateRegion
        ' --Set the rect to fill back color
        SetRect lpRect, 0, 0, lw, lh
        ' --Simply fill the button with one color (No gradient effect here!!)
        PaintRect m_bColors.tBackColor, lpRect

    ElseIf vState = eStateOver Then

        ' --Draws a gradient effect with the folowing colors
        DrawGradientEx 1, 1, lw - 2, lh - 2, sVistaColor(0), sVistaColor(1), gdVertical

        ' --Draws a gradient in half region to give a Light Effect
        DrawGradientEx 1, lh / 1.7, lw - 2, lh - 2, sVistaColor(1), sVistaColor(1), gdVertical

        ' --Draw outside borders
        DrawRectangle 0, 0, lw, lh, sVistaColor(2)
        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(sVistaColor(4), 0.5)

    ElseIf vState = eStateDown Then

        DrawGradientEx 1, 1, lw - 2, lh - 2, sVistaColor(3), sVistaColor(4), gdVertical

        ' --Draws outside borders
        DrawRectangle 0, 0, lw, lh, sVistaColor(2)
        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(sVistaColor(4), 0.5)

    End If

    If vState = eStateDown Or vState = eStateOver Then
        DrawCorners ShiftColor(sVistaColor(2), 0.3)
    End If

End Sub

Private Sub DrawVistaAero2Style(ByVal vState As enumButtonStates)

    '*************************************************************************
    '* Draws a cool Vista Aero Style Button                                  *
    '* Use a light background color for best result                          *
    '*************************************************************************

    Dim lpRect As RECT            'Used to set rect for drawing rectangles
    Dim Color1 As Long            'Shifted / Blended color
    Dim bColor As Long            'Original back Color
    Dim iColor2 As Long

    lh = ScaleHeight
    lw = ScaleWidth
    Color1 = ShiftColor(m_bColors.tBackColor, 0.1)
    bColor = m_bColors.tBackColor

    If Not m_bEnabled Then
        ' --Draw the Disabled Button
        '        CreateRegion
        ' --Fill the button with disabled color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect bColor, lpRect

        ' --Draws outside disabled color rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.25)
        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(bColor, 0.25)
        DrawCorners ShiftColor(bColor, -0.1)
    End If

    Select Case vState

        Case eStateNormal

            '        CreateRegion

            ' --Draws a gradient in the full region
            DrawGradientEx 1, 1, lw - 1, lh, Color1, ShiftColor(bColor, -0.1), gdVertical

            ' --Draws a gradient in half region to give a glassy look
            DrawGradientEx 1, lh / 2, lw - 2, lh - 2, ShiftColor(bColor, -0.05), ShiftColor(bColor, -0.3), gdVertical

            ' --Draws border rectangle
            'DrawRectangle 0, 0, lw, lh, &H707070   'outer
            '        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(bColor, 0.1) 'inner

            '        If m_DrawDarkerRect Then
            '            DrawRectangle 0, 0, lw, lh, &HA77532
            '            'DrawRectangle 1, 1, lw - 2, lh - 2, &HF0CD3D
            '            DrawRectangle 1, 1, lw - 2, lh - 2, &HF7EBD0
            '        Else
            ' --Draw darker outer rectangle
            DrawRectangle 0, 0, lw, lh, &HC0585E               '&HCF9B23    ' &HA77532
            ' --Draw light inner rectangle
            '            DrawRectangle 1, 1, lw - 2, lh - 2, &HF7EBD0
            DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(bColor, -0.05)
            '        End If

        Case eStateOver

            ' --Draw upper gradient
            DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(m_bColors.tBackColor, 0.1), ShiftColor(m_bColors.tBackColor, 0.05), gdVertical
            ' --Draw Bottom Gradient
            DrawGradientEx 0, lh / 2, lw, lh / 2, ShiftColor(m_bColors.tBackColor, -0.01), ShiftColor(m_bColors.tBackColor, -0.3), gdVertical
            ' --Draw Inner White Border

            ' --Draws border rectangle
            DrawRectangle 0, 0, lw, lh, &HA77532   'outer
            DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(m_bColors.tBackColor, 0.2) 'inner

        Case eStateDown

            ' --Draw a gradent in full region
            DrawGradientEx 0, 0, lw, lh / 2, ShiftColor(m_bColors.tBackColor, 0.1), ShiftColor(m_bColors.tBackColor, 0.2), gdVertical
            ' --Draw Bottom Gradient
            DrawGradientEx 0, lh / 2, lw, lh / 2, ShiftColor(m_bColors.tBackColor, 0.2), ShiftColor(m_bColors.tBackColor, -0.2), gdVertical

            ' --Draws down rectangle
            DrawRectangle 0, 0, lw, lh, &H8B622C
            DrawRectangle 1, 1, lw - 2, lh, ShiftColor(&HBAB09E, 0.02)  'inner gray color three sides rectangle

    End Select

    ' --Draw a focus rectangle if button has focus

    If m_bParentActive And (vState = eStateOver) Then
        ' --Draw darker outer rectangle
        DrawRectangle 0, 0, lw, lh, &HA77532
        ' --Draw light inner rectangle
        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(m_bColors.tBackColor, 0.2)

        If (m_bShowFocus And m_bHasFocus) Then
            SetRect lpRect, 1.5, 1.5, lw - 2, lh - 2
            DrawFocusRect hDC, lpRect
        End If
    End If

    ' --Create four corners which will be common to all states
    DrawCorners ShiftColor(bColor, -0.1)

    iColor2 = TranslateColor(m_BackColorBkg)

    SetPixel hDC, 0, 0, iColor2
    SetPixel hDC, lw - 1, 0, iColor2
    SetPixel hDC, 0, lh - 1, iColor2
    SetPixel hDC, lw - 1, lh - 1, iColor2
End Sub

Private Sub DrawVistaButton(ByVal vState As enumButtonStates)

    '*************************************************************************
    '* Draws a cool Vista Aero Style Button                                  *
    '* Use a light background color for best result                          *
    '*************************************************************************

    Dim lpRect As RECT            'Used to set rect for drawing rectangles
    Dim Color1 As Long            'Shifted / Blended color
    Dim bColor As Long            'Original back Color
    Dim iColor2 As Long

    lh = ScaleHeight
    lw = ScaleWidth
    Color1 = ShiftColor(m_bColors.tBackColor, 0.1)
    bColor = m_bColors.tBackColor

    If Not m_bEnabled Then
        ' --Draw the Disabled Button
        '        CreateRegion
        ' --Fill the button with disabled color
        SetRect lpRect, 0, 0, lw, lh
        PaintRect bColor, lpRect

        ' --Draws outside disabled color rectangle
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.25)
        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(bColor, 0.25)
        DrawCorners ShiftColor(bColor, -0.1)
    End If

    Select Case vState

        Case eStateNormal

            '        CreateRegion

            ' --Draws a gradient in the full region
            DrawGradientEx 1, 1, lw - 1, lh, Color1, ShiftColor(bColor, -0.1), gdVertical

            ' --Draws a gradient in half region to give a glassy look
            DrawGradientEx 1, lh / 2, lw - 2, lh - 2, ShiftColor(bColor, -0.05), ShiftColor(bColor, -0.3), gdVertical

            ' --Draws border rectangle
            'DrawRectangle 0, 0, lw, lh, &H707070   'outer
            '        DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(bColor, 0.1) 'inner

            '        If m_DrawDarkerRect Then
            '            DrawRectangle 0, 0, lw, lh, &HA77532
            '            'DrawRectangle 1, 1, lw - 2, lh - 2, &HF0CD3D
            '            DrawRectangle 1, 1, lw - 2, lh - 2, &HF7EBD0
            '        Else
            ' --Draw darker outer rectangle
            DrawRectangle 0, 0, lw, lh, &HCF9B23    ' &HA77532
            ' --Draw light inner rectangle
            '            DrawRectangle 1, 1, lw - 2, lh - 2, &HF7EBD0
            DrawRectangle 1, 1, lw - 2, lh - 2, ShiftColor(bColor, -0.05)
            '        End If

        Case eStateOver

            ' --Make gradient in the full region
            DrawGradientEx 1, 1, lw - 2, lh, ShiftColor(&HFFF7E3, 0.02), &HFEE6B9, gdVertical

            ' --Draw gradient in half button downside to give a glass look
            DrawGradientEx 1, lh / 2, lw - 2, lh - 2, &HFEE6B9, &HFEE6B9, gdVertical

            ' --Draws border rectangle
            DrawRectangle 0, 0, lw, lh, &HA77532   'outer
            DrawRectangle 1, 1, lw - 2, lh - 2, vbWhite 'inner

        Case eStateDown

            ' --Draw a gradent in full region
            DrawGradientEx 1, 1, lw - 1, lh, &HF9EDD5, &HE6C483, gdVertical

            ' --Draw gradient in half button downside to give a glass look
            DrawGradientEx 1, lh / 2, lw - 2, lh - 2, &HE6C483, ShiftColor(&HE6C483, -0.03), gdVertical

            ' --Draws down rectangle
            DrawRectangle 0, 0, lw, lh, &H8B622C
            DrawRectangle 1, 1, lw - 2, lh, ShiftColor(&HBAB09E, 0.02)  'inner gray color three sides rectangle

    End Select

    ' --Draw a focus rectangle if button has focus

    If m_bParentActive Then
        ' --Draw darker outer rectangle
        DrawRectangle 0, 0, lw, lh, &HA77532
        ' --Draw light inner rectangle
        DrawRectangle 1, 1, lw - 2, lh - 2, &HF0CD3D

        If (m_bShowFocus And m_bHasFocus) Then
            SetRect lpRect, 1.5, 1.5, lw - 2, lh - 2
            DrawFocusRect hDC, lpRect
        End If
    End If

    ' --Create four corners which will be common to all states
    DrawCorners ShiftColor(&H707070, 0.3)

    iColor2 = TranslateColor(m_BackColorBkg)

    SetPixel hDC, 0, 0, iColor2
    SetPixel hDC, lw - 1, 0, iColor2
    SetPixel hDC, 0, lh - 1, iColor2
    SetPixel hDC, lw - 1, lh - 1, iColor2
End Sub

Private Sub DrawOutlook2007(ByVal vState As enumButtonStates)

    'Dim lpRect As RECT
    Dim bColor As Long

    lh = ScaleHeight
    lw = ScaleWidth
    bColor = m_bColors.tBackColor

    If m_bCheckBoxMode And m_bValue Then
        DrawGradientEx 0, 0, lw, lh / 2.7, &HA9D9FF, &H6FC0FF, gdVertical
        DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), &H3FABFF, &H75E1FF, gdVertical
        DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
        If m_bMouseInCtl Then
            DrawGradientEx 0, 0, lw, lh / 2.7, &H58C1FF, &H51AFFF, gdVertical
            DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), &H468FFF, &H5FD3FF, gdVertical
            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
        End If
        Exit Sub
    End If

    Select Case vState
        Case eStateNormal
            PaintRect bColor, m_ButtonRect
            DrawGradientEx 0, 0, lw, lh / 2.7, BlendColors(ShiftColor(bColor, 0.09), vbWhite), BlendColors(ShiftColor(bColor, 0.07), bColor), gdVertical
            DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), bColor, ShiftColor(bColor, 0.03), gdVertical
            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
        Case eStateOver
            DrawGradientEx 0, 0, lw, lh / 2.7, &HE1FFFF, &HACEAFF, gdVertical
            DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), &H67D7FF, &H99E4FF, gdVertical
            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
        Case eStateDown
            DrawGradientEx 0, 0, lw, lh / 2.7, &H58C1FF, &H51AFFF, gdVertical
            DrawGradientEx 0, lh / 2.7, lw, lh - (lh / 2.7), &H468FFF, &H5FD3FF, gdVertical
            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, -0.34)
    End Select

End Sub

Private Sub DrawVectorButton(ByVal vState As enumButtonStates)

    'Dim lpRect          As RECT
    Dim bColor          As Long
    Dim m_lRgn          As Long

    lw = ScaleWidth
    lh = ScaleHeight
    bColor = m_bColors.tBackColor

    Select Case vState
        Case eStateNormal
            '        CreateRegion
            PaintRect bColor, m_ButtonRect
            m_lRgn = CreateRoundRectRgn(0, lh / 12, lw, lh * 2, 24, 24)

            DrawGradientEx 0, 0, lw, lh / 1.7, ShiftColor(bColor, 0.4), ShiftColor(bColor, 0.04), gdVertical
            DrawGradientEx 0, lh / 1.7, lw, lh - (lh / 1.7), bColor, ShiftColor(bColor, 0.3), gdVertical
            PaintRegion m_lRgn, ShiftColor(bColor, 0.05)

            DrawRectangle 0, 0, lw, lh, ShiftColor(bColor, 0.2)
    End Select

    DrawCorners ShiftColor(bColor, 0.2)
End Sub

Private Sub DrawPlasticButton(ByVal vState As enumButtonStates)

    Dim lpRect As RECT
    Dim bColor As Long
    Dim m_lRgn As Long
    'Dim lp As POINT, x As Long, y As Long

    lw = ScaleWidth
    lh = ScaleHeight
    bColor = m_bColors.tBackColor

    Select Case vState
        Case eStateNormal
            '        CreateRegion
            SetRect lpRect, 0, 0, lw, lh
            PaintRect ShiftColor(bColor, -0.4), lpRect

            m_lRgn = CreateRoundRectRgn(1, 1, lw, lh, 6, 6)
            PaintRegion m_lRgn, bColor

            m_lRgn = CreateRoundRectRgn(4, 2, lw - 3, lh / 4, 6, 6)

            PaintRegion m_lRgn, ShiftColor(bColor, 0.25)

        Case eStateOver
        Case eStateDown
    End Select

End Sub

Private Sub PaintRegion(ByVal lRgn As Long, ByVal lColor As Long)

    'Fills a specified region with specified color

    Dim hBrush As Long
    Dim hOldBrush As Long

    hBrush = CreateSolidBrush(lColor)
    hOldBrush = SelectObject(hDC, hBrush)

    FillRgn hDC, lRgn, hBrush

    SelectObject hDC, hOldBrush
    DeleteObject hBrush

End Sub

Private Sub PaintRect(ByVal lColor As Long, lpRect As RECT)

    'Fills a region with specified color

    Dim hOldBrush   As Long
    Dim hBrush      As Long

    hBrush = CreateSolidBrush(lColor)
    hOldBrush = SelectObject(UserControl.hDC, hBrush)

    FillRect UserControl.hDC, lpRect, hBrush

    SelectObject UserControl.hDC, hOldBrush
    DeleteObject hBrush

End Sub

Private Function ShiftColor(Color As Long, PercentInDecimal As Single) As Long

    '****************************************************************************
    '* This routine shifts a color value specified by PercentInDecimal          *
    '* Function inspired from DCbutton                                          *
    '* All Credits goes to Noel Dacara                                          *
    '* A Littlebit modified by me                                               *
    '****************************************************************************

    Dim R As Long
    Dim G As Long
    Dim B As Long

    '  Add or remove a certain color quantity by how many percent.

    R = Color And 255
    G = (Color \ 256) And 255
    B = (Color \ 65536) And 255

    R = R + PercentInDecimal * 255       ' Percent should already
    G = G + PercentInDecimal * 255       ' be translated.
    B = B + PercentInDecimal * 255       ' Ex. 50% -> 50 / 100 = 0.5

    '  When overflow occurs, ....
    If (PercentInDecimal > 0) Then       ' RGB values must be between 0-255 only
        If (R > 255) Then R = 255
        If (G > 255) Then G = 255
        If (B > 255) Then B = 255
    Else
        If (R < 0) Then R = 0
        If (G < 0) Then G = 0
        If (B < 0) Then B = 0
    End If

    ShiftColor = R + 256& * G + 65536 * B ' Return shifted color value

End Function

Private Sub mFont_FontChanged(ByVal PropertyName As String)
    UserControl.Refresh
    RedrawButton
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    If m_bEnabled Then                           'Disabled?? get out!!
        If m_bIsSpaceBarDown Then
            m_bIsSpaceBarDown = False
            m_bIsDown = False
        End If
        If m_bCheckBoxMode Then                'Checkbox Mode?
            If KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub 'Checkboxes dont repond to Enter/Escape'
            m_bValue = Not m_bValue             'Change Value (Checked/Unchecked)
            If Not m_bValue Then                'If value unchecked then
                m_Buttonstate = eStateNormal     'Normal State
            End If
            RedrawButton
        End If
        DoEvents                               'To remove focus from other button and Do events before click event
        RaiseEvent Click                       'Now Raiseevent
    End If

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

    m_bDefault = Ambient.DisplayAsDefault
    If Not m_bEnabled Or m_bMouseInCtl Then Exit Sub
    If PropertyName = "DisplayAsDefault" Then
        RedrawButton
    End If

    If PropertyName = "BackColor" Then
        RedrawButton
    End If

End Sub

Private Sub UserControl_DblClick()

    If m_lDownButton = 1 Then                    'React to only left button
        SetCapture (hWnd)                         'Preserve hWnd on DoubleClick
        'If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
        If m_Buttonstate <> eStateDown Then
            m_Buttonstate = eStateDown
        End If
        RedrawButton
        UserControl_MouseDown m_lDownButton, m_lDShift, m_lDX, m_lDY
        RaiseEvent DblClick
    End If

End Sub

Private Sub UserControl_EnterFocus()
    m_bHasFocus = True
End Sub

Private Sub UserControl_ExitFocus()
    m_bHasFocus = False
End Sub

Private Sub UserControl_GotFocus()

    m_bHasFocus = True
    RedrawButton

End Sub

Private Sub UserControl_Initialize()

    Dim OS As OSVERSIONINFO

    ' --Get the operating system version for text drawing purposes.
    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    m_WindowsNT = ((OS.dwPlatformID And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    mRedraw = False
    InitGlobal

End Sub

Private Sub UserControl_InitProperties()

    'Initialize Properties for User Control
    'Called on designtime everytime a control is added
    
    m_ButtonStyle = vxStandard
    m_bShowFocus = True
    m_bEnabled = True
    m_Caption = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
    m_PictureAlign = vxLeftOfCaption
    m_bUseMaskColor = True
    m_lMaskColor = &HFF00FF
    m_bUseMnemonic = True
    m_CaptionAlign = vxCenterAlign
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    m_BackColor = vbButtonFace
    m_ForeColor = vbButtonText
    m_BackColorBkg = vbButtonFace
    m_bColors.tBackColor = TranslateColor(m_BackColor)
    m_bColors.tForeColor = TranslateColor(m_ForeColor)
    m_BlendDisabledPicWithBackColor = False

    SetBackColorRGB

    SetPicToUse
    SetThemeColors

    Set mFont = UserControl.Font
    UserControl.Size 950, 375
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 13                                    'Enter Key
            RaiseEvent Click
        Case 37, 38                                'Left and Up Arrows
            SendKeysAPI "+{TAB}"                      'Button should transfer focus to other ctl
        Case 39, 40                                'Right and Down Arrows
            SendKeysAPI "{TAB}"                       'Button should transfer focus to other ctl
        Case 32                                    'SpaceBar held down
            If Not m_bIsDown Then
                If Shift = 4 Then Exit Sub         'System Menu Should pop up
                m_bIsSpaceBarDown = True           'Set space bar as pressed
                If (m_bCheckBoxMode) Then          'Is CheckBoxMode??
                    m_bValue = Not m_bValue        'Toggle Check Value
                    RedrawButton
                Else
                    If m_Buttonstate <> eStateDown Then
                        m_Buttonstate = eStateDown 'Button state should be down
                        RedrawButton
                    End If
                End If
            End If

            If (Not GetCapture = UserControl.hWnd) Then
                ReleaseCapture
                SetCapture UserControl.hWnd     'No other processing until spacebar is released
            End If                              'Thanks to APIGuide
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeySpace Then
        If m_bMouseInCtl And m_bIsDown Then
            If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
            RedrawButton
        ElseIf m_bMouseInCtl And Not m_bIsDown Then   'If spacebar released over ctl
            If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver 'Draw Hover State
            RedrawButton
            RaiseEvent Click
        Else                                         'If Spacebar released outside ctl
            m_Buttonstate = eStateNormal             'Draw Normal State
            RedrawButton
            RaiseEvent Click
        End If

        If (Not GetCapture = UserControl.hWnd) Then
            SetCapture UserControl.hWnd
        Else
            If (GetCapture = UserControl.hWnd) Then
                ReleaseCapture
            End If
        End If

        RaiseEvent KeyUp(KeyCode, Shift)
        m_bIsSpaceBarDown = False
        m_bIsDown = False
    End If

End Sub

Private Sub UserControl_LostFocus()

    m_bHasFocus = False                                 'No focus
    m_bIsDown = False                                   'No down state
    m_bIsSpaceBarDown = False                           'No spacebar held
    If m_bMouseInCtl Then
        m_Buttonstate = eStateOver
    Else
        m_Buttonstate = eStateNormal
    End If
    RedrawButton

    If m_bDefault = True Then                           'If default button,
        RedrawButton                                    'Show Focus
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    m_lDownButton = Button                       'Button pressed for Dblclick
    m_lDX = x
    m_lDY = y
    m_lDShift = Shift

    If Button = 1 Then
        m_bHasFocus = True
        m_bIsDown = True

        If m_bMouseInCtl Then
            If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
        End If
        RedrawButton
    End If
    RaiseEvent MouseDown(Button, Shift, x, y)

End Sub

Private Sub SetThemeColors()

    'Sets a style colors to default colors when button initialized
    'or whenever you change the style of Button

    With m_bColors

        Select Case m_ButtonStyle

            Case vxStandard, vxFlat, vxVistaToolbar, vx3DHover, vxFlatHover
                '            .tBackColor = m_bColors.tBackColor
            Case vxWindowsXP
                .tBackColor = &HE7EBEC
            Case vxOutlook2007, vxGelButton
                .tBackColor = &HFFD1AD
                .tForeColor = &H8B4215
            Case vxXPToolbar
                .tBackColor = &HECF1F1
            Case vxAOL
                .tBackColor = &HAA6D00
                .tForeColor = vbWhite
            Case vxVistaAero
                .tBackColor = ShiftColor(&HD4D4D4, 0.06)
            Case vxInstallShield
                .tBackColor = &HE1D6D5
            Case vxVisualStudio
                '            .tBackColor = m_bColors.tBackColor
        End Select

        '        If m_ButtonStyle <> vxAOL Then .tForeColor = TranslateColor(vbButtonText)
        If m_ButtonStyle = vxFlat Or m_ButtonStyle = vxStandard Then
            m_bShowFocus = True
        Else
            m_bShowFocus = False
        End If

    End With

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim p As Point

    GetCursorPos p

    If (Not WindowFromPoint(p.x, p.y) = UserControl.hWnd) Then
        m_bMouseInCtl = False
        RaiseEvent MouseLeave
    End If

    TrackMouseLeave UserControl.hWnd

    If m_bMouseInCtl Then
        If m_bIsDown Then
            If m_Buttonstate <> eStateDown Then m_Buttonstate = eStateDown
        ElseIf Not m_bIsDown And Not m_bIsSpaceBarDown Then
            If m_Buttonstate <> eStateOver Then m_Buttonstate = eStateOver
        End If
        RedrawButton True
    End If

    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        m_bIsDown = False
        If (x > 0 And y > 0) And (x < ScaleWidth And y < ScaleHeight) Then
            If m_bCheckBoxMode Then m_bValue = Not m_bValue
            RedrawButton
            RaiseEvent Click
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Private Sub UserControl_Resize()
    Dim iH As Long
    Dim iW As Long
    
    '   At least, a checkbox will also need this much of size!!!!
    iH = UserControl.ScaleHeight
    iW = UserControl.ScaleWidth
    
    If (iH < 15) Or (iW < 15) Then
        If (iH < 15) Then
            iH = 15
        End If
        If (iW < 15) Then
            iW = 15
        End If
        UserControl.Size UserControl.ScaleX(iW, vbPixels, vbTwips), UserControl.ScaleY(iH, vbPixels, vbTwips)
    End If
    '   On resize, create button region
    CreateRegion
    RedrawButton

End Sub

'Load property values from storage

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        m_ButtonStyle = .ReadProperty("ButtonStyle", vxStandard)
        m_bShowFocus = .ReadProperty("ShowFocusRect", False) 'for vxFlat style only
        On Error Resume Next
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        If Err.Number Then
            Set UserControl.Font = Ambient.Font
        End If
        On Error GoTo 0
        m_BackColor = .ReadProperty("BackColor", vbButtonFace)
        m_BackColorBkg = .ReadProperty("BackColorBkg", vbButtonFace)
        m_bEnabled = .ReadProperty("Enabled", True)
        m_Caption = .ReadProperty("Caption", "jcbutton")
        m_bValue = .ReadProperty("Value", False)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0) 'vbdefault
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        Set m_Picture = .ReadProperty("Picture", Nothing)
        Set m_Pic16 = .ReadProperty("Pic16", Nothing)
        Set m_Pic24 = .ReadProperty("Pic24", Nothing)
        Set m_Pic20 = .ReadProperty("Pic20", Nothing)
        m_bUseMaskColor = .ReadProperty("UseMaskCOlor", False)
        m_lMaskColor = .ReadProperty("MaskColor", &HFF00FF)
        Set m_DisabledPicture = Nothing
        m_bUseMnemonic = .ReadProperty("UseMnemonic", True)
        m_bCheckBoxMode = .ReadProperty("CheckBoxMode", False)
        m_PictureAlign = .ReadProperty("PictureAlign", vxLeftOfCaption)
        m_CaptionAlign = .ReadProperty("CaptionAlign", vxCenterAlign)
        m_ForeColor = .ReadProperty("ForeColor", vbButtonText)
        m_BlendDisabledPicWithBackColor = .ReadProperty("BlendDisabledPicWithBackColor", False)

        UserControl.ForeColor = m_bColors.tForeColor
        UserControl.Enabled = m_bEnabled
        SetAccessKey
        lh = UserControl.ScaleHeight
        lw = UserControl.ScaleWidth
        On Error Resume Next
        Err.Clear
        m_lParenthWnd = GetParentFormHwnd(UserControl.Parent.hWnd)
        If Err.Number > 0 Then
            m_lParenthWnd = GetParentFormHwnd(UserControl.hWnd)
            If m_lParenthWnd = UserControl.hWnd Then
                m_lParenthWnd = 0
            End If
        End If
        On Error GoTo 0
    End With

    m_bColors.tBackColor = TranslateColor(m_BackColor)
    m_bColors.tForeColor = TranslateColor(m_ForeColor)

    SetBackColorRGB

    SetPicToUse
    UserControl_Resize

    If Ambient.UserMode Then                                                              'If we're not in design mode
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If

        If bTrack Then
            'OS supports mouse leave so subclass for it
            With UserControl

                mUserControlHwnd = .hWnd
                AttachMessage Me, .hWnd, WM_MOUSEMOVE
                AttachMessage Me, .hWnd, WM_MOUSELEAVE
                '                If UserControl.Parent.MDIChild Then
                If m_lParenthWnd <> 0 Then
                    AttachMessage Me, m_lParenthWnd, WM_NCACTIVATE
                    AttachMessage Me, m_lParenthWnd, WM_ACTIVATE
                End If
                '                Else
                '                End If
            End With
        End If
    End If

    Set mFont = UserControl.Font

End Sub

Private Sub UserControl_Show()
    If m_lParenthWnd <> 0 Then
        If GetActiveWindow = m_lParenthWnd Then
            m_bParentActive = True
        End If
    Else
        m_bParentActive = True
    End If
    RedrawButton
    Redraw = True
End Sub

Private Sub UserControl_Terminate()
    'On Error GoTo Crash:
    Set m_Picture = Nothing
    Set m_DisabledPicture = Nothing
    Set m_Pic16 = Nothing
    Set m_Pic24 = Nothing
    Set m_Pic20 = Nothing
    Set m_PicToUse = Nothing

    If mUserControlHwnd <> 0 Then
        DetachMessage Me, mUserControlHwnd, WM_MOUSEMOVE
        DetachMessage Me, mUserControlHwnd, WM_MOUSELEAVE
        If m_lParenthWnd <> 0 Then
            DetachMessage Me, m_lParenthWnd, WM_ACTIVATE
            DetachMessage Me, m_lParenthWnd, WM_NCACTIVATE
        End If
    End If

    Set mFont = Nothing
End Sub

'Write property values to storage

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "ButtonStyle", m_ButtonStyle, vxStandard
        .WriteProperty "ShowFocusRect", m_bShowFocus, False
        .WriteProperty "Enabled", m_bEnabled, True
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "BackColor", m_BackColor, vbButtonFace
        .WriteProperty "BackColorBkg", m_BackColorBkg, vbButtonFace
        .WriteProperty "Caption", m_Caption, "jcbutton1"
        .WriteProperty "ForeColor", m_ForeColor, vbButtonText
        .WriteProperty "BlendDisabledPicWithBackColor", m_BlendDisabledPicWithBackColor, False
        .WriteProperty "CheckBoxMode", m_bCheckBoxMode, False
        .WriteProperty "Value", m_bValue, False
        .WriteProperty "MousePointer", UserControl.MousePointer, 0
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "Picture", m_Picture, Nothing
        .WriteProperty "Pic16", m_Pic16, Nothing
        .WriteProperty "Pic24", m_Pic24, Nothing
        .WriteProperty "Pic20", m_Pic20, Nothing
        .WriteProperty "PictureAlign", m_PictureAlign, vxLeftOfCaption
        .WriteProperty "UseMaskCOlor", m_bUseMaskColor, False
        .WriteProperty "UseMnemonic", m_bUseMnemonic, True
        .WriteProperty "MaskColor", m_lMaskColor, &HFF00FF
        .WriteProperty "CaptionAlign", m_CaptionAlign, vxCenterAlign
    End With

End Sub

'Determine if the passed function is supported

Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean

    Dim hMod        As Long
    Dim bLibLoaded  As Boolean

    hMod = GetModuleHandleA(sModule)

    If hMod = 0 Then
        hMod = LoadLibraryA(sModule)
        If hMod Then
            bLibLoaded = True
        End If
    End If

    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If

    If bLibLoaded Then
        FreeLibrary hMod
    End If

End Function

'Track the mouse leaving the indicated window

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

    Dim tme As TRACKMOUSEEVENT_STRUCT

    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With

        If bTrackUser32 Then
            TrackMouseEvent tme
        Else
            TrackMouseEventComCtl tme
        End If
    End If

End Sub

Public Property Get BackColor() As OLE_COLOR

    BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    m_BackColor = New_BackColor
    m_bColors.tBackColor = TranslateColor(m_BackColor)
    SetBackColorRGB
    If m_BlendDisabledPicWithBackColor Then
        Set m_DisabledPicture = Nothing
    End If
    RedrawButton
    PropertyChanged "BackColor"

End Property

Public Property Get BackColorBkg() As OLE_COLOR

    BackColorBkg = m_BackColorBkg

End Property

Public Property Let BackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)

    m_BackColorBkg = New_BackColorBkg
    RedrawButton
    PropertyChanged "BackColorBkg"

End Property

Public Property Get ButtonStyle() As vbExButtonStyleConstants

    ButtonStyle = m_ButtonStyle

End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As vbExButtonStyleConstants)

    m_ButtonStyle = New_ButtonStyle
    SetThemeColors          'Set colors
    CreateRegion            'Create Region Again
    RedrawButton            'Obviously, force redraw!!!
    PropertyChanged "ButtonStyle"

End Property

Public Property Get Caption() As String
Attribute Caption.VB_MemberFlags = "200"

    Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

    m_Caption = New_Caption
    SetAccessKey
    RedrawButton
    PropertyChanged "Caption"

End Property

Public Property Get CaptionAlign() As vbExButtonCaptionAlignConstants

    CaptionAlign = m_CaptionAlign

End Property

Public Property Let CaptionAlign(ByVal New_CaptionAlign As vbExButtonCaptionAlignConstants)

    m_CaptionAlign = New_CaptionAlign
    RedrawButton
    PropertyChanged "CaptionAlign"

End Property

Public Property Get CheckBoxMode() As Boolean

    CheckBoxMode = m_bCheckBoxMode

End Property

Public Property Let CheckBoxMode(ByVal New_CheckBoxMode As Boolean)

    m_bCheckBoxMode = New_CheckBoxMode
    'If Not m_bCheckBoxMode Then m_Buttonstate = eStateNormal
    If Not m_bCheckBoxMode Then
        m_Buttonstate = eStateNormal
    End If
    RedrawButton
    PropertyChanged "Value"
    PropertyChanged "CheckBoxMode"

End Property

Public Property Get Value() As Boolean

    Value = m_bValue

End Property

Public Property Let Value(ByVal New_Value As Boolean)
    Dim iPrev As Boolean

    If m_bCheckBoxMode Then
        iPrev = m_bValue
        m_bValue = New_Value
        'If Not m_bValue Then m_Buttonstate = eStateNormal
        If Not m_bValue Then
            m_Buttonstate = eStateNormal
        End If
        RedrawButton
        PropertyChanged "Value"
        If iPrev <> m_bValue Then
            RaiseEvent Click
        End If
    Else
        m_Buttonstate = eStateNormal
        RedrawButton
        RaiseEvent Click
    End If
End Property

Public Property Get Enabled() As Boolean

    Enabled = m_bEnabled
    'UserControl.Enabled = m_enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    If New_Enabled <> m_bEnabled Then
        m_bEnabled = New_Enabled
        UserControl.Enabled = m_bEnabled
        RedrawButton
        If Not m_bCheckBoxMode Then m_Buttonstate = eStateNormal
        UserControl_MouseMove 0, 0, 0, 0
        PropertyChanged "Enabled"
    End If
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)

    Set UserControl.Font = New_Font
    Set mFont = UserControl.Font
    UserControl.Refresh
    RedrawButton
    PropertyChanged "Font"

End Property

Public Property Get ForeColor() As OLE_COLOR

    ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)

    m_ForeColor = New_ForeColor
    m_bColors.tForeColor = TranslateColor(m_ForeColor)
    UserControl.ForeColor = m_bColors.tForeColor
    UserControl_Resize
    PropertyChanged "ForeColor"

End Property

Public Property Get hWnd() As Long

    hWnd = UserControl.hWnd

End Property

Public Property Get MaskColor() As OLE_COLOR

    MaskColor = m_lMaskColor

End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)

    m_lMaskColor = New_MaskColor
    Set m_DisabledPicture = Nothing
    RedrawButton
    PropertyChanged "MaskColor"

End Property

Public Property Get MouseIcon() As IPictureDisp

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal New_Icon As IPictureDisp)

    On Error Resume Next
    If Not New_Icon Is Nothing Then
        If New_Icon.Handle = 0 Then
            Set New_Icon = Nothing
        End If
    End If
    Set UserControl.MouseIcon = New_Icon
    If (New_Icon Is Nothing) Then
        UserControl.MousePointer = 0 ' vbDefault
    Else
        UserControl.MousePointer = 99 ' vbCustom
    End If
    PropertyChanged "MouseIcon"

End Property

Public Property Get MousePointer() As MousePointerConstants

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_Cursor As MousePointerConstants)

    UserControl.MousePointer = New_Cursor
    PropertyChanged "MousePointer"

End Property


Public Property Get Picture() As IPictureDisp
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As IPictureDisp)
    If New_Picture Is Nothing Then
        Set m_Picture = Nothing
    ElseIf New_Picture.Handle = 0 Then
        Set m_Picture = Nothing
    Else
        Set m_Picture = New_Picture
    End If
    
    Set m_DisabledPicture = Nothing
    SetPicToUse
    If Not New_Picture Is Nothing Then
        RedrawButton
    Else
        UserControl_Resize
    End If
    PropertyChanged "Picture"
End Property


Public Property Get Pic16() As IPictureDisp
    Set Pic16 = m_Pic16
End Property

Public Property Set Pic16(ByVal New_Pic16 As IPictureDisp)
    If New_Pic16 Is Nothing Then
        Set m_Pic16 = Nothing
    ElseIf New_Pic16.Handle = 0 Then
        Set m_Pic16 = Nothing
    Else
        Set m_Pic16 = New_Pic16
    End If
    
    Set m_DisabledPicture = Nothing
    SetPicToUse
    If Not New_Pic16 Is Nothing Then
        RedrawButton
    Else
        UserControl_Resize
    End If
    PropertyChanged "Pic16"
End Property


Public Property Get Pic20() As IPictureDisp
    Set Pic20 = m_Pic20
End Property

Public Property Set Pic20(ByVal New_Pic20 As IPictureDisp)
    If New_Pic20 Is Nothing Then
        Set m_Pic20 = Nothing
    ElseIf New_Pic20.Handle = 0 Then
        Set m_Pic20 = Nothing
    Else
        Set m_Pic20 = New_Pic20
    End If
    
    Set m_DisabledPicture = Nothing
    SetPicToUse
    If Not New_Pic20 Is Nothing Then
        RedrawButton
    Else
        UserControl_Resize
    End If
    PropertyChanged "Pic20"
End Property


Public Property Get Pic24() As IPictureDisp
    Set Pic24 = m_Pic24
End Property

Public Property Set Pic24(ByVal New_Pic24 As IPictureDisp)
    If New_Pic24 Is Nothing Then
        Set m_Pic24 = Nothing
    ElseIf New_Pic24.Handle = 0 Then
        Set m_Pic24 = Nothing
    Else
        Set m_Pic24 = New_Pic24
    End If
    
    Set m_DisabledPicture = Nothing
    SetPicToUse
    If Not New_Pic24 Is Nothing Then
        RedrawButton
    Else
        UserControl_Resize
    End If
    PropertyChanged "Pic24"
End Property


Public Property Get PictureAlign() As vbExButtonPictureAlignConstants

    PictureAlign = m_PictureAlign

End Property

Public Property Let PictureAlign(ByVal New_PictureAlign As vbExButtonPictureAlignConstants)

    m_PictureAlign = New_PictureAlign
    If Not m_PicToUse Is Nothing Then
        RedrawButton
    End If
    PropertyChanged "PictureAlign"

End Property

Public Property Get ShowFocusRect() As Boolean

    ShowFocusRect = m_bShowFocus

End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)

    m_bShowFocus = New_ShowFocusRect
    PropertyChanged "ShowFocusRect"

End Property

Public Property Get UseMaskColor() As Boolean

    UseMaskColor = m_bUseMaskColor

End Property

Public Property Let UseMaskColor(ByVal New_UseMaskColor As Boolean)

    m_bUseMaskColor = New_UseMaskColor
    If Not m_PicToUse Is Nothing Then
        Set m_DisabledPicture = Nothing
        RedrawButton
    End If
    PropertyChanged "UseMaskColor"

End Property

Public Property Get UseMnemonic() As Boolean

    UseMnemonic = m_bUseMnemonic

End Property

Public Property Let UseMnemonic(ByVal New_UseMnemonic As Boolean)

    m_bUseMnemonic = New_UseMnemonic
    RedrawButton
    PropertyChanged "UseMnemonic"

End Property

Private Function ISubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Function

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef bConsume As Boolean) As Long
    Select Case iMsg
        Case WM_MOUSEMOVE
            If Not m_bMouseInCtl Then
                m_bMouseInCtl = True
                TrackMouseLeave hWnd
                If m_bMouseInCtl Then
                    'If Not m_bIsSpaceBarDown Then m_Buttonstate = eStateOver
                    If Not m_bIsSpaceBarDown Then
                        m_Buttonstate = eStateOver
                    End If
                End If
                RedrawButton
                RaiseEvent MouseEnter
            End If

        Case WM_MOUSELEAVE

            m_bMouseInCtl = False
            If m_bIsSpaceBarDown Then Exit Function
            If m_bEnabled Then
                m_Buttonstate = eStateNormal
            End If
            RedrawButton
            RaiseEvent MouseLeave

        Case WM_NCACTIVATE, WM_ACTIVATE
            If wParam Then
                m_bParentActive = True
                '    If m_bDefault = True Then
                '        RedrawButton
                '    End If
                RedrawButton
            Else
                m_bIsDown = False
                m_bIsSpaceBarDown = False
                m_bHasFocus = False
                m_bParentActive = False
                If m_Buttonstate <> eStateNormal Then m_Buttonstate = eStateNormal
                RedrawButton
            End If
    End Select
End Function

Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Property Get FontBold() As Boolean
    FontBold = UserControl.Font.Bold
End Property

Public Property Let FontBold(nValue As Boolean)
    If UserControl.Font.Bold <> nValue Then
        UserControl.Font.Bold = nValue
        PropertyChanged "Font"
    End If
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = UserControl.Font.Italic
End Property

Public Property Let FontItalic(nValue As Boolean)
    If UserControl.Font.Italic <> nValue Then
        UserControl.Font.Italic = nValue
        PropertyChanged "Font"
    End If
End Property

Public Property Get FontName() As String
    FontName = UserControl.Font.Name
End Property

Public Property Let FontName(nValue As String)
    If UserControl.Font.Name <> nValue Then
        UserControl.Font.Name = nValue
        PropertyChanged "Font"
    End If
End Property

Public Property Get FontSize() As Long
    FontSize = UserControl.Font.Size
End Property

Public Property Let FontSize(nValue As Long)
    If UserControl.Font.Size <> nValue Then
        UserControl.Font.Size = nValue
        PropertyChanged "Font"
    End If
End Property

Public Property Get FontStrikeThru() As Boolean
    FontStrikeThru = UserControl.Font.Strikethrough
End Property

Public Property Let FontStrikeThru(nValue As Boolean)
    If UserControl.Font.Strikethrough <> nValue Then
        UserControl.Font.Strikethrough = nValue
        PropertyChanged "Font"
    End If
End Property

Public Property Get FontUnderLine() As Boolean
    FontUnderLine = UserControl.Font.Underline
End Property

Public Property Let FontUnderLine(nValue As Boolean)
    If UserControl.Font.Underline <> nValue Then
        UserControl.Font.Underline = nValue
        PropertyChanged "Font"
    End If
End Property

Private Function PictureToGrayScale(nPic As StdPicture) As StdPicture
    Dim iPb1 As cMemPictureBox
    Dim iPb2 As cMemPictureBox
    Dim x As Long
    Dim y As Long
    Dim iColor As Long

    If nPic Is Nothing Then Exit Function

    Set iPb1 = New cMemPictureBox
    Set iPb2 = New cMemPictureBox

    Set iPb1.Picture = nPic
    iPb2.Width = iPb1.Width
    iPb2.Height = iPb1.Height

    For x = 0 To iPb1.ScaleWidth - 1
        For y = 0 To iPb1.ScaleHeight - 1
            iColor = GetPixel(iPb1.hDC, x, y)
            If iColor <> m_lMaskColor Then
                iColor = ToGray(iColor)
            End If
            SetPixel iPb2.hDC, x, y, iColor
        Next y
    Next x

    Set PictureToGrayScale = iPb2.Image
    Set iPb1 = Nothing
    Set iPb2 = Nothing
End Function

Private Function ToGray(nColor As Long) As Long
    Dim iR As Long
    Dim iG As Long
    Dim iB As Long
    Dim iC As Long

    iR = nColor And 255
    iG = (nColor \ 256) And 255
    iB = (nColor \ 65536) And 255
    iC = (0.2125 * iR + 0.7154 * iG + 0.0721 * iB)

    If m_BlendDisabledPicWithBackColor Then
        ToGray = RGB(iC / 255 * mBackColorR * 0.7 + 88, iC / 255 * mBackColorG * 0.7 + 88, iC / 255 * mBackColorB * 0.7 + 88)
    Else
        ToGray = RGB(iC * 0.6 + 90, iC * 0.6 + 90, iC * 0.6 + 90)
    End If

End Function

Private Sub SetPicToUse()
    Dim iTx As Single

    If Not mRedraw Then
        mSetPicToUsePending = True
        Exit Sub
    End If
    mSetPicToUsePending = False

    iTx = Screen.TwipsPerPixelX
    If Not m_Pic16 Is Nothing Then
        If iTx >= 15 Then ' 96 DPI
            Set m_PicToUse = m_Pic16
        ElseIf iTx >= 12 Then ' 120 DPI
            If Not m_Pic20 Is Nothing Then
                Set m_PicToUse = m_Pic20
            Else
                Set m_PicToUse = m_Pic16
            End If
        ElseIf iTx >= 10 Then ' 144 DPI
            If Not m_Pic24 Is Nothing Then
                Set m_PicToUse = m_Pic24
            ElseIf Not m_Pic20 Is Nothing Then
                Set m_PicToUse = m_Pic20
            Else
                Set m_PicToUse = m_Pic16
            End If
        ElseIf iTx >= 7 Then ' 192 DPI
            Set m_PicToUse = StretchPicNN(m_Pic16, 2)
        ElseIf iTx >= 6 Then
            If Not m_Pic20 Is Nothing Then
                Set m_PicToUse = StretchPicNN(m_Pic20, 2)
            Else
                Set m_PicToUse = StretchPicNN(m_Pic16, 2)
            End If
        ElseIf iTx >= 5 Then
            If Not m_Pic24 Is Nothing Then
                Set m_PicToUse = StretchPicNN(m_Pic24, 2)
            ElseIf Not m_Pic20 Is Nothing Then
                Set m_PicToUse = StretchPicNN(m_Pic20, 2)
            Else
                Set m_PicToUse = StretchPicNN(m_Pic16, 3)
            End If
        ElseIf iTx >= 4 Then  ' 289 to 360 DPI
            If Not m_Pic20 Is Nothing Then
                Set m_PicToUse = StretchPicNN(m_Pic20, 3)
            Else
                Set m_PicToUse = StretchPicNN(m_Pic16, 4)
            End If
        ElseIf iTx >= 3 Then   ' 361 to 480 DPI
            If Not m_Pic24 Is Nothing Then
                Set m_PicToUse = StretchPicNN(m_Pic24, 3)
            ElseIf Not m_Pic20 Is Nothing Then
                Set m_PicToUse = StretchPicNN(m_Pic20, 4)
            Else
                Set m_PicToUse = StretchPicNN(m_Pic16, 6)
            End If
        ElseIf iTx >= 2 Then   ' 481 to 720 DPI
            If Not m_Pic24 Is Nothing Then
                Set m_PicToUse = StretchPicNN(m_Pic24, 5)
            Else
                Set m_PicToUse = StretchPicNN(m_Pic16, 8)
            End If
        Else ' greater than 720 DPI
            If Not m_Pic24 Is Nothing Then
                Set m_PicToUse = StretchPicNN(m_Pic24, 10)
            Else
                Set m_PicToUse = StretchPicNN(m_Pic16, 16)
            End If
        End If
    Else
        If Not m_Picture Is Nothing Then
            Set m_PicToUse = m_Picture
        Else
            Set m_PicToUse = Nothing
        End If
    End If

End Sub

Private Function StretchPicNN(nPic As StdPicture, nFactor As Long) As StdPicture
    Dim iPB As New cMemPictureBox
    Dim iPicInfo As BITMAP
    Dim PicSizeW As Long
    Dim PicSizeH As Long

    GetObjectAPI nPic.Handle, Len(iPicInfo), iPicInfo
    PicSizeW = iPicInfo.bmWidth
    PicSizeH = iPicInfo.bmHeight

    iPB.Width = PicSizeW * nFactor
    iPB.Height = PicSizeH * nFactor

    iPB.PaintPicture nPic, 0, 0, PicSizeW * nFactor, PicSizeH * nFactor

    Set StretchPicNN = iPB.Image
    iPB.Cls
    Set iPB = Nothing
End Function

Public Property Let BlendDisabledPicWithBackColor(nValue As Boolean)
    If nValue <> m_BlendDisabledPicWithBackColor Then
        m_BlendDisabledPicWithBackColor = nValue
        PropertyChanged "BlendDisabledPicWithBackColor"
        Set m_DisabledPicture = Nothing
        RedrawButton
    End If
End Property

Public Property Get BlendDisabledPicWithBackColor() As Boolean
    BlendDisabledPicWithBackColor = m_BlendDisabledPicWithBackColor
End Property

Private Sub SetBackColorRGB()
    Dim iColor As Long
    Dim iGreatest As Byte
    Dim iFactor As Single

    iColor = m_BackColor
    iColor = TranslateColor(iColor)
    
    mBackColorR = iColor And 255
    mBackColorG = (iColor \ 256) And 255
    mBackColorB = (iColor \ 65536) And 255

    iGreatest = mBackColorR
    If mBackColorG > iGreatest Then iGreatest = mBackColorG
    If mBackColorB > iGreatest Then iGreatest = mBackColorB
    
    If iGreatest > 0 Then
        iFactor = 255 / iGreatest
    Else
        iFactor = 1
    End If

    mBackColorR = mBackColorR * iFactor
    mBackColorG = mBackColorG * iFactor
    mBackColorB = mBackColorB * iFactor

    If mBackColorR > 255 Then mBackColorR = 255
    If mBackColorG > 255 Then mBackColorG = 255
    If mBackColorB > 255 Then mBackColorB = 255
End Sub

Public Property Let Redraw(nValor As Boolean)
    If nValor <> mRedraw Then
        mRedraw = nValor
        If mRedraw Then
            If mSetPicToUsePending Then
                SetPicToUse
            End If
            If mRedrawPending Then
                RedrawButton
            End If
        End If
    End If
End Property

Public Property Get Redraw() As Boolean
    Redraw = mRedraw
End Property

Private Function GetVistaColor(nColorIndex As Long) As Long
    Select Case nColorIndex
        Case 0
            GetVistaColor = &HFDF9F1
        Case 1
            GetVistaColor = &HF8ECD0
        Case 2
            GetVistaColor = &HCA9E61
        Case 3
            GetVistaColor = &HF1DEB0
        Case 4
            GetVistaColor = &HF9F1DB
        Case 5
            GetVistaColor = 14932157 ' 15841645
    End Select
End Function
