Attribute VB_Name = "mGridOrderByColumn"
Option Explicit

Public Enum efnFlexSortSettings
    flexSortNone = 0
    flexSortGenericAscending = 1
    flexSortGenericDescending = 2
    flexSortNumericAscending = 3
    flexSortNumericDescending = 4
    flexSortStringNoCaseAscending = 5
    flexSortStringNoCaseDescending = 6
    flexSortStringAscending = 7
    flexSortStringDescending = 8
End Enum

Private mGridsArrowUpImageCollection As New Collection
Private mGridsArrowDownImageCollection As New Collection

Public Function GetGridArrowUpImage(nGrid As Object) As StdPicture
    On Error Resume Next
    Set GetGridArrowUpImage = mGridsArrowUpImageCollection(CStr(nGrid.BackColorFixed))
    On Error GoTo 0
    If GetGridArrowUpImage Is Nothing Then
        StoreGridArrowImages nGrid
    End If
    On Error Resume Next
    Set GetGridArrowUpImage = mGridsArrowUpImageCollection(CStr(nGrid.BackColorFixed))
    On Error GoTo 0
End Function

Public Function GetGridArrowDownImage(nGrid As Object) As StdPicture
    On Error Resume Next
    Set GetGridArrowDownImage = mGridsArrowDownImageCollection(CStr(nGrid.BackColorFixed))
    On Error GoTo 0
    If GetGridArrowDownImage Is Nothing Then
        StoreGridArrowImages nGrid
    End If
    On Error Resume Next
    Set GetGridArrowDownImage = mGridsArrowDownImageCollection(CStr(nGrid.BackColorFixed))
    On Error GoTo 0
End Function

Private Sub StoreGridArrowImages(nGrid As Object)
    Dim iPic As Control
    Dim iImage As StdPicture
        
    On Error Resume Next
    nGrid.Parent.Controls.Add "VB.Picturebox", "picAux1x"
    Set iPic = nGrid.Parent.Controls("picAux1x")
    On Error GoTo TheExit:
    
    iPic.Width = nGrid.Parent.ScaleX(17 + nGrid.GridLineWidth, vbPixels, nGrid.Parent.ScaleMode)
    iPic.Height = nGrid.Parent.ScaleY(11, vbPixels, nGrid.Parent.ScaleMode)
    iPic.BorderStyle = 0
    iPic.AutoRedraw = True
    iPic.BackColor = nGrid.BackColorFixed
    iPic.ForeColor = nGrid.GridColorFixed
    iPic.ScaleMode = vbPixels
    
    iPic.Line (1, 8)-(15, 8)
    iPic.Line (1, 8)-(8, 1)
    iPic.Line (2, 8)-(8, 2)
    iPic.Line (8, 2)-(14, 8)
    iPic.Line (7, 2)-(13, 8)
    Set iImage = iPic.Image
    
    mGridsArrowUpImageCollection.Add iImage, CStr(nGrid.BackColorFixed)

    iPic.Cls
    Set iPic.Picture = Nothing
    iPic.PaintPicture iImage, 0, 0, iPic.ScaleWidth, iPic.ScaleHeight, 0, iPic.ScaleHeight, iPic.ScaleWidth, -iPic.ScaleHeight, vbSrcCopy
    Set iImage = iPic.Image
    mGridsArrowDownImageCollection.Add iImage, CStr(nGrid.BackColorFixed)
    
    On Error Resume Next
    nGrid.Parent.Controls.Remove "picAux1x"

TheExit:
End Sub

