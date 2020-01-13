VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DropShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'//
'// 名称：DropShadow(.cls)
'// 描述：为窗口画阴影
'// 作者：denglf
'// 日期：2018-05-06
'//

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type SIZEAPI
    cx As Long
    cy As Long
End Type

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Enum ShadowTypeConstants
    eRightShadow
    eBottomShadow
End Enum

Private m_eShadowType As ShadowTypeConstants
Private m_lShadowSize As Long
Private m_lShadowColor As Long

Private m_tBI As BITMAPINFO
Private m_lPtr As Long
Private m_hWnd As Long
Private m_hDC As Long
Private m_hDib As Long
Private m_hBmpOld As Long

'阴影大小
Public Property Get ShadowSize() As Long
    ShadowSize = m_lShadowSize
End Property
Public Property Let ShadowSize(ByVal lngValue As Long)
    m_lShadowSize = lngValue
End Property

'阴影颜色
Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_lShadowColor
End Property
Public Property Let ShadowColor(ByVal clrColor As OLE_COLOR)
    m_lShadowColor = clrColor
End Property

'阴影类型
Public Property Get ShadowType() As ShadowTypeConstants
    ShadowType = m_eShadowType
End Property

Public Property Let ShadowType(ByVal enmValue As ShadowTypeConstants)
    m_eShadowType = enmValue
End Property

Public Sub DropShadow(ByVal hWnd As Long, ByVal hDC As Long)
    Dim tR As RECT
    Dim tPt As POINTAPI
    Dim tSize As SIZEAPI
    Dim lhDeskWnd As Long
    Dim lhDC As Long
    Dim x As Long
    Dim y As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    
    m_hWnd = hWnd
    
    GetWindowRect m_hWnd, tR
    ClientToScreen m_hWnd, tPt
    
    x = tPt.x
    y = tPt.y
    lngWidth = tR.right - tR.left
    lngHeight = tR.bottom - tR.top
    
    '创建内存DC
    If (m_eShadowType = eBottomShadow) Then
       tSize.cx = lngWidth
       tSize.cy = m_lShadowSize
    Else
       tSize.cx = m_lShadowSize
       tSize.cy = lngHeight - m_lShadowSize
    End If
    If Not CreateDisplay(tSize.cx, tSize.cy) Then
        Exit Sub
    End If
    
    '用桌面背景填充内存DC
    lhDeskWnd = GetDesktopWindow()
    lhDC = GetDC(lhDeskWnd)
    If (m_eShadowType = eBottomShadow) Then
        BitBlt m_hDC, 0, 0, tSize.cx, tSize.cy, lhDC, x, y + lngHeight - m_lShadowSize, vbSrcCopy
    Else
        BitBlt m_hDC, 0, 0, tSize.cx, tSize.cy, lhDC, x + lngWidth - m_lShadowSize, y, vbSrcCopy
    End If
    ReleaseDC lhDeskWnd, lhDC
    
    '创建阴影
    If (m_eShadowType = eBottomShadow) Then
        CreateDropShadow True, False
    Else
        CreateDropShadow False, False
    End If
    
    '将结果拷贝到目标窗口
    GetClientRect m_hWnd, tR
    If (m_eShadowType = eBottomShadow) Then
        BitBlt hDC, tR.left, tR.bottom - m_lShadowSize, tR.right - tR.left, m_lShadowSize, m_hDC, 0, 0, vbSrcCopy
    Else
        BitBlt hDC, tR.right - m_lShadowSize, tR.top, m_lShadowSize, tR.bottom - tR.top, m_hDC, 0, 0, vbSrcCopy
    End If
    
    '清除资源
    ClearUpDisplay
    
End Sub

Private Function CreateDIB(ByVal lhDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByRef hDib As Long) As Boolean
    
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    hDib = CreateDIBSection(lhDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
    CreateDIB = (hDib <> 0)
   
End Function

Private Function CreateDisplay(ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
    
   ClearUpDisplay
   m_hDC = CreateCompatibleDC(0)
   If (m_hDC <> 0) Then
       If (CreateDIB(m_hDC, lWidth, lHeight, m_hDib)) Then
           m_hBmpOld = SelectObject(m_hDC, m_hDib)
           CreateDisplay = True
       Else
           DeleteObject m_hDC
           m_hDC = 0
       End If
   End If
   
End Function

Private Sub ClearUpDisplay()
    
    If (m_hDC <> 0) Then
        If (m_hDib <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDib
        End If
        DeleteObject m_hDC
    End If
    m_hDC = 0
    m_hDib = 0
    m_hBmpOld = 0
    m_lPtr = 0
    
End Sub

Private Property Get BytesPerScanLine() As Long
    
    ' Scans must align on dword boundaries; easy with an alpha bitmap!
    BytesPerScanLine = m_tBI.bmiHeader.biWidth * 4
    
End Property

Private Property Get DibWidth() As Long
    DibWidth = m_tBI.bmiHeader.biWidth
End Property
Private Property Get DibHeight() As Long
    DibHeight = m_tBI.bmiHeader.biHeight
End Property

Private Sub CreateDropShadow( _
    ByVal bHorizontal As Boolean, _
    ByVal bLeftTop As Boolean)
    Dim bDib() As Byte
    Dim x As Long, y As Long
    Dim lC As Long, lInitC As Long, lSize As Long
    Dim lR As Long, lG As Long, lB As Long
    Dim lNewR As Long, lNewG As Long, lNewB As Long
    Dim tSA As SAFEARRAY2D
    
    ' Get the bits in the from DIB section:
    With tSA
       .cbElements = 1
       .cDims = 2
       .Bounds(0).lLbound = 0
       .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
       .Bounds(1).lLbound = 0
       .Bounds(1).cElements = BytesPerScanLine()
       .pvData = m_lPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    
    lR = m_lShadowColor And &HFF&
    lG = (m_lShadowColor And &HFF00&) \ &H100&
    lB = (m_lShadowColor And &HFF0000) \ &H10000
    
    If (bHorizontal) Then
       ' horizontal:
       lSize = DibHeight
       If (bLeftTop) Then
       
       Else
          ' at bottom:
          For x = 0 To BytesPerScanLine - 1 Step 4
             ' if x is within lSize of the start or end, then
             ' it contributes to alpha:
             If (x < lSize * 4) Then
                lInitC = (255 * x) \ (lSize * 4)
             ElseIf (x >= (BytesPerScanLine - lSize * 4)) Then
                lInitC = (((BytesPerScanLine - x) * 255) \ (4 * lSize))
             Else
                lInitC = 255
             End If
             
             For y = 0 To DibHeight - 1
                lC = (lInitC * y) \ DibHeight
                lC = lC * 3 / 8
                bDib(x + 3, y) = 0 'lC
                bDib(x + 2, y) = ((lR * lC) / &HFF&) + ((bDib(x + 2, y) * (&HFF& - lC)) / &HFF&)
                bDib(x + 1, y) = ((lG * lC) / &HFF&) + ((bDib(x + 1, y) * (&HFF& - lC)) / &HFF&)
                bDib(x, y) = ((lB * lC) / &HFF&) + ((bDib(x, y) * (&HFF& - lC)) / &HFF&)
             Next y
             
          Next x
       End If
    Else
       ' vertical:
       lSize = BytesPerScanLine \ 4
       If (bLeftTop) Then
       Else
          ' at right:
          For y = 0 To DibHeight - 1
             ' if y is within lSize of the start or end, then
             ' it contributes to the alpha:
             
             'If (y < lSize) Then
             '   lInitC = (255 * y) \ lSize
             If (y >= (DibHeight - lSize)) Then
                lInitC = (255 * (DibHeight - y)) \ lSize
             Else
                lInitC = 255
             End If
             
             For x = 0 To BytesPerScanLine - 1 Step 4
                ' the amount of alpha depends on how far we are from the left:
                lC = (lInitC * (BytesPerScanLine - x)) \ BytesPerScanLine
                lC = lC * 3 / 8
                bDib(x + 3, y) = 0 'lC
                bDib(x + 2, y) = ((lR * lC) / &HFF&) + ((bDib(x + 2, y) * (&HFF& - lC)) / &HFF&)
                bDib(x + 1, y) = ((lG * lC) / &HFF&) + ((bDib(x + 1, y) * (&HFF& - lC)) / &HFF&)
                bDib(x, y) = ((lB * lC) / &HFF&) + ((bDib(x, y) * (&HFF& - lC)) / &HFF&)
             Next x
          Next y
       End If
    End If
     
    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
End Sub

Private Sub Class_Initialize()
    m_lShadowSize = 5
    m_lShadowColor = RGB(0, 0, 192)
End Sub

Private Sub Class_Terminate()
    ClearUpDisplay
End Sub
 
