VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'//-- Name:    ExToolTip Class 1.0
'//-- Author : Mario Flores Gonzalez
'//-- E:mail : sistec_de_juarez@hotmail.com

'//-- Using this code in commercial software, it's not prohibited,
'     however, selling or asking money for parcial or complete part of the code it's prohibited,
'     please help keep this code Open Source.

'     You are welcome to upload this code on any site, as long as you keep the original credits.


'//-- Updates at http://www.geocities.com/sistec_de_juarez/ExTooltip/

'==================================================================================================
' Apis Declarations
'==================================================================================================
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByRef phiconLarge As Long, ByRef phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

'==================================================================================================
'Class Public Enumerations
'==================================================================================================

Public Enum ttIconSize
    TTIcon16 = 16            '//-- 16x16 Icon Size
    TTIcon24 = 24            '//-- 24x24 Icon Size
    TTIcon32 = 32            '//-- 32x32 Icon Size
    TTIcon48 = 48            '//-- 48x48 Icon Size
    TTIcon72 = 72            '//-- 72x72 Icon Size
End Enum

Public Enum ttIcon
    TTI_ERROR = 1           '//--Predefined Error Icon
    TTI_INFO = 2            '//--Predefined Info Icon
    TTI_WARNING = 3         '//--Predefined Warning Icon
End Enum

Public Enum ttStyleEnum
    TTStandard               '//-- Classic Rectanlge Tooltip
    TTBalloon                '//-- Ballon Style Tooltip
End Enum

Public Enum ttFillStyle
    TTNrml = 1               '//-- Normal (Solid) Fill.
    TTGdtH = 2               '//-- Gradient Fill Vertical.
    TTGdtV = 3               '//-- Gradient Fill Horizontal.
    TTPict = 4               '//-- Picture BackGround.
End Enum

'==================================================================================================
'THE POINTAPI STRUCTURE
'--------------------------------------------------------------------------------------------------
'The POINTAPI structure defines the x- and y-coordinates of a point.
'==================================================================================================

Private Type POINTAPI
    x As Long
    y As Long
End Type

'==================================================================================================
'THE RECT STRUCTURE
'--------------------------------------------------------------------------------------------------
'The RECT structure defines the coordinates of the upper-left and lower-right corners of a rectangle.
'==================================================================================================

Private Type RECT
    left   As Long
    top    As Long
    right  As Long
    bottom As Long
End Type

'==================================================================================================
'THE RGBQUAD STRUCTURE
'--------------------------------------------------------------------------------------------------
'The RGBQUAD structure describes a color consisting of relative intensities of red, green, and blue.
'==================================================================================================

Private Type RGBQUAD
    rgbBlue     As Byte
    rgbGreen    As Byte
    rgbRed      As Byte
    rgbReserved As Byte
End Type

'==================================================================================================
'THE BITMAPINFOHEADER STRUCTURE
'--------------------------------------------------------------------------------------------------
'The BITMAPINFOHEADER structure contains information about the dimensions and color format
'of a device-independent bitmap (DIB).
'==================================================================================================

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

'==================================================================================================
'THE BITMAPINFO STRUCTURE
'--------------------------------------------------------------------------------------------------
'The BITMAPINFO structure defines the dimensions and color information for a Windows
'device-independent bitmap (DIB).
'==================================================================================================

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

'==================================================================================================
'THE EXWINDOW STRUCTURE
'--------------------------------------------------------------------------------------------------
'The EXWINDOW structure Handles the API Created Window ToolTip information
'==================================================================================================

Private Type EXWINDOW
    hwnd    As Long
    x       As Long
    y       As Long
    aExt    As Long
    bExt    As Long
    Width   As Long
    Height  As Long
    Hover   As Boolean
    Flag    As Boolean
    Visible As Boolean
    Alpha   As Byte
    hRect   As RECT
    Title   As String
    Text    As String
    TextH   As Long
    TextW   As Long
    Icon    As Long
    Pic     As StdPicture
End Type


'==================================================================================================
'ExTooltip Class Private Constants
'==================================================================================================

'=====================================================
'WINDOW STYLES
Private Const GWL_STYLE         As Long = -16
Private Const GWL_EXSTYLE       As Long = -20
'=====================================================

'=====================================================
'WINDOWS SUBCLASS MESSAGES
Private Const WM_DESTROY        As Long = &H2
Private Const WM_TIMER          As Long = &H113
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_MOUSEACTIVATE  As Long = &H21

'=====================================================


'=====================================================
'EXTENDED WINDOW STYLES
Private Const WS_EX_TOOLWINDOW  As Long = &H80&
Private Const WS_EX_CLIENTEDGE  As Long = &H200
'=====================================================

'=====================================================
'WINDOWS STYLES
Private Const WS_BORDER         As Long = &H800000
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOZORDER      As Long = &H4
'=====================================================

'=====================================================
'TEXT FORMAT CONST
Private Const DT_CALCRECT       As Long = &H400
'=====================================================

'=====================================================
'SHADOW CONSTS
Private Const CS_DROPSHADOW     As Long = &H20000
Private Const GCL_STYLE         As Long = (-26)
'=====================================================

'=====================================================
'AREA REGION CONSTS
Private Const RGN_DIFF          As Long = 4
'=====================================================

'=====================================================
'DRAWICONEX CONSTS
Private Const DI_NORMAL         As Long = &H3
'=====================================================

'=====================================================
'VARIABLES USED IN PROCESS
'=====================================================
Private iFnt          As IFont        '//-- Font Value
Private m_fnt         As IFont        '//-- Font Value
Private hFntOld       As Long         '//-- Font Value
Private CPOS          As POINTAPI     '//-- Coordinate Values of the Position of the Mouse
Private CMOS          As POINTAPI     '//-- Memory Coordinate Values of the Position of the Mouse
Private bTime         As Long         '//-- API Timer Variable Counter
Private LWnd          As Long         '//-- Control Handle Window
Private M_ExWindow    As EXWINDOW     '//-- Ex Window Tooltip Values
Private mToolTipStyle As ttStyleEnum
Private mIconSize     As ttIconSize
Private mTextColor    As Long
Private mBackColor    As Long
Private mGCEnd        As Long
Private mGCStart      As Long
Private mBackStyle    As Long
Private mDelayTime    As Long
Private mShadow       As Boolean
Private mKillTime     As Long
Private mAlive        As Boolean
Private mCP(100, 255) As Byte         '//-- Holds the percent calculation
Private mCC(255, 255) As Byte         '//-- Holds the Add Colors math

'=====================================================
'VARIABLES USED TO DRAW TEMP DC
'=====================================================
Private m_hDC      As Long
Private m_ThDC     As Long
Private m_hBmp     As Long             '//-- Temporary DC (Where All The Drawings are made)
Private m_hBmpOld  As Long             '//-- It Also Provides an Anti-Flicker Draw
Private m_lWidth   As Long
Private m_lHeight  As Long
'=====================================================

'=====================================================
'VARIABLES USED TO DRAW REGION
Private rgnNorm    As Long
Private rgn2       As Long
Private rgn1       As Long
'=====================================================

Private WithEvents mclsSubOwner As TSubclass.Subclass
Attribute mclsSubOwner.VB_VarHelpID = -1
Private WithEvents mclsSubTip As TSubclass.Subclass
Attribute mclsSubTip.VB_VarHelpID = -1

Public Property Let Alive(ByVal vData As Boolean)
    mAlive = vData
End Property
Public Property Get Alive() As Boolean
    Alive = mAlive
End Property

Public Property Let Shadow(ByVal vData As Boolean)
    mShadow = vData
End Property
Public Property Get Shadow() As Boolean
    Shadow = mShadow
End Property

Public Property Get Font() As IFont
    Set Font = m_fnt
End Property
Public Property Set Font(ByRef fnt As IFont)
    Set m_fnt = fnt    'Defined By System but can change by user choice.
End Property
Public Property Let Font(ByRef fnt As IFont)
    Set m_fnt = fnt
End Property

Public Property Let IconSize(ByVal vData As ttIconSize)
    mIconSize = vData
End Property
Public Property Get IconSize() As ttIconSize
    IconSize = mIconSize
End Property

Public Property Let ToolTipStyle(ByVal Style As ttStyleEnum)
    mToolTipStyle = Style
End Property
Public Property Get ToolTipStyle() As ttStyleEnum
    ToolTipStyle = mToolTipStyle
End Property

Public Property Let TextColor(ByVal vcolor As Long)
    mTextColor = vcolor
End Property
Public Property Get TextColor() As Long
    TextColor = mTextColor
End Property

Public Property Let GradientColorStart(ByVal vcolor As Long)
    mGCStart = vcolor
End Property
Public Property Get GradientColorStart() As Long
   GradientColorStart = mGCStart
End Property

Public Property Let GradientColorEnd(ByVal vcolor As Long)
     mGCEnd = vcolor
End Property
Public Property Get GradientColorEnd() As Long
    GradientColorEnd = mGCEnd
End Property

Public Property Let BackColor(ByVal vcolor As Long)
    mBackColor = vcolor
End Property
Public Property Get BackColor() As Long
    BackColor = mBackColor
End Property

Public Property Let BackStyle(ByVal bStyle As Long)
    mBackStyle = bStyle
End Property
Public Property Get BackStyle() As Long
   BackStyle = mBackStyle
End Property

Public Property Let DelayTime(ByVal Time As Long)
    mDelayTime = Time
End Property
Public Property Get DelayTime() As Long
    DelayTime = mDelayTime
End Property

Public Property Let KillTime(ByVal Time As Long)
    mKillTime = Time
End Property
Public Property Get KillTime() As Long
     KillTime = mKillTime
End Property

Public Property Set Picture(xPic As StdPicture)
    Set M_ExWindow.Pic = xPic
End Property
Public Property Get Picture() As StdPicture
    Set Picture = M_ExWindow.Pic
End Property

'CONVERTION FUNCTION
Private Function GetLngColor(Color As Long) As Long

    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If

End Function

'DRAWS A LINE WITH A DEFINED COLOR
Private Sub DrawLine( _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal cHdc As Long, _
    ByVal Color As Long)

    Dim Pen1    As Long
    Dim Pen2    As Long
    Dim Outline As Long
    Dim POS     As POINTAPI

    Pen1 = CreatePen(0, 1, GetLngColor(Color))
    Pen2 = SelectObject(cHdc, Pen1)
    
    MoveToEx cHdc, x, y, POS
    LineTo cHdc, Width, Height
          
    SelectObject cHdc, Pen2
    DeleteObject Pen2
    DeleteObject Pen1

End Sub

'DRAWS A FILL RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawFillRectangle(ByRef dRect As RECT, ByVal Color As Long, ByVal MyHdc As Long)
    Dim hBrush    As Long
   
    hBrush = CreateSolidBrush(GetLngColor(Color))
    FillRect MyHdc, dRect, hBrush
    DeleteObject hBrush

End Sub

'DRAWS A BORDER RECTANGLE AREA OF AN SPECIFIED COLOR
Private Sub DrawRgnRectangle(ByRef hRect As Long, ByVal Color As Long, ByVal MyHdc As Long)
    Dim hBrush As Long

    hBrush = CreateSolidBrush(GetLngColor(Color))
    FrameRgn MyHdc, hRect, hBrush, 1, 1
    DeleteObject hBrush

End Sub


'DRAWS A IMAGE AREA OF AN SPECIFIED DC
Private Sub DrawBkgImgDC(ByRef dRect As RECT, ByVal m_Picture As Long)
    Dim Brush      As Long
    Dim origBrush  As Long

    If m_Picture = 0 Then Exit Sub                        '//-- In Case No Picture is Choosen
           
    Brush = CreatePatternBrush(m_Picture)                 '//-- Use Pattern Picture Draw
    origBrush = SelectObject(m_hDC, Brush)
    
    PatBlt m_hDC, 0, 0, dRect.right - dRect.left, dRect.bottom - dRect.top, vbPatCopy
      
    SelectObject m_hDC, origBrush
    DeleteObject Brush
       
End Sub

Private Function CreateNewWindow() As Long
    
    '销毁提示窗口
    If M_ExWindow.hwnd <> 0 Then
        KillTimer LWnd, 1                       '//-- Destroy the API Timer
        DestroyWindow M_ExWindow.hwnd           '//-- Destroy the ToolTip Window when the Mouse Leaves the Control
        M_ExWindow.hwnd = 0                     '//-- Reset the HWnd    Value
        CreateNewWindow = -1                    '//-- If The ToolTip is Visible Exit Function
        Exit Function
    End If
    
    M_ExWindow.Hover = False                    '//-- Reset the Hover Value
    M_ExWindow.Flag = False                     '//-- Reset the Flag  Value
   
    If mclsSubOwner Is Nothing Then
        Set mclsSubOwner = New TSubclass.Subclass
        mclsSubOwner.hwnd = LWnd
    Else
        If mclsSubOwner.hwnd <> LWnd Then '不要反复子类化窗口
             mclsSubOwner.hwnd = LWnd '重设窗口时会先卸载
        End If
    End If
    
End Function

Private Sub DrawToolTipBody()
    Dim ClientRect As RECT
    Dim hIcon      As Long
        
    '================================================================
    'DRAW THE FILL OF THE TOOLTIP (BACKGROUND,TEXT,ICON)
    '================================================================
        
    With M_ExWindow
       
        '============================================================================================
        '/-- Get The ToolTip Area Cords.
        GetClientRect .hwnd, ClientRect
        '============================================================================================
        '/-- Solid BackGround Fill
        If mBackStyle = TTNrml Then
           DrawFillRectangle ClientRect, mBackColor, m_hDC
        '============================================================================================
        '/-- Picture BackGround Fill
        ElseIf mBackStyle = TTPict Then
           If Not Nothing Is M_ExWindow.Pic Then DrawBkgImgDC ClientRect, M_ExWindow.Pic.handle
        '============================================================================================
        '/-- Gradient BackGround Fill
        Else
           DrawGradient mGCEnd, mGCStart, ClientRect, m_hDC, IIf(mBackStyle = TTGdtH, True, False)
        End If
        '============================================================================================
        '/-- Draw the Tooltip Body Border.
        DrawRgnRectangle rgn2, mTextColor, m_hDC
               
     End With
        

     DeleteObject rgnNorm  '//-- Delete Rgn
     DeleteObject rgn2     '//-- Delete Rgn
     SetBkMode m_hDC, 1    '//-- Sets the background mix mode (TRANSPARENT Background remains untouched)
     
     '============================================================================================
     '/-- Draw The Icon
     With M_ExWindow
        If .Icon <> 0 Then
            Select Case .Icon
            Case TTI_ERROR, TTI_INFO, TTI_WARNING
                If mIconSize > TTIcon24 Then
                   ExtractIconEx GetSystemDir & "\comctl32.dll", .Icon, hIcon, 0, 1
                Else
                   ExtractIconEx GetSystemDir & "\comctl32.dll", .Icon, 0, hIcon, 1
                End If
                DrawIconEx m_hDC, 12, 8 + IIf(.aExt > 0, .aExt, 0), hIcon, mIconSize, mIconSize, 0, 0, DI_NORMAL
                DestroyIcon hIcon
            Case Else
                DrawIconEx m_hDC, 12, 8 + IIf(.aExt > 0, .aExt, 0), .Icon, mIconSize, mIconSize, 0, 0, DI_NORMAL
            End Select
        End If
        '============================================================================================
        '/-- Draw the Title.
        DrawTexto .Title, True
        '============================================================================================
        '/-- Draw the Text.
        DrawTexto .Text, False, .Title
     End With
     
  End Sub

'RETRIVES THE SYSTEM DIRECTORY
Private Function GetSystemDir() As String
    Dim sSave   As String
    Dim Ret     As Long

    'Create a buffer
     sSave = Space(255)
    'Get the system directory
     Ret = GetSystemDirectory(sSave, 255)
    'Remove all unnecessary chr$(0)'s
     GetSystemDir = left$(sSave, Ret)

End Function

Public Sub ShowToolTip(ByVal aWnd As Long, ByVal Title As String, ByVal Text As String, ByVal cIcon As Long, Optional TransparentAlpha As Long = 100)
   
    With M_ExWindow
        
        LWnd = aWnd                            '//-- Pass Parameters to memory
        .Title = Title                         '//-- Pass Parameters to memory
        .Text = Text                           '//-- Pass Parameters to memory
        .Alpha = TransparentAlpha              '//-- Pass Parameters to memory
        .TextH = 0                             '//-- Reset Height Value
        .TextW = 0                             '//-- Reset Widht Value
        .Icon = cIcon                          '//-- Pass Parameters to memory
        Call SetToolTipRect(True)              '//-- Calculate Title Dimensions
        Call SetToolTipRect(False)             '//-- Calculate Text Dimensions
        .Height = M_ExWindow.TextH + IIf(Len(Trim(M_ExWindow.Title)) > 0, 13, 0)
        .Width = M_ExWindow.TextW + 20
                  
        If CreateNewWindow() <> -1 Then
            mAlive = True   '/-- Start Core Work
        End If
        
    End With
    
End Sub

Private Function CreateToolTip() As Boolean
    Dim XL As Long
    Dim YL As Long
       
    With M_ExWindow
           
        'The CreateWindowEx function creates an overlapped, pop-up, or child window with an extended style;
        'otherwise, this function is identical to the CreateWindow function.
        '//-- FYI : #32770: Is the class for dialog boxes.
        '//-- Create a "MFGTOOLTIP" window where the mouse cursor is located.
        '//-- FYI : WS_EX_TOOLWINDOW disables the ShowInTaskBar Window Style.
        .hwnd = CreateWindowEx(WS_EX_TOOLWINDOW, "#32770", "MFG-EXTOOLTIP", 0, CPOS.x, CPOS.y, .Width, .Height, 0, 0, App.hInstance, 0)
            
        If .hwnd = 0 Then Exit Function '//-- hWnd not found, system error.
                          
        .x = CPOS.x        '//-- Assign X Cursor Coords to Mem.
        .y = CPOS.y        '//-- Assign Y Cursor Coords to Mem.
    
        XL = .Width + .x   '//-- ToolTip Rectangle Right-Side.
        YL = .Height + .y  '//-- ToolTip Rectangle Bottom-Side.
        
        .Width = .Width
        .Height = .Height
        
        '============================================================================
        '//-- Calculate if ToolTip fits on Screen X Coords.
        '============================================================================
        If XL > (Screen.Width / Screen.TwipsPerPixelX) Then
            .bExt = IIf(mToolTipStyle = TTStandard, .aExt = -2, -17)
            .x = .x - .Width
        Else
            .bExt = IIf(mToolTipStyle = TTStandard, .aExt = 2, 17)
            .x = .x + 1
        End If
               
        '============================================================================
        '//-- Calculate if ToolTip fits on Screen Y Coords.
        '============================================================================
        If YL > (Screen.Height / Screen.TwipsPerPixelY) Then
            .aExt = IIf(mToolTipStyle = TTStandard, .aExt = -2, -17)
            .y = .y - .Height
        ElseIf .y - .Height - 20 < 0 Then
            .aExt = IIf(mToolTipStyle = TTStandard, .aExt = 2, 17)
            .y = .y + 1
        Else
        '//-- FYI : By default ToolTips always appear on the TOP-LEFT of the screen.
            .aExt = IIf(mToolTipStyle = TTStandard, .aExt = -2, -17)
            .y = .y - .Height
        End If
         
        
        '============================================================================
        '//-- Change the window styles, so it can look like a real Tooltip.
        If ChangeWindowStyles(.hwnd) = 0 Then Exit Function
        '============================================================================
        
        '============================================================================
        '//-- Add Shadow to the Tooltip Window, if selected. <Only Layered OS>
        SetClassLong .hwnd, GCL_STYLE, IIf(mShadow = True, GetClassLong(.hwnd, GCL_STYLE) Or CS_DROPSHADOW, 0)
        '============================================================================
                  
        GetWindowRect .hwnd, .hRect                                                    '//-- Get The New ToolTip Coords.
            
        With .hRect
            m_hDC = ThDC(.right - .left, .bottom - .top)                               '//-- Create a Temporary DC, with given Coordinates.
            DCImageDraw m_hDC, 0, 0, .right - .left, .bottom - .top                    '//-- Function that Draws the Tooltip in the temp DC.
            ShowWindow M_ExWindow.hwnd, 4                                               '//-- Displays a window in its most recent size and position. The active window remains active.
            '===============================================================================
            pDraw GetDC(M_ExWindow.hwnd), 0, 0, .right - .left, .bottom - .top, 0, 0   '//-- Draw The Temp DC into the Real Window DC.
            '===============================================================================
        End With
        
        pDestroy                '//-- Destroy the Temporary DC.
        ReleaseDC m_hDC, .hwnd  '//-- The ReleaseDC function releases the window device context (DC), freeing it for use by other applications.
        
        '//-- SubClass The Tooltip.
        If mclsSubTip Is Nothing Then
            Set mclsSubTip = New TSubclass.Subclass
            mclsSubTip.hwnd = .hwnd
        End If
              
    End With
        
    CreateToolTip = True         '//-- The MFG ExTooltip was successfully created.
       
End Function


Private Sub DCImageDraw( _
    ByVal DstDC As Long, _
    ByVal DstX As Long, _
    ByVal DstY As Long, _
    ByVal DstW As Long, _
    ByVal DstH As Long)
    
    Dim INFO       As BITMAPINFO  'INFO
    Dim Data1()    As RGBQUAD     'RGB
    Dim Data2()    As RGBQUAD     'RGB
    Dim SrcDC      As Long        'DC
    Dim DataDC     As Long        'DC
    Dim DataBmp    As Long        'BMP
    Dim DataObj    As Long        'OBJ
    Dim i          As Long        'COUNTER
    Dim F          As Long        'COUNTER
    Dim A          As Long        'COUNTER
    Dim AlphaD     As Long        'ALPHA PERCENT
    Dim AlphaS     As Long        'ALPHA PERCENT
    
    
    With INFO.bmiHeader
        .biSize = Len(INFO.bmiHeader)   '//-- Specifies the number of bytes required by the structure.
        .biWidth = DstW                 '//-- Specifies the width of the bitmap, in pixels.
        .biHeight = DstH                '//-- Specifies the height of the bitmap, in pixels.
        .biPlanes = 1                   '//-- Specifies the number of planes for the target device. This value must be set to 1.
        .biBitCount = 32                '//-- Specifies the number of bits per pixel.
        .biCompression = 0              '//-- Specifies the type of compression for a compressed bottom-up bitmap.
        .biClrUsed = 0                  '//-- Specifies the number of color indexes in the color table that are actually used by the bitmap.
        .biClrImportant = 0             '//-- Specifies the number of color indexes that are required for displaying the bitmap.
        .biSizeImage = DstW * DstH      '//-- Specifies the size, in bytes, of the image.
    End With
       
    '-----------------------------------------------------------------
    '//-- Fill the Color Table
    '-----------------------------------------------------------------
    For i = 0 To 255
        For F = 0 To 255
            A = i + F
            If A > 255 Then A = 255     '//-- Sanity Color Check
            mCC(i, F) = A
        Next F
    Next i
    
    '-----------------------------------------------------------------
    '//-- Fill the Percent Table
    '-----------------------------------------------------------------
    For i = 0 To 100
        For F = 0 To 255
            mCP(i, F) = F / 100 * i     '//-- Alpha Percent (255 RGB)
        Next F
    Next i
    
    CreateBallonRegion
    DrawToolTipBody
    
    SrcDC = CreateCompatibleDC(0)       '//-- Create a Temp DC.
       
    With M_ExWindow.hRect
       DataObj = CaptureWindow(0, .left, .top, .right - .left, .bottom - .top)
    End With
    
    SelectObject SrcDC, DataObj
    DeleteObject DataObj

    GetDCDBitsFromDevice DataDC, DataBmp, DataObj, SrcDC, DstDC, DstH, DstW, Data1, INFO, False
    GetDCDBitsFromDevice DataDC, DataBmp, DataObj, SrcDC, DstDC, DstH, DstW, Data2, INFO, True
    
   '=================================================================================================
   'A device-independent bitmap consists of two distinct parts: a BITMAPINFO structure describing
   'the dimensions and colors of the bitmap, and an array of bytes defining the pixels of the bitmap.
   'The bits in the array are packed together, but each scan line must be padded with zeroes to end
   'on a LONG boundary. If the height is positive, the origin of the bitmap is the lower-left corner.
   'If the height is negative, the origin is the upper-left corner.
   '=================================================================================================
   
   For i = 0 To DstW - 1
        For F = 0 To DstH - 1
            AlphaS = 100 - M_ExWindow.Alpha
            AlphaD = M_ExWindow.Alpha
            Data2(i, F).rgbBlue = mCC(mCP(AlphaD, Data1(i, F).rgbBlue), mCP(AlphaS, Data2(i, F).rgbBlue))
            Data2(i, F).rgbGreen = mCC(mCP(AlphaD, Data1(i, F).rgbGreen), mCP(AlphaS, Data2(i, F).rgbGreen))
            Data2(i, F).rgbRed = mCC(mCP(AlphaD, Data1(i, F).rgbRed), mCP(AlphaS, Data2(i, F).rgbRed))
        Next F
   Next i

   '=================================================================================================
   ' The SetDIBitsToDevice function sets the pixels in the specified rectangle on the device that is
   ' associated with the destination device context using color data from a device-independent bitmap (DIB).
   '=================================================================================================
    
   '//-- Draw the new buffer Bits into the Destination DC.
    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data2(0, 0), INFO, 0
    
    Erase Data1                                     'Free Array
    Erase Data2                                     'Free Array
    DeleteDC SrcDC                                  'Free Mem
             
End Sub

Private Sub GetDCDBitsFromDevice( _
    ByRef DataHDC As Long, _
    ByRef DataBmp As Long, _
    ByRef DataObj As Long, _
    ByRef DataSRC As Long, _
    ByRef DataDST As Long, _
    ByRef DataHGT As Long, _
    ByRef DataWDT As Long, _
    ByRef DataC() As RGBQUAD, _
    ByRef DataINF As BITMAPINFO, _
    ByVal DataXOR As Boolean)
    
    DataHDC = CreateCompatibleDC(DataSRC)                                                   '//--Create a memory device context (DC) compatible with the specified device.
    DataBmp = CreateCompatibleBitmap(DataDST, DataWDT, DataHGT)                             '//--Create a bitmap compatible with the device that is associated with the specified device context.
    DataObj = SelectObject(DataHDC, DataBmp)                                                '//--Select object into the specified device context.
    ReDim DataC(0 To DataWDT - 1, 0 To DataHGT - 1)                                         '//--Pointer to a new safe array bound structure containing the new array bound.
    BitBlt DataHDC, 0, 0, DataWDT, DataHGT, IIf(DataXOR, DataSRC, DataDST), 0, 0, vbSrcCopy '//--Perform a bit-block transfer of the color data corresponding to the rectangle of pixels from the specified DC.
    GetDIBits DataHDC, DataBmp, 0, DataHGT, DataC(0, 0), DataINF, 0                         '//--Retrieve the bits of the specified bitmap and copy them into a buffer using the specified format.
    DataBmp = SelectObject(DataHDC, DataObj)                                                '//--Select object into the specified device context. The new object replaces the previous object of the same type.
    DeleteObject DataBmp                                                                    '//--Delete logical bitmap, freeing all system resources associated with the object. After the object is deleted, the specified handle is no longer valid.
    DeleteDC DataHDC                                                                        '//--Delete the specified device context (DC)

End Sub

Private Sub CreateBallonRegion()
    Dim H       As Long
    Dim W       As Long
    Dim i       As Long


    H = IIf(M_ExWindow.aExt <> 0, Abs(M_ExWindow.aExt) + M_ExWindow.Height, M_ExWindow.Height)
    W = M_ExWindow.Width
    
    If mToolTipStyle = TTStandard Then
        rgn2 = CreateRectRgn(0, 0, W, H)
        Exit Sub
    End If
    
    
    rgnNorm = CreateRectRgn(0, 0, W, H)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
       

    With M_ExWindow
              
        SetAreaRegion IIf(.bExt < 0, .Width, 0), IIf(.aExt > 0, 0, H), IIf(.bExt < 0, .Width - 17, 18), IIf(.aExt > 0, 17, H - 17), rgn1, rgn2, rgnNorm, True
        SetAreaRegion IIf(.bExt < 0, .Width - 35, 35), IIf(.aExt > 0, 0, H), IIf(.bExt < 0, W - .Width, W), IIf(.aExt > 0, 17, H - 17), rgn1, rgn2, rgnNorm, False
    
        For i = 0 To IIf(.bExt < 0, 17, 15)
            SetAreaRegion IIf(.bExt < 0, W - 18 - i, 19 + i), IIf(.aExt > 0, i, H - i), IIf(.bExt < 0, W - 35, 35), IIf(.aExt > 0, i + 1, H - i - 1), rgn1, rgn2, rgnNorm, IIf((i + 1) Mod 2 = 1, True, False)
        Next i
         
      
        '//-- TOP LEFT CORNER ====================================
        SetAreaRegion 0, IIf(.aExt > 0, 17, 0), 5, IIf(.aExt > 0, 18, 1), rgn1, rgn2, rgnNorm, True
        SetAreaRegion 0, IIf(.aExt > 0, 18, 1), 4, IIf(.aExt > 0, 19, 2), rgn1, rgn2, rgnNorm, False
        SetAreaRegion 0, IIf(.aExt > 0, 19, 2), 2, IIf(.aExt > 0, 20, 3), rgn1, rgn2, rgnNorm, True
        SetAreaRegion 0, IIf(.aExt > 0, 20, 3), 1, IIf(.aExt > 0, 22, 5), rgn1, rgn2, rgnNorm, False
        '//-- ====================================================
        
        '//-- BOTTOM LEFT CORNER =================================
        SetAreaRegion 0, H - 5 + IIf(.aExt > 0, 0, .aExt), 1, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, True
        SetAreaRegion 1, H - 3 + IIf(.aExt > 0, 0, .aExt), 2, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, False
        SetAreaRegion 2, H - 2 + IIf(.aExt > 0, 0, .aExt), 3, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, True
        SetAreaRegion 3, H + IIf(.aExt > 0, 0, .aExt), 5, H - 1 + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, False
        '//-- ====================================================
        
        '//-- TOP RIGHT CORNER ===================================
        SetAreaRegion W - 5, IIf(.aExt > 0, 17, 0), W, IIf(.aExt > 0, 18, 1), rgn1, rgn2, rgnNorm, True
        SetAreaRegion W - 3, IIf(.aExt > 0, 18, 1), W, IIf(.aExt > 0, 19, 2), rgn1, rgn2, rgnNorm, False
        SetAreaRegion W - 2, IIf(.aExt > 0, 19, 2), W, IIf(.aExt > 0, 20, 3), rgn1, rgn2, rgnNorm, True
        SetAreaRegion W - 1, IIf(.aExt > 0, 20, 3), W, IIf(.aExt > 0, 22, 5), rgn1, rgn2, rgnNorm, False
        '//-- ====================================================
    
        '//-- BOTTOM RIGHT CORNER ================================
        SetAreaRegion W - 1, H - 5 + IIf(.aExt > 0, 0, .aExt), W, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, True
        SetAreaRegion W - 2, H - 3 + IIf(.aExt > 0, 0, .aExt), W - 1, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, False
        SetAreaRegion W - 3, H - 2 + IIf(.aExt > 0, 0, .aExt), W - 2, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, True
        SetAreaRegion W - 5, H - 1 + IIf(.aExt > 0, 0, .aExt), W - 3, H + IIf(.aExt > 0, 0, .aExt), rgn1, rgn2, rgnNorm, False
        '//-- ====================================================
        
    End With
        
    Call SetWindowRgn(M_ExWindow.hwnd, rgnNorm, True)

End Sub

Private Sub SetAreaRegion(ByRef A1 As Long, ByRef B1 As Long, ByRef A2 As Long, ByRef B2 As Long, ByRef R1 As Long, ByRef R2 As Long, ByRef RN As Long, ByVal aOR As Boolean)
   
    R1 = CreateRectRgn(A1, B1, A2, B2)
    CombineRgn IIf(aOR, R2, RN), IIf(aOR, RN, R2), R1, RGN_DIFF
    DeleteObject R1
    
End Sub

'======================================================================
'DRAWS THE TEXT ON THE TOOLTIP
Private Sub DrawTexto(ByVal ThisText As String, ByVal Title As Boolean, Optional TitleST As String)
    Dim TextRect As RECT
             
    Set iFnt = Font                             '//--New Font
    iFnt.Bold = IIf(Title, True, False)         '//--Bold State
    hFntOld = SelectObject(m_hDC, iFnt.hFont)   '//--Use the New Font
        
    DrawText m_hDC, ThisText, Len(ThisText), TextRect, DT_CALCRECT  '//-- calculate text Dimensions using DT_CALCRECT
    
    With TextRect
     
        If Title Then
            .top = 1 + IIf(M_ExWindow.Icon <> 0, (mIconSize / 2), 0) + IIf(M_ExWindow.aExt > 0, M_ExWindow.aExt, 0)
            .left = (20 + IIf(M_ExWindow.aExt > 0, M_ExWindow.aExt, 0))
        Else
            DrawText m_hDC, TitleST, Len(TitleST), TextRect, DT_CALCRECT
            .top = IIf(Len(Trim(M_ExWindow.Title)) > 0, .bottom + IIf(M_ExWindow.Icon <> 0, mIconSize, 3), (.bottom / 2) - 5) + IIf(M_ExWindow.aExt > 0, M_ExWindow.aExt, 0)
            If Len(Trim(M_ExWindow.Title)) = 0 And M_ExWindow.Icon <> 0 Then
                .top = .top + mIconSize
            End If
        End If
    
        If M_ExWindow.Icon <> 0 And Title Then
            .left = .left + mIconSize
        Else
            .top = .top + 5
            .left = 10
        End If
    
       .right = .right + .left
       .bottom = .bottom + .top
    
    End With
              
    SetTextColor m_hDC, GetLngColor(mTextColor)                 '//-- Select Custom Color for Text
    DrawText m_hDC, ThisText, Len(ThisText), TextRect, &H100    '//-- Draw Text on DC
    SelectObject m_hDC, hFntOld                                 '//-- Delete the Used Font
      
End Sub
'======================================================================

Private Sub SetToolTipRect(ByVal Title As Boolean)
    Dim FDC      As Long
    Dim TextRect As RECT

    FDC = CreateDC("DISPLAY", "", "", ByVal 0&)   '//--Create a DC
          
    If Title And Len(M_ExWindow.Title) = 0 Then M_ExWindow.Title = " "
    
    Set iFnt = Font                             '//--New Font
    iFnt.Bold = IIf(Title, True, False)         '//--Bold State
    hFntOld = SelectObject(FDC, iFnt.hFont)     '//--Use the New Font

    DrawText FDC, IIf(Title, M_ExWindow.Title, M_ExWindow.Text), Len(IIf(Title, M_ExWindow.Title, M_ExWindow.Text)), TextRect, DT_CALCRECT
    M_ExWindow.TextH = M_ExWindow.TextH + TextRect.bottom
    M_ExWindow.TextW = IIf(M_ExWindow.TextW > TextRect.right, M_ExWindow.TextW, TextRect.right)
              
    If (M_ExWindow.Icon <> 0 And Title) Then
        M_ExWindow.TextW = M_ExWindow.TextW + 25 + mIconSize
        M_ExWindow.TextH = M_ExWindow.TextH + mIconSize
    End If
                     
    SelectObject FDC, hFntOld                   '//--Delete the Used Font
    DeleteDC FDC                                '//--Delete the Used DC
            
End Sub

'CHECKS-CREATES CORRECT DIMENSIONS OF THE TEMP DC
Private Function ThDC(Width As Long, Height As Long) As Long
   If m_ThDC = 0 Then
      If (Width > 0) And (Height > 0) Then
         pCreate Width, Height
      End If
   Else
      If Width > m_lWidth Or Height > m_lHeight Then
         pCreate Width, Height
      End If
   End If
   ThDC = m_ThDC
End Function

'CREATES THE TEMP DC
Private Sub pCreate(ByVal Width As Long, ByVal Height As Long)
    Dim lhDCC As Long
    
    pDestroy
    lhDCC = CreateDC("DISPLAY", "", "", ByVal 0&)
    If Not (lhDCC = 0) Then
        m_ThDC = CreateCompatibleDC(lhDCC)
        If Not (m_ThDC = 0) Then
            m_hBmp = CreateCompatibleBitmap(lhDCC, Width, Height)
            If Not (m_hBmp = 0) Then
                m_hBmpOld = SelectObject(m_ThDC, m_hBmp)
                If Not (m_hBmpOld = 0) Then
                    m_lWidth = Width
                    m_lHeight = Height
                    DeleteDC lhDCC
                    Exit Sub
                End If
            End If
        End If
        DeleteDC lhDCC
        pDestroy
    End If
    
End Sub

'DRAWS THE TEMP DC
Private Sub pDraw( _
    ByVal hDC As Long, _
    Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
    Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
    Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
   )
   If WidthSrc <= 0 Then WidthSrc = M_ExWindow.Width
   If HeightSrc <= 0 Then HeightSrc = M_ExWindow.Height
   BitBlt hDC, xDst, yDst, WidthSrc, HeightSrc, m_ThDC, xSrc, ySrc, vbSrcCopy

End Sub

'DESTROYS THE TEMP DC
Private Sub pDestroy()
    
    If Not m_hBmpOld = 0 Then
        SelectObject m_ThDC, m_hBmpOld
        m_hBmpOld = 0
    End If
    If Not m_hBmp = 0 Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If
    If Not m_ThDC = 0 Then
        DeleteDC m_ThDC
        m_ThDC = 0
    End If
    m_lWidth = 0
    m_lHeight = 0
    
End Sub

Private Function ChangeWindowStyles(ByVal ThWnd As Long) As Long
  Dim TFlat As Long
  
  '===================================================================================================================
  '//-- Experimenting with some window styles to create a Tooltip look window, this flags give me the best results.
  '===================================================================================================================
  
  TFlat = GetWindowLong(ThWnd, GWL_STYLE)    '//-- Retrieve information about the specified window.
  TFlat = TFlat And WS_BORDER And 1600000    '//-- Parameters of a window that has a flat border.
  SetWindowLongA ThWnd, GWL_STYLE, TFlat      '//-- Change attributes of the specified window.
  
  TFlat = GetWindowLong(ThWnd, GWL_EXSTYLE)  '//-- Retrieve information about the specified window.
  TFlat = TFlat And Not WS_EX_CLIENTEDGE     '//-- Specifies that the window doesn't have a 3D look ?that is, Flat Border Style.
  SetWindowLongA ThWnd, GWL_EXSTYLE, TFlat   '//-- Change attributes of the specified window.
      
  '//-- The SetWindowPos function changes the size, position, and Z order of a child, pop-up, or top-level window.
  '//-- Child, pop-up, and top-level windows are ordered according to their appearance on the screen.
  '//-- The topmost window receives the highest rank and is the first window in the Z order.
  
  With M_ExWindow
    ChangeWindowStyles = SetWindowPos(ThWnd, 0, .x - .bExt, .y + .aExt, .Width, .Height + Abs(.aExt), SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
  End With
  
  '//-- FYI: SWP_FRAMECHANGED Sends a WM_NCCALCSIZE message to the window
  '//-- FYI: SWP_NOZORDER Retains the current Z order.
  '//-- FYI: SWP_NOACTIVATE Does not activate the window.
  
End Function

Private Function InsideArea(cHandle As Long) As Boolean
    Dim POS As POINTAPI
        
    GetCursorPos POS

    If (WindowFromPoint(POS.x, POS.y) <> cHandle) Then
        InsideArea = False
        Else
        InsideArea = True
    End If

End Function


Private Function CaptureWindow(ByVal hWndSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Long
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim hDCSrc As Long
      
    hDCSrc = GetDC(hWndSrc) ' Get device context for entire window.
    
    ' Create a memory device context for the copy process.
    hDCMemory = CreateCompatibleDC(hDCSrc)
    ' Create a bitmap and place it in the memory DC.
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    
    ' Copy the on-screen image into the memory DC.
    BitBlt hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy
    
    ' Remove the new copy of the  on-screen image.
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    ' Release the device context resources back to the system.
    Call DeleteDC(hDCMemory)
    Call ReleaseDC(hWndSrc, hDCSrc)
    
    CaptureWindow = hBmp
   
End Function

Private Sub DrawGradient(lStartcolor As Long, lEndColor As Long, ByRef dRect As RECT, ByVal hDC As Long, Optional bH As Boolean)
    'Draw a Gradient in the current HDC
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    
    lEndColor = GetLngColor(lEndColor)
    lStartcolor = GetLngColor(lStartcolor)

    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    With dRect
    
        sR = (sR - eR) / IIf(bH, .right, .bottom)
        sG = (sG - eG) / IIf(bH, .right, .bottom)
        sB = (sB - eB) / IIf(bH, .right, .bottom)
    
        
        For ni = 0 To IIf(bH, .right, .bottom)
        
            If bH Then
                DrawLine .left + ni, .top, .left + ni, .bottom, hDC, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
            Else
                DrawLine .left, .top + ni, .right, .top + ni, hDC, RGB(eR + (ni * sR), eG + (ni * sG), eB + (ni * sB))
            End If
        
        Next ni
    
    End With
    
End Sub

Private Sub Class_Initialize()
    
    Set Font = New StdFont              '//-- Font Assigned by the current System Font
         
    mDelayTime = 100                    '//-- 1 Sec
    mKillTime = 500                     '//-- 5 Sec
    mBackStyle = 1                      '//-- Default BackStyle is 'Solid'
    mToolTipStyle = TTStandard          '//-- Default Tooltip Style is 'Standard'
    mShadow = False                     '//-- Default Shadow is OFF
    mIconSize = TTIcon16                '//-- Default Icon is 16x16
    mTextColor = vbButtonText           '//-- Default Color
    mBackColor = vbInfoBackground       '//-- Default Color
    mGCStart = vbInfoBackground         '//-- Default Color
    mGCEnd = vbDesktop                  '//-- Default Color
    
End Sub

Private Sub Class_Terminate()
    
    '必须销毁子类
    If Not mclsSubOwner Is Nothing Then '// Terminate all subclassing
        Set mclsSubOwner = Nothing
    End If
    If Not mclsSubTip Is Nothing Then
        Set mclsSubTip = Nothing
    End If
    
    pDestroy                            '//-- Destroy Temp DC
    DestroyWindow M_ExWindow.hwnd       '//-- Destroy API Window
    
End Sub

Private Function mclsSubOwner_OnMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean) As Long
 
    Select Case uMsg

    '==============================================================================================
    '==============================================================================================
    'The WM_MOUSEMOVE message is posted to a window when the cursor moves.
    'If the mouse is not captured, the message is posted to the window that contains the cursor.
    'Otherwise, the message is posted to the window that has captured the mouse.
    '==============================================================================================
    '==============================================================================================
    Case WM_MOUSEMOVE

       '//-- If the User is Over the Calling Control Turn 'ON' The Hover Flag.
        If InsideArea(LWnd) Then
            If M_ExWindow.Hover = False Then
                M_ExWindow.Hover = True         '/-- If The Mouse Enter The Control Area
                SetTimer LWnd, 1, 1, 0          '/-- Start API Timer
            End If
        Else
            M_ExWindow.Hover = False            '/-- If The Mouse Left The Control Area
        End If
    
    '==============================================================================================
    '==============================================================================================
    'The WM_TIMER message is posted to the installing thread's message queue when a timer expires.
    'You can process the message by providing a WM_TIMER case in the window procedure.
    'Otherwise, the default window procedure will call the TimerProc callback function specified
    'in the call to the SetTimer function used to install the timer.
    '==============================================================================================
    '==============================================================================================
   Case WM_TIMER
        If M_ExWindow.Hover And (Not M_ExWindow.Flag) Then
            
            CMOS.x = CPOS.x     '//-- Assign Memory X POS.
            CMOS.y = CPOS.y     '//-- Assign Memory Y POS.
            GetCursorPos CPOS   '//-- Capture New X-Y Mouse Coordinates.

            '//-- ToolTips Don't Popup Just When the mouse is moving,
            '     they do so until the mouse is "still" for a few moments.
            '     Lets do it until that's 1 second.
            If (CMOS.x = CPOS.x) And (CMOS.y = CPOS.y) Then
                bTime = bTime + 1
                If bTime > mDelayTime Then      '//-- At This Point The Mouse Cursor Has been still for 1 second in the same position
                    bTime = 0                   '//-- Lets Reset the Timer, To count the time that the Tooltip is going to be Alive.
                    If mAlive = True Then       '//-- Ensure no child controls are on same Mouse Coords(in case Lwnd has Child Controls)
                        CreateToolTip           '//-- Begin The ToolTip Drawing.
                        M_ExWindow.Flag = True  '//-- We Now Know That We are ready to draw the ToolTip.
                    End If
                    Exit Function
                End If
            Else
                bTime = 0                       '//-- At This Point, the Mouse Cursor has not been still for 1 second in the same position, let's reset the timer.
            End If

        End If

        If InsideArea(LWnd) = False Then
             KillTimer LWnd, 1                  '//-- Destroy the API Timer
             DestroyWindow M_ExWindow.hwnd      '//-- Destroy the ToolTip Window when the Mouse Leaves the Control
             M_ExWindow.hwnd = 0                '//-- Reset the HWnd    Value
             M_ExWindow.Hover = False           '//-- Reset the Hover   Value
             M_ExWindow.Flag = False            '//-- Reset the Flag    Value
             'Class_Initialize                   '//-- Reset Default Values for next Instance.
        ElseIf bTime >= mKillTime Then
             DestroyWindow M_ExWindow.hwnd      '//-- AT This Point, the Tool tip as past out the visible time, and must be Destroyed.
             'Class_Initialize                   '//-- Reset Default Values for next Instance.
        ElseIf M_ExWindow.Hover And M_ExWindow.hwnd <> 0 Then
             bTime = bTime + 1
        End If

    '==============================================================================================
    '==============================================================================================
    'The WM_MOUSEACTIVATE message is sent when the cursor is in an inactive window and the user
    'presses a mouse button.
    '==============================================================================================
    '==============================================================================================
    Case WM_MOUSEACTIVATE
        
        '//-- At This Point, the Tool tip has to be destroyed,
        '     because the user click on the calling control (natural Tooltip behavior).
                              
        KillTimer LWnd, 1                  '//-- Destroy the API Timer
        DestroyWindow M_ExWindow.hwnd      '//-- Send WM_DESTROY Message.
        M_ExWindow.Hover = False           '//-- Reset the Hover Value
        'Class_Initialize                   '//-- Reset Default Values for next Instance.
        mAlive = True                      '//-- Force User To Move Mouse Outside Control Before another Tooltip is Displayed.
                
    End Select
    
End Function

Private Function mclsSubTip_OnMessage(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean) As Long
    
    '==============================================================================================
    '==============================================================================================
    'The WM_DESTROY message is sent when a window is being destroyed. It is sent to the window procedure
    'of the window being destroyed after the window is removed from the screen.
    '==============================================================================================
    '==============================================================================================
    Select Case uMsg
    Case WM_DESTROY
        mAlive = False

    End Select
    
End Function
