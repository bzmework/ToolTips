VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SPACEAREA As Long = 4     ' 最小间隙
Private Const BARWIDTH As Long = 16     ' 垂直条最小宽度
Private Const SHADOWWIDTH As Long = 3   ' 阴影最小宽度

Public Enum ToolTipStyleConstants
    Custom = 1
    BorderBar = 2
End Enum

Public Enum ToolTipIconSizeConstants
    Icon16 = 0
    Icon24 = 1
    Icon32 = 2
    Icon48 = 3
    Icon64 = 4
    Icon72 = 5
    Icon96 = 6
    Icon128 = 7
End Enum

Public Enum ToolTipSysTypeConstants
    None = 0
    Warning = 1
    Error = 2
    Asterisk = 3 ' Information
    Question = 4
End Enum

Public Enum GradientDirectionConstants
    LeftToRight = 0
    TopToBottom = 1
End Enum

Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSEACTIVATE As Long = &H21
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_RBUTTONDBLCLK As Long = &H206
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MBUTTONDBLCLK As Long = &H209


Private Const WM_TIMER As Long = &H113

' GetClassLong/SetClassLong nIndex constants
Private Const GCL_MENUNAME As Long = -8
Private Const GCL_HBRBACKGROUND As Long = -10
Private Const GCL_HCURSOR As Long = -12
Private Const GCL_HICON As Long = -14
Private Const GCL_HMODULE As Long = -16
Private Const GCL_CBWNDEXTRA As Long = -18
Private Const GCL_CBCLSEXTRA As Long = -20
Private Const GCL_WNDPROC As Long = -24
Private Const GCL_STYLE As Long = -26
Private Const GCL_HICONSM As Long = -34
Private Const GCL_CONVERSION As Long = &H1
Private Const GCL_REVERSECONVERSION As Long = &H2
Private Const GCL_REVERSE_LENGTH As Long = &H3


' GetWindowLong/SetWindowLong Style constants
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE As Long = -16

Private Const CW_USEDEFAULT As Long = &H80000000

' CreateWindowEx dwExStyle parameter constants
Private Const WS_ACTIVECAPTION As Long = &H1
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_BORDER As Long = &H800000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_CHILDWINDOW As Long = (WS_CHILD)
Private Const WS_CLIPCHILDREN As Long = &H2000000
Private Const WS_CLIPSIBLINGS As Long = &H4000000
Private Const WS_DISABLED As Long = &H8000000

Private Const WS_EX_LEFT As Long = &H0&
Private Const WS_EX_LTRREADING As Long = &H0&
Private Const WS_EX_RIGHTSCROLLBAR As Long = &H0&
Private Const WS_EX_DLGMODALFRAME As Long = &H1&
Private Const WS_EX_NOPARENTNOTIFY As Long = &H4&
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_ACCEPTFILES As Long = &H10&
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WS_EX_MDICHILD As Long = &H40&
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_CONTEXTHELP As Long = &H400&
Private Const WS_EX_RIGHT As Long = &H1000&
Private Const WS_EX_RTLREADING As Long = &H2000&
Private Const WS_EX_LEFTSCROLLBAR As Long = &H4000&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_CONTROLPARENT As Long = &H10000
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_APPWINDOW As Long = &H40000
Private Const WS_EX_NOINHERITLAYOUT As Long = &H100000
Private Const WS_EX_LAYOUTRTL As Long = &H400000
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const WS_EX_OVERLAPPEDWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Private Const WS_EX_PALETTEWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

' SetWindowPos hWndInsertAfter parameter constants
Private Const HWND_MESSAGE As Long = -3
Private Const HWND_NOTOPMOST As Long = -2
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_DESKTOP As Long = 0
Private Const HWND_TOP As Long = 0
Private Const HWND_BOTTOM As Long = 1
Private Const HWND_BROADCAST As Long = &HFFFF&

' SetWindowPos wFlags parameter constants
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_NOCOPYBITS As Long = &H100
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_DEFERERASE As Long = &H2000
Private Const SWP_ASYNCWINDOWPOS As Long = &H4000
Private Const SWP_NOREPOSITION As Long = SWP_NOOWNERZORDER
Private Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED

' ShowWindow nCmdShow parameter constants
' GCW_ATOM: Retrieves an ATOM value that uniquely identifies the window class.
' This is the same atom that the RegisterClassEx function returns.
' GCL_CBCLSEXTRA: Retrieves the size, in bytes, of the extra memory associated with the class.
' GCL_CBWNDEXTRA: Retrieves the size, in bytes, of the extra window memory associated with each window in the class.
' For information on how to access this memory, see GetWindowLong.
' GCL_HBRBACKGROUND: Retrieves a handle to the background brush associated with the class.
' GCL_HCURSOR: Retrieves a handle to the cursor associated with the class.
' GCL_HICON: Retrieves a handle to the icon associated with the class.
' GCL_HICONSM: Retrieves a handle to the small icon associated with the class.
' GCL_HMODULE: Retrieves a handle to the module that registered the class.
' GCL_MENUNAME: Retrieves the address of the menu name string. The string identifies the menu resource associated with the class.
' GCL_STYLE: Retrieves the window-class style bits.
' GCL_WNDPROC: Retrieves the address of the window procedure, or a handle representing the address of the window procedure.
' You must use the CallWindowProc function to call the window procedure.
Private Const SW_HIDE As Long = 0
Private Const SW_NORMAL As Long = 1
Private Const SW_SHOWMINIMIZED As Long = 2
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWNOACTIVATE As Long = 4
Private Const SW_SHOW As Long = 5
Private Const SW_MINIMIZE As Long = 6
Private Const SW_SHOWMINNOACTIVE As Long = 7
Private Const SW_SHOWNA As Long = 8
Private Const SW_RESTORE As Long = 9
Private Const SW_MAX As Long = 10
Private Const SW_FORCEMINIMIZE As Long = 11


' Class Styles Constants
' CS_BYTEALIGNCLIENT: Aligns the window's client area on a byte boundary (in the x direction).
' This style affects the width of the window and its horizontal placement on the display.
' CS_BYTEALIGNWINDOW: Aligns the window on a byte boundary (in the x direction).
' This style affects the width of the window and its horizontal placement on the display.
' CS_CLASSDC: Allocates one device context to be shared by all windows in the class.
' Because window classes are process specific, it is possible for multiple threads of an application to create a window of the same class.
' It is also possible for the threads to attempt to use the device context simultaneously.
' When this happens, the system allows only one thread to successfully finish its drawing operation.
' CS_DBLCLKS: Sends a double-click message to the window procedure when the user double-clicks the mouse while the cursor is within a window belonging to the class.
' CS_DROPSHADOW: Windows XP: Enables the drop shadow effect on a window.
' The effect is turned on and off through SPI_SETDROPSHADOW.
' Typically, this is enabled for small, short-lived windows such as menus to emphasize their Z order relationship to other windows.
' CS_GLOBALCLASS: Specifies that the window class is an application global class.
' For more information, see Application Global Classes.
' CS_HREDRAW: Redraws the entire window if a movement or size adjustment changes the width of the client area.
' CS_NOCLOSE: Disables Close on the window menu.
' CS_OWNDC: Allocates a unique device context for each window in the class.
' CS_PARENTDC: Sets the clipping rectangle of the child window to that of the parent window
' so that the child can draw on the parent. A window with the CS_PARENTDC
' style bit receives a regular device context from the system's cache of device contexts.
' It does not give the child the parent's device context or device context settings.
' Specifying CS_PARENTDC enhances an application's performance.
' CS_SAVEBITS: Saves, as a bitmap, the portion of the screen image obscured by a window of this class.
' When the window is removed, the system uses the saved bitmap to restore the screen image,
' including other windows that were obscured. Therefore, the system does not send WM_PAINT messages to
' windows that were obscured if the memory used by the bitmap has not been discarded and if other screen actions have not invalidated the stored image.
' This style is useful for small windows (for example, menus or dialog boxes) that are displayed briefly and
' then removed before other screen activity takes place.
' This style increases the time required to display the window,
' because the system must first allocate memory to store the bitmap.
' CS_VREDRAW: Redraws the entire window if a movement or size adjustment changes the height of the client area.
Private Const CS_VREDRAW As Long = &H1
Private Const CS_HREDRAW As Long = &H2
Private Const CS_KEYCVTWINDOW As Long = &H4
Private Const CS_DBLCLKS As Long = &H8
Private Const CS_OWNDC As Long = &H20
Private Const CS_CLASSDC As Long = &H40
Private Const CS_PARENTDC As Long = &H80
Private Const CS_NOKEYCVT As Long = &H100
Private Const CS_NOCLOSE As Long = &H200
Private Const CS_SAVEBITS As Long = &H800
Private Const CS_BYTEALIGNCLIENT As Long = &H1000
Private Const CS_BYTEALIGNWINDOW As Long = &H2000
Private Const CS_GLOBALCLASS As Long = &H4000
Private Const CS_IME As Long = &H10000
Private Const CS_DROPSHADOW  As Long = &H20000

' CombineRgn fnCombineMode parameter constants
Private Const RGN_AND As Long = 1
Private Const RGN_OR As Long = 2
Private Const RGN_XOR As Long = 3
Private Const RGN_DIFF As Long = 4
Private Const RGN_COPY As Long = 5

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function UpdateWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function BeginPath Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function RoundRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function EndPath Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function PathToRegion Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private WithEvents mfrmTip As frmToolTip
Attribute mfrmTip.VB_VarHelpID = -1
Private mclsGdi As TGdi.WinGdi
Private mclsDrops As DropShadow

Private mblnIsXp As Boolean
Private mblnActive As Boolean
Private mstrTitle As String
Private mstrText As String
Private mlngTitleWidthMax As Long
Private mlngTextWidthMax As Long
Private mclrTitleColor As OLE_COLOR
Private mclrTitleBarColor As OLE_COLOR
Private mclrTextColor As OLE_COLOR
Private mclrBorderColor As OLE_COLOR
Private mclrBeginColor As OLE_COLOR
Private mclrEndColor As OLE_COLOR
Private mimgTitleIcon As StdPicture
Private menmTitleIconSize As ToolTipIconSizeConstants
Private mimgBackPicture As StdPicture
Private menmTipStyle As ToolTipStyleConstants
Private menmTipType As ToolTipSysTypeConstants
Private menmDirection As GradientDirectionConstants
Private mlngAutomaticDelay As Long
Private mlngInitialDelay As Long
Private mlngAutoPopDelay As Long
Private mlngAlpha As Long
Private mlngMargin As Long

Private mTipArea As RECT
Private mTipPos As POINTAPI
Private mblnMouseIn As Boolean
Private mhCurrCtlWnd As Long
Private mlngTimerID As Long
Private mlngWait As Long
Private mlngElapse As Long

' 工具提示窗口是否处于活动状态，只有处于活动状态才会显示提示
Public Property Get Active() As Boolean
    Active = mblnActive
End Property

' 标题文本
Public Property Get Title() As String
    Title = mstrTitle
End Property
Public Property Let Title(ByVal strValue As String)
    mstrTitle = strValue
End Property

' 内容文本
Public Property Get Text() As String
    Text = mstrText
End Property
Public Property Let Text(ByVal strValue As String)
    mstrText = strValue
End Property

' 标题字体颜色
Public Property Get TitleColor() As OLE_COLOR
    TitleColor = mclrTitleColor
End Property
Public Property Let TitleColor(ByVal clrColor As OLE_COLOR)
    mclrTitleColor = clrColor
End Property

' 标题条颜色
Public Property Get TitleBarColor() As OLE_COLOR
    TitleBarColor = mclrTitleBarColor
End Property
Public Property Let TitleBarColor(ByVal clrColor As OLE_COLOR)
    mclrTitleBarColor = clrColor
End Property

' 文本字体颜色
Public Property Get TextColor() As OLE_COLOR
    TextColor = mclrTextColor
End Property
Public Property Let TextColor(ByVal clrColor As OLE_COLOR)
    mclrTextColor = clrColor
End Property

' 边框颜色
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mclrBorderColor
End Property
Public Property Let BorderColor(ByVal clrColor As OLE_COLOR)
    mclrBorderColor = clrColor
End Property

' 渐变开始颜色
Public Property Get BeginColor() As OLE_COLOR
    BeginColor = mclrBeginColor
End Property
Public Property Let BeginColor(ByVal clrColor As OLE_COLOR)
    mclrBeginColor = clrColor
End Property

' 渐变结束颜色
Public Property Get EndColor() As OLE_COLOR
    EndColor = mclrEndColor
End Property
Public Property Let EndColor(ByVal clrColor As OLE_COLOR)
    mclrEndColor = clrColor
End Property

' 标题最大宽度
Public Property Get TitleWidthMax() As Long
    TitleWidthMax = mlngTitleWidthMax
End Property
Public Property Let TitleWidthMax(ByVal lngValue As Long)
    mlngTitleWidthMax = lngValue
End Property

' 内容最大宽度
Public Property Get TextWidthMax() As Long
    TextWidthMax = mlngTextWidthMax
End Property
Public Property Let TextWidthMax(ByVal lngValue As Long)
    mlngTextWidthMax = lngValue
End Property

' 标题图标
Public Property Get TitleIcon() As StdPicture
    Set TitleIcon = mimgTitleIcon
End Property
Public Property Set TitleIcon(ByVal imgPic As StdPicture)
    Set mimgTitleIcon = imgPic
End Property

' 标题图标尺寸
Public Property Get TitleIconSize() As ToolTipIconSizeConstants
    TitleIconSize = menmTitleIconSize
End Property
Public Property Let TitleIconSize(ByVal enmValue As ToolTipIconSizeConstants)
    menmTitleIconSize = enmValue
End Property

' 背景图片
Public Property Get BackPicture() As StdPicture
    Set BackPicture = mimgBackPicture
End Property
Public Property Set BackPicture(ByVal imgPic As StdPicture)
    Set mimgBackPicture = imgPic
End Property

' 工具提示样式
Public Property Get TipStyle() As ToolTipStyleConstants
    TipStyle = menmTipStyle
End Property
Public Property Let TipStyle(ByVal enmValue As ToolTipStyleConstants)
    menmTipStyle = enmValue
End Property

' 工具提示类型
Public Property Get TipType() As ToolTipSysTypeConstants
    TipType = menmTipType
End Property
Public Property Let TipType(ByVal enmValue As ToolTipSysTypeConstants)
    menmTipType = enmValue
End Property

' 渐变方向
Public Property Get Direction() As GradientDirectionConstants
    Direction = menmDirection
End Property
Public Property Let Direction(ByVal enmValue As GradientDirectionConstants)
    menmDirection = enmValue
End Property

' 自动延迟（以毫秒计）。默认值为 500。
' 该属性使您能够设置单个延迟值，该值转而用于设置 AutoPopDelay、InitialDelay和ReshowDelay三个属性的值。
' 每次设置AutomaticDelay属性时，默认情况下都会设置下列值：
' AutoPopDelay  AutomaticDelay属性值的10倍。
' InitialDelay  等于AutomaticDelay属性的值。
' ReshowDelay   AutomaticDelay属性值的1/5。
' 一旦设置了AutomaticDelay属性，也可以单独设置上述属性。请参见AutoPopDelay、InitialDelay和ReshowDelay三个属性。
Public Property Get AutomaticDelay() As Long
    AutomaticDelay = mlngAutomaticDelay
End Property
Public Property Let AutomaticDelay(ByVal lngValue As Long)
    mlngAutomaticDelay = lngValue
    mlngInitialDelay = mlngAutomaticDelay
    mlngAutoPopDelay = mlngAutomaticDelay * 10
End Property

' 获取或设置工具提示显示之前经过的时间。
' 此属性使您可以缩短或延长显示“工具提示”窗口前ToolTip等待的时间。
' 如果InitialDelay属性的值设置得太长，则应用程序的用户可能就不知道应用程序提供了工具提示帮助。
' 使用此属性，就可以确保用户通过缩短指定的时间快速显示工具提示。
Public Property Get InitialDelay() As Long
    InitialDelay = mlngInitialDelay
End Property
Public Property Let InitialDelay(ByVal lngValue As Long)
    mlngInitialDelay = lngValue
End Property

' 获取或设置当指针在包含指定工具提示文本的控件内保持静止时，工具提示保持的时段。
' 此属性使您可以缩短或延长鼠标指针位于控件上时ToolTip窗口的显示时间。
' 例如，如果在“工具提示”窗口中显示大量帮助信息，就可以增大此属性的值，以确保用户有足够的时间查阅此文本。
Public Property Get AutoPopDelay() As Long
    AutoPopDelay = mlngAutoPopDelay
End Property
Public Property Let AutoPopDelay(ByVal lngValue As Long)
    mlngAutoPopDelay = lngValue
End Property

' 透明程度
Public Property Get Alpha() As Long
    Alpha = mlngAlpha
End Property
Public Property Let Alpha(ByVal lngValue As Long)
    If lngValue < 0 Then lngValue = 0
    If lngValue > 255 Then lngValue = 255
    mlngAlpha = lngValue
End Property

' 边缘距离
Public Property Get Margin() As Long
    Margin = mlngMargin
End Property
Public Property Let Margin(ByVal lngValue As Long)
    mlngMargin = lngValue
End Property

' 显示工具提示
Public Sub Show(ByVal hWnd As Long)
    
    '已经进入则需要移出后再次移进后才能显示
    If mblnMouseIn Then
        Exit Sub
    End If
    
    '同一窗口不能反复显示
    If mhCurrCtlWnd = hWnd Then
        Exit Sub
    End If
    
    ' 不同风格对标题和内容的限制
    Select Case menmTipStyle
    Case Custom
        If Len(Trim$(mstrTitle)) = 0 And Len(Trim$(mstrText)) = 0 Then
            Exit Sub
        End If
    Case BorderBar
        If Len(Trim$(mstrText)) = 0 Then
            Exit Sub
        End If
    End Select
   
    ' 创建定时器
    If Not mfrmTip.tmrControl.Enabled Then
        mfrmTip.tmrControl.Enabled = True
    End If
    
    '初始化
    HideToolTipWindow
    mhCurrCtlWnd = hWnd
    mlngElapse = 0
    mlngWait = 0
    mblnActive = False
    mblnMouseIn = True
    
End Sub

' 创建透明阴影窗口
Private Sub CreateTransparentShadowWindow( _
    ByVal lngWndWidth As Long, ByVal lngWndHeight As Long, _
    Optional ByVal Alpha As Long = 255)
    Dim lngScrWidth As Long
    Dim lngScrHeight As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    
    ' 调整窗口高度和宽度，以使其能够显示阴影
    lngWndWidth = lngWndWidth + SHADOWWIDTH
    lngWndHeight = lngWndHeight + SHADOWWIDTH
    ' 获取屏幕宽度和高度
    lngScrWidth = Screen.Width / Screen.TwipsPerPixelX
    lngScrHeight = Screen.Height / Screen.TwipsPerPixelY
    ' 获得光标位置
    GetCursorPos mTipPos
    ' 计算正确显示位置
    If mTipPos.y <= lngScrHeight / 2 Then '上边区域
       If ((mTipPos.x < lngWndWidth And mTipPos.y < lngScrHeight)) Or _
          ((mTipPos.x < lngWndWidth)) Then '左上边
           lngLeft = mTipPos.x + 16
           lngTop = mTipPos.y + 16
       ElseIf ((lngScrWidth - mTipPos.x) < lngWndWidth And mTipPos.y < lngScrHeight) Or _
          ((lngScrWidth - mTipPos.x) < lngWndWidth) Then '右上边
           lngLeft = (mTipPos.x - lngWndWidth) + 16
           lngTop = mTipPos.y + 24
       Else
           lngLeft = mTipPos.x + 16
           lngTop = mTipPos.y + 16
       End If
    Else '下边区域
       If ((mTipPos.x < lngWndWidth And (lngScrHeight - mTipPos.y) < lngWndHeight)) Or _
          (mTipPos.x < lngWndWidth) Then '左下边
           lngLeft = mTipPos.x
           lngTop = mTipPos.y - lngWndHeight
       ElseIf ((lngScrWidth - mTipPos.x) < lngWndWidth And (lngScrHeight - mTipPos.y) < lngWndHeight) Or _
          ((lngScrWidth - mTipPos.x) < lngWndWidth) Then '右下边
           lngLeft = mTipPos.x - lngWndWidth
           lngTop = mTipPos.y - lngWndHeight
       Else
           lngLeft = mTipPos.x
           lngTop = mTipPos.y - lngWndHeight
       End If
    End If
    
    ' 隐藏提示窗口
    SetWindowPos mfrmTip.hWnd, HWND_TOPMOST, lngLeft, lngTop, lngWndWidth, lngWndHeight, _
    SWP_HIDEWINDOW Or SWP_NOACTIVATE
    
    ' 创建透明阴影窗体
'    mclsGdi.TransparentWindow mfrmTip.hDC, lngLeft, lngTop, lngWndWidth, lngWndHeight, Alpha
    mclsDrops.ShadowSize = SHADOWWIDTH
    mclsDrops.ShadowColor = vbButtonShadow
    mclsDrops.ShadowType = eRightShadow
    mclsDrops.DropShadow mfrmTip.hWnd, mfrmTip.hDC
    mclsDrops.ShadowType = eBottomShadow
    mclsDrops.DropShadow mfrmTip.hWnd, mfrmTip.hDC
    
End Sub

' 显示自定义工具提示窗口
Private Sub ShowCustomTipWindow()
    Dim lngTitleWidth As Long
    Dim lngTitleHeight As Long
    Dim lngTextWidth As Long
    Dim lngTextHeight As Long
    Dim lngWndWidth As Long
    Dim lngWndHeight As Long
    Dim lngTitleIconSize As Long
    Dim tRIcon As RECT
    Dim tRTitle As RECT
    Dim tRText As RECT
    
    '删除已经生成的Bitmap，防止调整窗口大小时，导致Bitmap资源泄漏
    mclsGdi.DeleteObject mfrmTip.Image.handle
    
    ' 计算标题宽度和高度
    mfrmTip.FontBold = True
    lngTitleWidth = mfrmTip.TextWidth(mstrTitle)
    lngTitleWidth = IIf(lngTitleWidth > mlngTitleWidthMax, mlngTitleWidthMax, lngTitleWidth)
    lngTitleHeight = mfrmTip.TextHeight(mstrTitle)
    lngTitleWidth = IIf(Len(Trim$(mstrTitle)) = 0, 0, lngTitleWidth)
    lngTitleHeight = IIf(Len(Trim$(mstrTitle)) = 0, 0, lngTitleHeight)
    
    ' 计算文本宽度和高度，文本不能为空
    mfrmTip.FontBold = False
    lngTextWidth = mfrmTip.TextWidth(mstrText)
    lngTextWidth = IIf(lngTextWidth > mlngTextWidthMax, mlngTextWidthMax, lngTextWidth)
    lngTextHeight = mfrmTip.TextHeight(mstrText)
    lngTextWidth = IIf(Len(Trim$(mstrText)) = 0, 0, lngTextWidth)
    lngTextHeight = IIf(Len(Trim$(mstrText)) = 0, 0, lngTextHeight)
    
    ' 计算图标高度
    Select Case menmTitleIconSize
    Case Icon16: lngTitleIconSize = 16
    Case Icon24: lngTitleIconSize = 24
    Case Icon32: lngTitleIconSize = 32
    Case Icon48: lngTitleIconSize = 48
    Case Icon64: lngTitleIconSize = 64
    Case Icon72: lngTitleIconSize = 72
    Case Icon96: lngTitleIconSize = 96
    Case Icon128: lngTitleIconSize = 128
    End Select
    If menmTipType = None Then
        lngTitleIconSize = 0
    Else
        If ObjPtr(mimgTitleIcon) = 0 Then
            lngTitleIconSize = 0
        End If
    End If
    
    ' 计算窗体宽度和高度
    If (lngTitleWidth + lngTitleIconSize + mlngMargin * 2 + SPACEAREA) < (lngTextWidth + mlngMargin * 2) Then
        lngWndWidth = lngTextWidth + mlngMargin * 2
    Else
        lngWndWidth = lngTitleWidth + lngTitleIconSize + mlngMargin * 2 + _
        IIf(lngTitleWidth = 0 Or lngTextWidth = 0, 0, SPACEAREA)
    End If
    If lngTitleHeight < lngTitleIconSize Then
        lngWndHeight = lngTitleIconSize + lngTextHeight + mlngMargin * 2 + SPACEAREA
    Else
        lngWndHeight = lngTitleHeight + lngTextHeight + mlngMargin * 2 + _
        IIf(lngTitleHeight = 0 Or lngTextHeight = 0, 0, SPACEAREA)
    End If
    
    ' 计算图标区域
    tRIcon.left = mlngMargin
    tRIcon.top = mlngMargin
    tRIcon.right = tRIcon.left + lngTitleIconSize
    tRIcon.bottom = tRIcon.top + lngTitleIconSize
    
    ' 计算标题区域
    tRTitle.left = IIf(lngTitleIconSize > 0, tRIcon.right + SPACEAREA, mlngMargin)
    tRTitle.top = tRIcon.top
    tRTitle.right = tRTitle.left + lngTitleWidth
    tRTitle.bottom = tRTitle.top + IIf(lngTitleHeight < lngTitleIconSize, lngTitleIconSize, lngTitleHeight)
    
    ' 计算文本区域
    tRText.left = mlngMargin
    tRText.top = IIf(lngTitleHeight > 0, tRTitle.bottom + SPACEAREA, mlngMargin)
    tRText.right = tRText.left + lngTextWidth
    tRText.bottom = tRText.top + lngTextHeight
    
    ' 填充背景并画边框
    If ObjPtr(mimgBackPicture) > 0 Then
        mclsGdi.DrawImage mfrmTip.hDC, mimgBackPicture.handle, 0, 0, lngWndWidth - 1, lngWndHeight - 1, ilStretch
        mclsGdi.DrawBorder mfrmTip.hDC, 0, 0, lngWndWidth - 1, lngWndHeight - 1, mclrBorderColor
    Else
        mclsGdi.DrawGradient mfrmTip.hDC, 0, 0, lngWndWidth - 1, lngWndHeight - 1, _
        mclrBeginColor, mclrEndColor, menmDirection
        mclsGdi.DrawBorder mfrmTip.hDC, 0, 0, lngWndWidth - 1, lngWndHeight - 1, mclrBorderColor
    End If
    
    ' 创建阴影和透明窗体
    CreateTransparentShadowWindow lngWndWidth, lngWndHeight, Alpha
    
    ' 画图标
    Select Case TipType
    Case Warning
        mclsGdi.DrawResIcon mfrmTip.hDC, mlngMargin, mlngMargin, lngTitleIconSize, lngTitleIconSize, idWarning
    Case Error
        mclsGdi.DrawResIcon mfrmTip.hDC, mlngMargin, mlngMargin, lngTitleIconSize, lngTitleIconSize, idError
    Case Asterisk
        mclsGdi.DrawResIcon mfrmTip.hDC, mlngMargin, mlngMargin, lngTitleIconSize, lngTitleIconSize, idInformation
    Case Question
        mclsGdi.DrawResIcon mfrmTip.hDC, mlngMargin, mlngMargin, lngTitleIconSize, lngTitleIconSize, idQuestion
    Case Else
        If ObjPtr(mimgTitleIcon) Then
            mclsGdi.DrawPic mfrmTip.hDC, mimgTitleIcon, mlngMargin, mlngMargin, lngTitleIconSize, lngTitleIconSize
        End If
    End Select
    
    ' 画标题
    mfrmTip.FontBold = True
    mclsGdi.DrawText mfrmTip.hDC, mstrTitle, tRTitle.left, tRTitle.top, _
    tRTitle.right - tRTitle.left, tRTitle.bottom - tRTitle.top, mclrTitleColor, _
    dtSingleLine Or dtLeft Or dtVCenter Or dtEndEllipsis

    ' 画文本
    mfrmTip.FontBold = False
    mclsGdi.DrawText mfrmTip.hDC, mstrText, tRText.left, tRText.top, _
    tRText.right - tRText.left, tRText.bottom - tRText.top, _
    mclrTextColor, dtWordEllipsis
    
    ' 刷新输出
    mfrmTip.Refresh
    
    ' 显示提示窗口
    ShowToolTipWindow
    
End Sub

' 显示边框条工具提示窗口
Private Sub ShowBorderBarTipWindow()
    Dim tRBar As RECT
    Dim tRText As RECT
    Dim lngTextWidth As Long
    Dim lngTextHeight As Long
    Dim lngWndWidth As Long
    Dim lngWndHeight As Long
    Dim lngScrWidth As Long
    Dim lngScrHeight As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim bytDirection As Byte
    
    '删除已经生成的Bitmap，防止调整窗口大小时，导致Bitmap资源泄漏
    mclsGdi.DeleteObject mfrmTip.Image.handle
    
    ' 计算文本宽度和高度，文本不能为空
    mfrmTip.FontBold = False
    lngTextWidth = mfrmTip.TextWidth(mstrText)
    lngTextHeight = mfrmTip.TextHeight(mstrText)
    
    ' 计算窗体宽度和高度
    lngWndWidth = BARWIDTH + lngTextWidth + SPACEAREA * 2
    lngWndHeight = lngTextHeight + SPACEAREA * 2
    
    ' 获取屏幕宽度和高度
    lngScrWidth = Screen.Width / Screen.TwipsPerPixelX
    lngScrHeight = Screen.Height / Screen.TwipsPerPixelY
    
    ' 获得光标位置
    GetCursorPos mTipPos
    
    ' 计算正确显示位置、计算文本区域、垂直条区域
    If mTipPos.y <= lngScrHeight / 2 Then '上边区域
        If ((mTipPos.x < lngWndWidth And mTipPos.y < lngScrHeight)) Or _
           ((mTipPos.x < lngWndWidth)) Then '左上边
            lngLeft = mTipPos.x + 16
            lngTop = mTipPos.y + 16
            tRText.left = BARWIDTH + SPACEAREA
            tRText.top = SPACEAREA
            tRText.right = tRText.left + lngTextWidth
            tRText.bottom = tRText.top + lngTextHeight
            tRBar.left = 0
            tRBar.top = 0
            tRBar.right = BARWIDTH
            tRBar.bottom = lngWndHeight
            bytDirection = 1
        ElseIf ((lngScrWidth - mTipPos.x) < lngWndWidth And mTipPos.y < lngScrHeight) Or _
           ((lngScrWidth - mTipPos.x) < lngWndWidth) Then '右上边
            lngLeft = (mTipPos.x - lngWndWidth) + 16
            lngTop = mTipPos.y + 24
            tRText.left = SPACEAREA
            tRText.top = SPACEAREA
            tRText.right = tRText.left + lngTextWidth
            tRText.bottom = tRText.top + lngTextHeight
            tRBar.left = lngWndWidth - BARWIDTH
            tRBar.top = 0
            tRBar.right = lngWndWidth
            tRBar.bottom = lngWndHeight
            bytDirection = 2
        Else
            lngLeft = mTipPos.x + 16
            lngTop = mTipPos.y + 16
            tRText.left = BARWIDTH + SPACEAREA
            tRText.top = SPACEAREA
            tRText.right = tRText.left + lngTextWidth
            tRText.bottom = tRText.top + lngTextHeight
            tRBar.left = 0
            tRBar.top = 0
            tRBar.right = BARWIDTH
            tRBar.bottom = lngWndHeight
            bytDirection = 1
        End If
    Else '下边区域
        If ((mTipPos.x < lngWndWidth And (lngScrHeight - mTipPos.y) < lngWndHeight)) Or _
          (mTipPos.x < lngWndWidth) Then '左下边
            lngLeft = mTipPos.x
            lngTop = mTipPos.y - lngWndHeight - SHADOWWIDTH
            tRText.left = BARWIDTH + SPACEAREA
            tRText.top = SPACEAREA
            tRText.right = tRText.left + lngTextWidth
            tRText.bottom = tRText.top + lngTextHeight
            tRBar.left = 0
            tRBar.top = 0
            tRBar.right = BARWIDTH
            tRBar.bottom = lngWndHeight
            bytDirection = 3
        ElseIf ((lngScrWidth - mTipPos.x) < lngWndWidth And (lngScrHeight - mTipPos.y) < lngWndHeight) Or _
          ((lngScrWidth - mTipPos.x) < lngWndWidth) Then '右下边
            lngLeft = mTipPos.x - lngWndWidth
            lngTop = mTipPos.y - lngWndHeight - SHADOWWIDTH
            tRText.left = SPACEAREA
            tRText.top = SPACEAREA
            tRText.right = tRText.left + lngTextWidth
            tRText.bottom = tRText.top + lngTextHeight
            tRBar.left = lngWndWidth - BARWIDTH
            tRBar.top = 0
            tRBar.right = lngWndWidth
            tRBar.bottom = lngWndHeight
            bytDirection = 4
        Else
            lngLeft = mTipPos.x
            lngTop = mTipPos.y - lngWndHeight - SHADOWWIDTH
            tRText.left = BARWIDTH + SPACEAREA
            tRText.top = SPACEAREA
            tRText.right = tRText.left + lngTextWidth
            tRText.bottom = tRText.top + lngTextHeight
            tRBar.left = 0
            tRBar.top = 0
            tRBar.right = BARWIDTH
            tRBar.bottom = lngWndHeight
            bytDirection = 3
        End If
    End If
    
    ' 填充背景并画边框
    mclsGdi.DrawGradient mfrmTip.hDC, 0, 0, lngWndWidth - 1, lngWndHeight - 1, mclrBeginColor, mclrEndColor, menmDirection
    mclsGdi.DrawGradient mfrmTip.hDC, tRBar.left, tRBar.top, tRBar.right - tRBar.left, tRBar.bottom - tRBar.top, mclrTitleBarColor, mclrTitleBarColor, menmDirection
    mclsGdi.DrawBorder mfrmTip.hDC, 0, 0, lngWndWidth - 1, lngWndHeight - 1, mclrBorderColor

    ' 画方向箭头
    Select Case bytDirection
    Case 1 ' 左上边
        'mclsGdi.DrawIcon mfrmTip.hdc, mhIconArrowWN, tRBar.Left + 3, tRBar.TOP + 2, 8, 8
        mclsGdi.DrawArrow mfrmTip.hDC, tRBar.left + 3, tRBar.top + 3, 8, 8, adWestNorth
    Case 2 ' 右上边
        'mclsGdi.DrawIcon mfrmTip.hdc, mhIconArrowEN, tRBar.Right - 11, tRBar.TOP + 2, 7, 7
        mclsGdi.DrawArrow mfrmTip.hDC, tRBar.right - 11, tRBar.top + 3, 7, 7, adEastNorth
    Case 3 ' 左下边
        'mclsGdi.DrawIcon mfrmTip.hdc, mhIconArrowWS, tRBar.Left + 3, tRBar.Bottom - 11, 7, 7
        mclsGdi.DrawArrow mfrmTip.hDC, tRBar.left + 3, tRBar.bottom - 11, 7, 7, adWestSouth
    Case 4 ' 右下边
        'mclsGdi.DrawIcon mfrmTip.hdc, mhIconArrowES, tRBar.Right - 11, tRBar.Bottom - 11, 7, 7
        mclsGdi.DrawArrow mfrmTip.hDC, tRBar.right - 11, tRBar.bottom - 11, 7, 7, adEastSouth
    End Select
    
    ' 画文本
    mclsGdi.DrawText mfrmTip.hDC, mstrText, tRText.left, tRText.top, _
    tRText.right - tRText.left, tRText.bottom - tRText.top, _
    mclrTextColor, dtWordEllipsis

    ' 调整窗口高度和宽度，以使其能够显示阴影
    lngWndWidth = lngWndWidth + SHADOWWIDTH
    lngWndHeight = lngWndHeight + SHADOWWIDTH
    
    ' 隐藏提示窗口
    SetWindowPos mfrmTip.hWnd, HWND_DESKTOP, lngLeft, lngTop, lngWndWidth, lngWndHeight, _
    SWP_HIDEWINDOW Or SWP_NOACTIVATE
    
    ' 创建透明阴影窗体
    mclsGdi.TransparentPaintAlpha mfrmTip.hDC, lngLeft, lngTop, lngWndWidth, lngWndHeight, 0, 0, 0, 0, 0, Alpha
    mclsDrops.ShadowSize = SHADOWWIDTH
    mclsDrops.ShadowColor = vbButtonShadow
    mclsDrops.ShadowType = eRightShadow
    mclsDrops.DropShadow mfrmTip.hWnd, mfrmTip.hDC
    mclsDrops.ShadowType = eBottomShadow
    mclsDrops.DropShadow mfrmTip.hWnd, mfrmTip.hDC
    
    ' 刷新输出
    mfrmTip.Refresh
    
    ' 显示提示窗口
    ShowToolTipWindow
    
End Sub

' 显示工具提示窗口
Private Sub ShowToolTipWindow()
    
    'SetCapture mfrmTip.hwnd
    ShowWindow mfrmTip.hWnd, SW_SHOWNOACTIVATE
    
End Sub

' 隐藏工具提示窗口
Private Sub HideToolTipWindow()

    mblnActive = False
    mhCurrCtlWnd = 0
    mlngWait = 0
    mlngElapse = 0
    If mfrmTip.Visible Then
        'ReleaseCapture
        mfrmTip.Hide
    End If
    
End Sub

' 卸载工具提示窗口
Private Sub UnloadToolTipWindow()

    HideToolTipWindow
    Unload mfrmTip
    Set mfrmTip = Nothing
    
End Sub

' 设置阴影窗口
Private Function SetShadowWindow() As Boolean
    Dim lngStyle As Long
    
    SetShadowWindow = False
    If mclsGdi.IsWinXpOrLater Then
        lngStyle = GetClassLong(mfrmTip.hWnd, GCL_STYLE)
        If (lngStyle And CS_DROPSHADOW) = CS_DROPSHADOW Then
            SetShadowWindow = True
        Else
            lngStyle = SetClassLong(mfrmTip.hWnd, GCL_STYLE, lngStyle Or CS_DROPSHADOW)
            SetShadowWindow = (lngStyle <> 0)
        End If
    End If
    
End Function

Private Sub Class_Initialize()
    
    Set mclsGdi = New TGdi.WinGdi
    Set mclsDrops = New DropShadow
    Set mfrmTip = frmToolTip

    mblnIsXp = mclsGdi.IsWinXpOrLater
    mclrTitleBarColor = RGB(255, 210, 83)
    mclrTitleColor = vbButtonText
    mclrTextColor = vbButtonText
    mclrBorderColor = vbWindowFrame
    mclrBeginColor = vbInfoBackground
    mclrEndColor = vbInfoBackground
    Set mimgTitleIcon = Nothing
    Set mimgBackPicture = Nothing
    mlngTitleWidthMax = 256
    mlngTextWidthMax = 512
    menmTitleIconSize = Icon16
    menmTipStyle = Custom
    menmDirection = TopToBottom
    mlngAutomaticDelay = 500
    mlngAutoPopDelay = 5000
    mlngInitialDelay = 500
    mlngElapse = 0
    mlngAlpha = 255
    mlngMargin = 6
    mblnActive = False
    
End Sub

Private Sub Class_Terminate()

    UnloadToolTipWindow
    
    If ObjPtr(mimgTitleIcon) > 0 Then
        mclsGdi.DeleteObject mimgTitleIcon.handle
    End If
    If ObjPtr(mimgBackPicture) > 0 Then
        mclsGdi.DeleteObject mimgBackPicture.handle
    End If
    Set mimgTitleIcon = Nothing
    Set mimgBackPicture = Nothing

    Set mclsGdi = Nothing
    Set mclsDrops = Nothing
    
End Sub

Private Sub mfrmTip_Click()
    If mblnActive Then
        HideToolTipWindow
    End If
End Sub

Private Sub mfrmTip_DblClick()
    If mblnActive Then
        HideToolTipWindow
    End If
End Sub

Private Sub mfrmTip_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mblnActive Then
        HideToolTipWindow
    End If
End Sub

'Private Sub mfrmTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    If GetCapture <> mfrmTip.hWnd Then
'        SetCapture mfrmTip.hWnd
'    End If
'
'End Sub
 
Private Sub mfrmTip_Tick()
    Dim tPt As POINTAPI
    Dim hWnd As Long
    
    GetCursorPos tPt
    hWnd = WindowFromPoint(tPt.x, tPt.y)
    If hWnd = mhCurrCtlWnd Then  '在当前窗口区域
        If Not mfrmTip.Visible Then
            '显示
            mlngWait = mlngWait + 1
            If mlngWait = (mlngInitialDelay \ 100) Then
                Select Case menmTipStyle
                Case Custom: ShowCustomTipWindow
                Case BorderBar: ShowBorderBarTipWindow
                End Select
                mblnActive = True
            End If
        Else
            '时间到
            mlngElapse = mlngElapse + 1
            If mlngElapse = (mlngAutoPopDelay \ 100) Then
                HideToolTipWindow
            End If
        End If
    Else '未在当前窗口区域
        If mhCurrCtlWnd <> 0 Then
            HideToolTipWindow
        End If
        mblnMouseIn = False
    End If
    
End Sub
