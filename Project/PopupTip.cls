VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PopupTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MINMARGIN As Long = 8    ' ��С��Ե����

Private Const MK_LBUTTON As Long = &H1
Private Const MK_MBUTTON As Long = &H10
Private Const MK_RBUTTON As Long = &H2
Private Const MK_SHIFT As Long = &H4
Private Const MK_CONTROL As Long = &H8

Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Const WM_TIMER As Long = &H113

' GetSystemMetrics nIndex parameter constants
Private Const SM_CYCAPTION = 4        ' Height of window caption
Private Const SM_CXFULLSCREEN = 16    ' Width of window client area
Private Const SM_CYFULLSCREEN = 17    ' Height of window client area

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

' ������ʽ
Public Enum BackStyleConstants
    BakColor = 0
    Gradient = 1
    ImagePic = 2
End Enum

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SleepEx Lib "kernel32.dll" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Private WithEvents mfrmTip As frmPopupTip
Attribute mfrmTip.VB_VarHelpID = -1
Private mclsGdi As TGdi.WinGdi
Private mclsPlay As MediaPlayer

Private mclrBackColor As OLE_COLOR              ' ������ɫ���ı���ʾ������ɫ��
Private mclrForeColor As OLE_COLOR              ' ǰ����ɫ���������ɫ��
Private mclrBorderColor As OLE_COLOR            ' �߿���ɫ
Private mclrGradientStartColor As OLE_COLOR     ' ���俪ʼ��ɫ
Private mclrGradientEndColor As OLE_COLOR       ' ���������ɫ
Private mstrFontName As String                  ' ��������
Private mlngFontSize As Long                    ' �����С
Private mblnFontBold As Boolean                 ' �Ƿ����
Private mblnFontItalic As Boolean               ' �Ƿ�б��
Private mlngGradientAngle As Long               ' ����Ƕ�
Private mimgBackPicture As StdPicture           ' ����ͼƬ
Private menmBackStyle As BackStyleConstants     ' ������ʽ
Private mlngHeight As Long                      ' ����߶�
Private mlngWidth As Long                       ' �������

Private mlngTimerID As Long                     ' ��ʱ��ID
Private mlngAutoPopDelay As Long                ' �Զ�������ʱ(����Ϊ��λ)
Private mlngElapse As Long                      ' ����ʱ��(����Ϊ��λ)
Private mblnDelaying As Boolean                 ' ������ʱ��
Private mhCursorHand As StdPicture              ' ���ι��

' ������ɫ
Public Property Get BackColor() As OLE_COLOR
    BackColor = mclrBackColor
End Property
Public Property Let BackColor(ByVal clrValue As OLE_COLOR)
    mclrBackColor = clrValue
End Property

' ������ɫ
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mclrForeColor
End Property
Public Property Let ForeColor(ByVal clrValue As OLE_COLOR)
     mclrForeColor = clrValue
End Property

' �߿���ɫ
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mclrBorderColor
End Property
Public Property Let BorderColor(ByVal clrValue As OLE_COLOR)
    mclrBorderColor = clrValue
End Property

' ��������
Public Property Get FontName() As String
    FontName = mstrFontName
End Property
Public Property Let FontName(ByVal strValue As String)
    mstrFontName = strValue
End Property

' �����С
Public Property Get FontSize() As Long
    FontSize = mlngFontSize
End Property
Public Property Let FontSize(ByVal lngValue As Long)
    mlngFontSize = lngValue
End Property

' �����Ƿ�Ϊ����
Public Property Get FontBold() As Boolean
    FontBold = mblnFontBold
End Property
Public Property Let FontBold(ByVal blnValue As Boolean)
    mblnFontBold = blnValue
End Property

' �����Ƿ����»���
Public Property Get FontItalic() As Boolean
    FontItalic = mblnFontItalic
End Property
Public Property Let FontItalic(ByVal blnValue As Boolean)
    mblnFontItalic = blnValue
End Property

' ����߶�
Public Property Get Height() As Long
    Height = mlngHeight
End Property
Public Property Let Height(ByVal lngValue As Long)
    mlngHeight = lngValue
End Property

' �������
Public Property Get Width() As Long
    Width = mlngWidth
End Property
Public Property Let Width(ByVal lngValue As Long)
    mlngWidth = lngValue
End Property

' ������ʽ
Public Property Get BackgroundStyle() As BackStyleConstants
    BackgroundStyle = menmBackStyle
End Property
Public Property Let BackgroundStyle(ByVal enmValue As BackStyleConstants)
    menmBackStyle = enmValue
End Property

' ���ñ�������
Public Sub SetBackGradient(ByVal lngAngle As Long, ByVal lngStartColor As Long, ByVal lngEndColor As Long)
    mlngGradientAngle = lngAngle
    mclrGradientStartColor = lngStartColor
    mclrGradientEndColor = lngEndColor
End Sub

' ���ñ���ͼƬ
Public Sub SetBackPicture(picImage As StdPicture)
    If Not picImage Is Nothing Then
       Set mimgBackPicture = picImage
    End If
End Sub

' ��ʾ��Ϣ
Public Sub DisplayAlert( _
    ByVal MessageText As String, _
    Optional ByVal AutoPopDelay As Long = 5, _
    Optional ByVal PlayAudio As Boolean = True, _
    Optional ByVal ResourceSound As ResoureSoundConstants = Reminder, _
    Optional ByVal SystemSound As SystemSoundConstants = SystemDefault, _
    Optional ByVal PlayResourceSound As Boolean = True)
    Dim tRText As RECT
    
    ' ������ʱ�У�����������
    If mblnDelaying Then
        Exit Sub
    End If
    
    'ɾ���Ѿ����ɵ�Bitmap����ֹ�������ڴ�Сʱ������Bitmap��Դй©
    mclsGdi.DeleteObject mfrmTip.Image.handle
 
    ' ���ô��ڿ��Ⱥ͸߶�
    mfrmTip.Width = mlngWidth * Screen.TwipsPerPixelX
    mfrmTip.Height = mlngHeight * Screen.TwipsPerPixelY
    
    ' ��������
    mfrmTip.FontName = mstrFontName      '��������
    mfrmTip.FontSize = mlngFontSize      '�����С
    mfrmTip.FontBold = mblnFontBold      '�Ƿ����
    mfrmTip.FontItalic = mblnFontItalic  '�Ƿ�б��
    
    ' ���õ�����ʱ
    mlngAutoPopDelay = AutoPopDelay
    
    ' ���ñ��������߿�
    Select Case menmBackStyle
    Case BakColor
        mclsGdi.DrawGradient mfrmTip.hDC, 0, 0, mfrmTip.ScaleWidth, mfrmTip.ScaleHeight, _
        mclrBackColor, mclrBackColor, dHorizontal
    Case Gradient
        mclsGdi.DrawGradient mfrmTip.hDC, 0, 0, mfrmTip.ScaleWidth, mfrmTip.ScaleHeight, _
        mclrGradientStartColor, mclrGradientEndColor, mlngGradientAngle
    Case ImagePic
        mclsGdi.DrawImage mfrmTip.hDC, mimgBackPicture.handle, 0, 0, mfrmTip.ScaleWidth, mfrmTip.ScaleHeight, ilStretch
    End Select
    mclsGdi.DrawBorder mfrmTip.hDC, 0, 0, mfrmTip.ScaleWidth - 1, mfrmTip.ScaleHeight - 1, mclrBorderColor
    
    ' ���رհ�Ŧ
    'mclsGdi.DrawBorder mfrmTip.hdc, mfrmTip.ScaleWidth - 16, 3, 12, 12, mclrBorderColor
    mclsGdi.DrawClose mfrmTip.hDC, mfrmTip.ScaleWidth - 13, 6, 6, mclrBorderColor
    
    ' �����ı����򲢻����ı�
    SetRect tRText, MINMARGIN, MINMARGIN * 2, mfrmTip.ScaleWidth - MINMARGIN, mfrmTip.ScaleHeight - MINMARGIN
    mclsGdi.DrawText mfrmTip.hDC, MessageText, _
    tRText.left, tRText.top, tRText.right - tRText.left, _
    tRText.bottom - tRText.top, _
    mclrForeColor, dtWordBreak

    
    ' ˢ�����
    mfrmTip.Refresh
    
    ' ��������
    If PlayResourceSound Then
        mclsPlay.PlayResAudio ResourceSound
    Else
        mclsPlay.PlaySysAudio SystemSound
    End If
    
    ' ��ʾ��������
    ShowPopupWindow
    
End Sub

' ��ʾ��������
Private Sub ShowPopupWindow()
    Dim lngCurHeight As Long
    Dim lngNewHeight As Long
    
    ' ��ʾ����
    mlngElapse = 0
    lngNewHeight = 0
    lngCurHeight = mfrmTip.Height
    SetWindowPos mfrmTip.hWnd, HWND_TOPMOST, _
    (GetSystemMetrics(SM_CXFULLSCREEN) - mfrmTip.Width \ Screen.TwipsPerPixelX) - 6, _
    GetSystemMetrics(SM_CYFULLSCREEN) + GetSystemMetrics(SM_CYCAPTION), _
    mfrmTip.ScaleWidth, 0&, SWP_SHOWWINDOW Or SWP_NOACTIVATE
    mblnDelaying = True
    Do
       mfrmTip.top = mfrmTip.top - 40
       mfrmTip.Height = mfrmTip.Height + 40
       lngNewHeight = lngNewHeight + 40
       SleepEx 1, True
       DoEvents
    Loop Until lngNewHeight >= lngCurHeight
    mfrmTip.Height = lngCurHeight
    mblnDelaying = False
    
    ' ������ʱ��
    If Not mfrmTip.tmrControl.Enabled Then
        mfrmTip.tmrControl.Enabled = True
    End If
    
End Sub

' ���ص�������
Private Sub HidePopupWindow()
    Dim lngCurHeight As Long
    Dim lngNewHeight As Long
    
    ' ���ش���
    mlngElapse = 0
    lngCurHeight = mfrmTip.Height
    lngNewHeight = 0
    mblnDelaying = True
    Do
        mfrmTip.top = mfrmTip.top + 40
        mfrmTip.Height = mfrmTip.Height - 40
        lngNewHeight = lngNewHeight + 40
        SleepEx 1, True
        DoEvents
    Loop Until lngNewHeight >= lngCurHeight
    mfrmTip.Hide
    mfrmTip.Height = lngCurHeight
    mblnDelaying = False
    
End Sub

' ж�ص���ʽ��ʾ����
Private Sub UnloadPopupTipWindow()
    
    If mfrmTip.Visible Then
        mfrmTip.Hide
    End If
    Unload mfrmTip
    
End Sub


Private Sub mfrmTip_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub mfrmTip_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If x >= (mfrmTip.ScaleWidth - 15) _
        And x <= (mfrmTip.ScaleWidth - 5) _
        And y >= 3 And y < 15 Then
        mfrmTip.MousePointer = vbCustom
        Set mfrmTip.MouseIcon = mhCursorHand
    Else
        If mfrmTip.MousePointer <> vbDefault Then
            mfrmTip.MousePointer = vbDefault
        End If
    End If
    
End Sub

Private Sub mfrmTip_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    RaiseEvent MouseUp(Button, Shift, x, y)
    If x >= (mfrmTip.ScaleWidth - 15) And x <= (mfrmTip.ScaleWidth - 5) _
        And y >= 3 And y < 15 Then
        HidePopupWindow
    End If
    
End Sub

Private Sub mfrmTip_Tick()
    
    mlngElapse = mlngElapse + 1
    If mlngElapse = mlngAutoPopDelay Then
        HidePopupWindow
    End If
    
End Sub

Private Sub Class_Initialize()

    Set mclsGdi = New TGdi.WinGdi
    Set mclsPlay = New MediaPlayer
    Set mfrmTip = frmPopupTip
    Set mhCursorHand = LoadResPicture(101, vbResCursor)

    mclrBackColor = vbInfoBackground
    mclrForeColor = vbButtonText
    mclrBorderColor = vbWindowFrame
    mstrFontName = mfrmTip.FontName
    mlngFontSize = mfrmTip.FontSize
    mblnFontBold = False
    mblnFontItalic = False
    mlngGradientAngle = -100
    mclrGradientStartColor = RGB(0, 192, 255)
    mclrGradientEndColor = RGB(255, 255, 255)
    menmBackStyle = Gradient
    mlngHeight = 120
    mlngWidth = 160
    
End Sub

Private Sub Class_Terminate()
    
    UnloadPopupTipWindow
    
    If ObjPtr(mhCursorHand) > 0 Then
        mclsGdi.DeleteObject mhCursorHand.handle
    End If
    Set mhCursorHand = Nothing
    
    Set mclsPlay = Nothing
    Set mclsGdi = Nothing
    
End Sub

