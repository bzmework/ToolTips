VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediaPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************
'
'   名称: MediaPlayer.cls
'   日期: 2005-09-03
'   作者: Qf
'   描述: 多媒体文件播放控制类,该类需要引入"ActiveMovie control type library(QUARTZ.DLL)"
'
'*********************************************************************************************
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'flag values for Flags parameter
'
' SND_APPLICATION: 用应用程序指定的关联来播放声音
' SND_ALIAS: pszSound参数指定了注册表或WIN.INI中的系统事件的别名
' SND_ALIAS_ID: pszSound参数指定了预定义的声音标识符
' SND_ASYNC: 用异步方式播放声音，PlaySound函数在开始播放后立即返回。
' SND_FILENAME: pszSound参数指定了WAVE文件名
' SND_LOOP: 重复播放声音，必须与SND_ASYNC标志一块使用。
' SND_MEMORY: 播放载入到内存中的声音，此时pszSound是指向声音数据的指针。
' SND_NODEFAULT: 不播放缺省声音，若无此标志，则PlaySound在没找到声音时会播放缺省声音。
' SND_NOSTOP: PlaySound不打断原来的声音播出并立即返回FALSE?
' SND_NOWAIT: 如果驱动程序正忙则函数就不播放声音并立即返回?
' SND_PURGE: 停止所有与调用任务有关的声音。若参数pszSound为NULL，就停止所有的声音，否则，停止pszSound指定的声音。
' SND_RESOURCE: pszSound参数是WAVE资源的标识符，这时要用到hmod参数。
' SND_SYNC: 同步播放声音，在播放完后PlaySound函数才返回。
Private Const SND_SYNC = &H0              '  play synchronously (default)
Private Const SND_ASYNC = &H1             '  play asynchronously
Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Private Const SND_MEMORY = &H4            '  lpszSoundName points to a memory file
Private Const SND_ALIAS = &H10000         '  name is a WIN.INI [sounds] entry
Private Const SND_FILENAME = &H20000      '  name is a file name
Private Const SND_RESOURCE = &H40004      '  name is a resource name or atom
Private Const SND_ALIAS_ID = &H110000     '  name is a WIN.INI [sounds] entry identifier
Private Const SND_ALIAS_START = 0         '  must be > 4096 to keep strings in same section of resource file
Private Const SND_LOOP = &H8              '  loop the sound until next sndPlaySound
Private Const SND_NOSTOP = &H10           '  don't stop any currently playing sound
Private Const SND_VALID = &H1F            '  valid flags          / ;Internal /
Private Const SND_NOWAIT = &H2000         '  don't wait if the driver is busy
Private Const SND_VALIDFLAGS = &H17201F   '  Set of valid flag bits.  Anything outside this range will raise an error
Private Const SND_RESERVED = &HFF000000   '  In particular these flags are reserved
Private Const SND_TYPE_MASK = &H170007

'waveform audio error return values
'
Private Const WAVERR_BASE = 32
Private Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)       '  unsupported wave format
Private Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '  still something playing
Private Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)      '  header not prepared
Private Const WAVERR_SYNC = (WAVERR_BASE + 3)            '  device is synchronous
Private Const WAVERR_LASTERROR = (WAVERR_BASE + 3)       '  last error in range

' 声音类型
Public Enum ResoureSoundConstants
    Online = 101
    Reminder = 102
    Request = 103
    Send = 104
    SendOk = 105
    TrashDropItem = 106
End Enum

Public Enum SystemSoundConstants
    SystemDefault = 101
    SystemAsterisk = 102
    SystemExclamation = 103
    SystemQuestion = 104
    SystemError = 105
    SystemStart = 106
    SystemExit = 107
End Enum

Public Enum PlayStatusConstants
    eStop = 0
    ePlay = 1
    ePause = 2
    eRePlay = 3
    eNo = 4
End Enum

Public Enum MediaTypeConstants
    eAudio = 0
    eVideo = 1
    eUnkonwn = 2
End Enum

Private Type RECT
    Left As Long
    TOP As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private m_FilgrpManager As FilgraphManager
Private m_VideoWnd As IVideoWindow
Private m_MediaPos As IMediaPosition
Private m_Audio As IBasicAudio
Private m_Video As IBasicVideo

Private m_OwnerhWnd As Long
Private m_VideohWnd As Long
Private m_OpenMediaSuccess As Boolean
Private m_VideoWidth As Long
Private m_VideoHeight As Long
Private m_MediaFileType As MediaTypeConstants
Private m_PlayStatus As PlayStatusConstants
Private m_Sound() As Byte

Private Sub Class_Initialize()
    
    Set m_FilgrpManager = New FilgraphManager
    Set m_VideoWnd = m_FilgrpManager
    Set m_Video = m_FilgrpManager
    Set m_MediaPos = m_FilgrpManager
    Set m_Audio = m_FilgrpManager
    
End Sub

Private Sub Class_Terminate()
    
    If m_OpenMediaSuccess = True Then m_FilgrpManager.Stop
    Set m_Audio = Nothing
    Set m_Video = Nothing
    Set m_MediaPos = Nothing
    Set m_VideoWnd = Nothing
    Set m_FilgrpManager = Nothing
    
End Sub

'//
'// 播放Windows系统WAVE音频
'//
Public Function PlaySysAudio(ByVal SoundID As SystemSoundConstants) As Long
    On Error Resume Next
    Dim strSoundAlias As String
    
    Select Case SoundID
    Case SystemAsterisk
        strSoundAlias = "SystemAsterisk"
    Case SystemExclamation
        strSoundAlias = "SystemExclamation"
    Case SystemExit
        strSoundAlias = "SystemExit"
    Case SystemError
        strSoundAlias = "SystemHand"
    Case SystemQuestion
        strSoundAlias = "SystemQuestion"
    Case SystemStart
        strSoundAlias = "SystemStart"
    Case Else
        strSoundAlias = "Gobbledygook"
    End Select
    PlaySysAudio = PlaySound(strSoundAlias, 0&, SND_ALIAS Or SND_ASYNC)

End Function

'//
'// 播放资源中的WAVE音频
'//
Public Function PlayResAudio(ByVal SoundID As ResoureSoundConstants) As Long
    On Error Resume Next
    Dim lngFlags As Long
    
    lngFlags = SND_MEMORY Or SND_NODEFAULT Or SND_ASYNC
    m_Sound = LoadResData(SoundID, "CUSTOM")
    PlayResAudio = PlaySoundData(m_Sound(0), 0&, lngFlags)

End Function

'//
'// 播放WAVE波形文件
'// 例如:PlayAudio "C:\MyAudio.WAV"
'//
Public Function PlayAudio(ByVal strWaveFile As String, _
    Optional ByVal Async As Boolean = True) As Long
    On Error Resume Next
    Dim lngFlags As Long
    
    lngFlags = SND_MEMORY Or SND_NODEFAULT
    If Async Then lngFlags = lngFlags Or SND_ASYNC
    PlayAudio = PlaySoundData(ByVal strWaveFile, 0, lngFlags)
   
End Function

'//
'// 函数:打开多媒体视频文件
'// 说明:对于视频文件则必须指定容纳视频图像的窗口句柄;对于音频文件则不用.
'// 示例:OpenMediaFile "C:\MyVideo.AVI",Me.hWnd
'//
Public Function OpenMediaFile(ByVal strFileName As String, Optional ByVal hWndOwner As Long) As Boolean
    On Error GoTo ErrorHandler
    Dim strTemp As String
    Dim lpRect As RECT
    
    ' 如果已经有多媒体文件被打开,则必须先关闭
    If m_OpenMediaSuccess = True Then
        m_MediaPos.CurrentPosition = 0
        m_FilgrpManager.Stop
    End If
    
    strTemp = UCase$(Right$(strFileName, 3))
    If strTemp = "AVI" Or _
        strTemp = "ASF" Or _
        strTemp = "WMV" Or _
        strTemp = "MLV" Or _
        strTemp = "MOV" Or _
        strTemp = "MPG" Or _
        strTemp = "DAT" Or _
        strTemp = "QT" Then '视频文件
       
        ' 打开视频文件
        m_FilgrpManager.RenderFile strFileName
        ' 必须设置窗口样式为0
        m_VideoWnd.WindowStyle = 0
        ' 指定容纳视频图像的目标窗口句柄
        m_VideoWnd.Owner = hWndOwner
        ' 获得目标窗口的客户区域大小
        GetWindowRect hWndOwner, lpRect
        ' 将视频图像填满整个窗口
        m_VideoWnd.SetWindowPosition 0, 0, lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.TOP
        ' 保存目标窗口句柄,在设置视频的回放,改变视频大小时要用到
        m_OwnerhWnd = hWndOwner
        ' 标记当前打开的文件类型为视频
        m_MediaFileType = eVideo
        ' 标记当前打开多媒体文件成功
        m_OpenMediaSuccess = True
        ' 获得视频窗口句柄
        m_VideohWnd = GetVideoWindowHandle(m_VideoWnd.Owner)
        ' 返回
        OpenMediaFile = True
    ElseIf strTemp = "MP3" Or _
        strTemp = "MP2" Or _
        strTemp = "MP1" Or _
        strTemp = "WAV" Or _
        strTemp = "MID" Then  '音频文件
        
        ' 打开音频文件
        m_FilgrpManager.RenderFile strFileName
        ' 标记当前打开的文件类型为音频
        m_MediaFileType = eAudio
        ' 标记当前打开多媒体文件成功
        m_OpenMediaSuccess = True
        ' 返回
        OpenMediaFile = True
    Else
        ' 媒体文件类型不能判定
        m_MediaFileType = eUnkonwn
        ' 标记当前打开多媒体文件失败
        m_OpenMediaSuccess = False
        ' 返回
        OpenMediaFile = False
    End If
    
    Exit Function
ErrorHandler:
    
    m_MediaFileType = eUnkonwn
    m_OpenMediaSuccess = False
    OpenMediaFile = False
    
End Function

'//
'// 获得视频窗口的句柄
'// 返回：成功返回视频窗口句柄，失败返回0
'// 例如：m_VideohWnd = GetVideoWindowHandle(hWndOwner)
'//
Private Function GetVideoWindowHandle(ByVal hWndOwner As Long) As Long
    Dim lnghWnd As Long

    lnghWnd = FindWindowEx(hWndOwner, ByVal 0&, "VideoRenderer", vbNullString)
    'lnghWnd = FindWindow(vbNullString, "ActiveMovie Window")
    GetVideoWindowHandle = lnghWnd
    
End Function

'//
'// 获得视频窗口句柄
'//
Public Property Get hwnd() As Long
    
    hwnd = m_VideohWnd
    
End Property

'//
'// 播放多媒体文件
'//
Public Sub PlayMedia()
    
    If m_OpenMediaSuccess = True Then
       
       m_FilgrpManager.Run
       m_PlayStatus = ePlay
       
    End If
    
End Sub

'//
'// 重新播放多媒体文件
'//
Public Sub RePlayMedia()
    
    If m_OpenMediaSuccess = True Then
       
       m_MediaPos.CurrentPosition = 0
       m_FilgrpManager.Run
       m_PlayStatus = eRePlay
       
    End If
    
End Sub

'//
'// 停止播放多媒体文件
'//
Public Sub StopMedia()
    
    If m_OpenMediaSuccess = True Then
    
       m_MediaPos.CurrentPosition = 0
       m_FilgrpManager.Stop
       m_PlayStatus = eStop
       
    End If
    
End Sub

'//
'// 暂停播放多媒体文件
'//
Public Sub PauseMedia()
    
    If m_OpenMediaSuccess = True Then
        
        m_FilgrpManager.Pause
        m_PlayStatus = ePause
        
    End If
    
End Sub

'//
'// 改变视频窗口尺寸
'//
Public Sub ChangeVideoWindowSize(Optional ByVal bStretch As Boolean = True, _
    Optional ByVal Left As Long, Optional ByVal TOP As Long, _
    Optional ByVal Width As Long, Optional ByVal Height As Long)
    Dim lpRect As RECT
    Dim lWidth As Long, lHeight As Long
    
    If m_OpenMediaSuccess = True Then
        
        If m_MediaFileType = "Audio" Then Exit Sub
        
        If IsWindow(m_OwnerhWnd) Then
       
           If bStretch = True Then
          
              GetWindowRect m_OwnerhWnd, lpRect
              lWidth = lpRect.Right - lpRect.Left
              lHeight = lpRect.Bottom - lpRect.TOP
              m_VideoWnd.SetWindowPosition 0, 0, _
              lWidth, lHeight
          
           Else
              lWidth = Width
              lHeight = Height
              GetWindowRect m_OwnerhWnd, lpRect
              m_VideoWnd.SetWindowPosition Left, TOP, _
              lWidth, lHeight
              
           End If
           
           ' 保存新的视频尺寸大小
           m_VideoWidth = lWidth
           m_VideoHeight = lHeight
           
       End If
       
    End If
    
End Sub

'//
'// 自动调节视频窗口尺寸
'//
Public Sub AutoChangeVideoSize()
    Dim lpRect As RECT
    Dim lWidth As Long, lHeight As Long
    Dim pLeft As Long, pTop As Long
    Dim pWidth As Long, pHeight As Long
    Dim lNewW As Long, lNewH As Long
    Dim nGene As Single
    
    If m_OpenMediaSuccess = True Then
       
       If m_MediaFileType = eAudio Then Exit Sub
        
       If IsWindow(m_OwnerhWnd) Then
          
          GetWindowRect m_OwnerhWnd, lpRect
          lWidth = lpRect.Right - lpRect.Left: lHeight = lpRect.Bottom - lpRect.TOP
          pLeft = m_Video.SourceLeft: pTop = m_Video.SourceTop
          pWidth = m_Video.SourceWidth: pHeight = m_Video.SourceHeight
          lNewW = pWidth: lNewH = pHeight
          nGene = 1
          
          If lNewW > lWidth Or lNewH > lHeight Then
             ' 等比缩小图像
             Do While lNewW > lWidth Or lNewH > lHeight
                lNewW = Int(pWidth / nGene)
                lNewH = Int(pHeight / nGene)
                nGene = nGene + 0.1
             Loop
          Else
             ' 等比放大图像
             Do While lNewW < lWidth And lNewH < lHeight
                lNewW = Int(pWidth * nGene)
                lNewH = Int(pHeight * nGene)
                nGene = nGene + 0.1
             Loop
          End If
          
          pWidth = lNewW: pHeight = lNewH
          ' 调整视频的显示尺寸
          m_VideoWnd.SetWindowPosition pLeft, pTop, pWidth, pHeight
          ' 保存新的视频尺寸大小
          m_VideoWidth = pWidth
          m_VideoHeight = pHeight
          
       End If
       
    End If
    
End Sub

'//
'// 跳转到指定位置进行播放
'//
Public Sub GoToPosition(ByVal Position As Double)
    
    If m_OpenMediaSuccess = True Then
        m_MediaPos.CurrentPosition = Position
    End If
    
End Sub

'//
'// 设置左右声道,取值范围0-10000
'//
Public Property Let Balance(ByVal Value As Long)
    On Error Resume Next
    
    If m_OpenMediaSuccess = True Then
        m_Audio.Balance = Value
    End If
    
End Property

'//
'// 设置音量大小
'//
Public Property Let Volume(ByVal Value As Long)
    On Error Resume Next
    
    If m_OpenMediaSuccess = True Then
        m_Audio.Volume = Value
    End If
    
End Property

'//
'// 获得视频的宽度
'//
Public Property Get VideoWidth() As Long
    
    If m_OpenMediaSuccess = True Then
        VideoWidth = m_VideoWidth
    Else
        VideoWidth = 0
    End If
    
End Property

'//
'// 获得视频的高度
'//
Public Property Get VideoHeight() As Long
    
    If m_OpenMediaSuccess = True Then
        VideoHeight = m_VideoHeight
    Else
        VideoHeight = 0
    End If
    
End Property

'//
'// 获得媒体速率
'//
Public Property Get Rate() As Double
    
    If m_OpenMediaSuccess = True Then
        Rate = m_MediaPos.Rate
    Else
        Rate = 0
    End If
    
End Property

'//
'// 获得媒体停止时间
'//
Public Property Get StopTime() As Double
    
    If m_OpenMediaSuccess = True Then
        StopTime = m_MediaPos.StopTime
    Else
        StopTime = 0
    End If
    
End Property

'//
'// 获得媒体播放总长度（单位：时间；精度：？）应该将其转换成(HH:MM:SS)，即时分秒的格式查看
'//
Public Property Get Duration() As Double
    
    If m_OpenMediaSuccess = True Then
        Duration = m_MediaPos.Duration
    Else
        Duration = 0
    End If
    
End Property

'//
'// 获得媒体当前播放长度（单位：时间；精度：？）应该将其转换成(HH:MM:SS)，即时分秒的格式查看
'//
Public Property Get CurrentPosition() As Double

    If m_OpenMediaSuccess = True Then
        CurrentPosition = m_MediaPos.CurrentPosition
    Else
        CurrentPosition = 0
    End If
    
End Property
 
'//
'// 获得播放状态
'//
Public Property Get PlayStatus() As PlayStatusConstants
    
    If m_OpenMediaSuccess = True Then
        PlayStatus = m_PlayStatus
    Else
        PlayStatus = eNo
    End If
    
End Property



