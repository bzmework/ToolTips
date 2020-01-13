VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "测试"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   503
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdBorderBar 
      Caption         =   "BorderBarTip"
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdSystem 
      Caption         =   "SystemTip"
      Height          =   615
      Left            =   5760
      TabIndex        =   3
      Top             =   2190
      Width           =   1335
   End
   Begin VB.CommandButton cmdPopup 
      Caption         =   "PopupTip"
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   3090
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   3105
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmMain.frx":0000
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "CustomTip"
      Height          =   615
      Left            =   5760
      TabIndex        =   0
      Top             =   540
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   3105
      Left            =   120
      Picture         =   "frmMain.frx":00EE
      Top             =   3720
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   150
      Picture         =   "frmMain.frx":216DE
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsTip As TToolTip.CustomTip
Private mclsPopTip As TToolTip.PopupTip
Private mclSysTip As TToolTip.SystemTip

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTitle As String
    Dim strText As String
    
    strTitle = "沁园春·雪"
    
    strText = _
    "北国风光，千里冰封，万里雪飘。" & vbCrLf & _
    "望长城内外，惟余莽莽；大河上下，顿失滔滔。" & vbCrLf & _
    "山舞银蛇，原驰蜡象，欲与天公试比高。" & vbCrLf & _
    "须晴日，看红装素裹，分外妖娆。" & vbCrLf & _
    "江山如此多娇，引无数英雄竞折腰。" & vbCrLf & _
    "惜秦皇汉武，略输文采；唐宗宋祖，稍逊风骚。"
    
    mclsTip.Title = strTitle
    mclsTip.Text = strText
    mclsTip.Show Text1.hWnd
    
End Sub

Private Sub cmdCustom_Click()
    
    mclsTip.TipStyle = Custom
    mclsTip.TipType = Warning
    Set mclsTip.TitleIcon = Image1.Picture
    Set mclsTip.BackPicture = Image2.Picture
    mclsTip.TitleIconSize = Icon32
    mclsTip.TitleColor = vbBlue
    mclsTip.TextColor = vbRed
    mclsTip.BeginColor = vbGreen
    mclsTip.BorderColor = vbButtonText
    mclsTip.Alpha = 192
    
End Sub

Private Sub cmdCustom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTitle As String
    Dim strText As String
    
    strTitle = "沁园春·雪"
    
    strText = _
    "北国风光，千里冰封，万里雪飘。" & vbCrLf & _
    "望长城内外，惟余莽莽；大河上下，顿失滔滔。" & vbCrLf & _
    "山舞银蛇，原驰蜡象，欲与天公试比高。" & vbCrLf & _
    "须晴日，看红装素裹，分外妖娆。" & vbCrLf & _
    "江山如此多娇，引无数英雄竞折腰。" & vbCrLf & _
    "惜秦皇汉武，略输文采；唐宗宋祖，稍逊风骚。"

    mclsTip.Title = strTitle
    mclsTip.Text = strText
    mclsTip.Show cmdCustom.hWnd
    
End Sub

Private Sub cmdBorderBar_Click()
    
    mclsTip.TipStyle = BorderBar
    Set mclsTip.TitleIcon = Image1.Picture
    Set mclsTip.BackPicture = Image2.Picture
    mclsTip.TitleIconSize = Icon32
    mclsTip.TitleColor = vbBlue
    mclsTip.TextColor = vbRed
    mclsTip.BeginColor = vbGreen
    mclsTip.BorderColor = vbButtonText
    mclsTip.Alpha = 192
    
End Sub

Private Sub cmdBorderBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTitle As String
    Dim strText As String
    
    strTitle = "沁园春·雪"
    
    strText = _
    "北国风光，千里冰封，万里雪飘。" & vbCrLf & _
    "望长城内外，惟余莽莽；大河上下，顿失滔滔。" & vbCrLf & _
    "山舞银蛇，原驰蜡象，欲与天公试比高。" & vbCrLf & _
    "须晴日，看红装素裹，分外妖娆。" & vbCrLf & _
    "江山如此多娇，引无数英雄竞折腰。" & vbCrLf & _
    "惜秦皇汉武，略输文采；唐宗宋祖，稍逊风骚。"

    mclsTip.Title = strTitle
    mclsTip.Text = strText
    mclsTip.Show cmdBorderBar.hWnd
    
End Sub


Private Sub cmdSystem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTitle As String
    Dim strText As String
    
    If mclSysTip.Alive = False Then
        
        With mclSysTip
            .DelayTime = 25 '毫秒
            .KillTime = 500 '毫秒
            
            .BackColor = vbRed
            .TextColor = vbGreen
            .GradientColorStart = vbRed
            .GradientColorEnd = vbYellow
            .BackStyle = 3  '1 - Solid 2 - H Gradient  3 - V Gradient  4 - Picture
            .Font.Name = Me.FontName
            .Shadow = True
            .ToolTipStyle = TTBalloon
            .IconSize = TTIcon16
            Set .Picture = Image2.Picture
            
        End With
    
        strTitle = "沁园春·雪"
        
        strText = _
        "北国风光，千里冰封，万里雪飘。" & vbCrLf & _
        "望长城内外，惟余莽莽；大河上下，顿失滔滔。" & vbCrLf & _
        "山舞银蛇，原驰蜡象，欲与天公试比高。" & vbCrLf & _
        "须晴日，看红装素裹，分外妖娆。" & vbCrLf & _
        "江山如此多娇，引无数英雄竞折腰。" & vbCrLf & _
        "惜秦皇汉武，略输文采；唐宗宋祖，稍逊风骚。"


        mclSysTip.ShowToolTip cmdSystem.hWnd, strTitle, strText, Image1.Picture.Handle, 90
        
    End If
    
End Sub

Private Sub cmdPopup_Click()
    Dim strText As String
    
    strText = _
    "北国风光，千里冰封，万里雪飘。" & vbCrLf & _
    "望长城内外，惟余莽莽；大河上下，顿失滔滔。" & vbCrLf & _
    "山舞银蛇，原驰蜡象，欲与天公试比高。" & vbCrLf & _
    "须晴日，看红装素裹，分外妖娆。" & vbCrLf & _
    "江山如此多娇，引无数英雄竞折腰。" & vbCrLf & _
    "惜秦皇汉武，略输文采；唐宗宋祖，稍逊风骚。"
    
    mclsPopTip.BackgroundStyle = ImagePic
    mclsPopTip.SetBackPicture Image2.Picture
    mclsPopTip.DisplayAlert strText, , , Online
    
End Sub

Private Sub Form_Load()

    Set mclsTip = New TToolTip.CustomTip
    Set mclsPopTip = New TToolTip.PopupTip
    Set mclSysTip = New TToolTip.SystemTip
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mclsTip = Nothing
    Set mclsPopTip = Nothing
    Set mclSysTip = Nothing
    
End Sub


