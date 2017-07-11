VERSION 5.00
Begin VB.Form FrmTest 
   Caption         =   "图像去雾算法效果锦集"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18735
   Icon            =   "FrmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1249
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame 
      Caption         =   "自适应直方图调节"
      Height          =   1155
      Index           =   4
      Left            =   3060
      TabIndex        =   55
      Top             =   9750
      Visible         =   0   'False
      Width           =   15585
      Begin VB.CheckBox ChkSepChannel 
         Caption         =   "通道分离"
         Height          =   345
         Left            =   5490
         TabIndex        =   68
         Top             =   750
         Width           =   1335
      End
      Begin VB.HScrollBar Contrast 
         Height          =   300
         Left            =   1230
         Max             =   500
         Min             =   50
         TabIndex        =   65
         Top             =   750
         Value           =   150
         Width           =   3750
      End
      Begin VB.HScrollBar TileX 
         Height          =   300
         Left            =   1230
         Max             =   20
         Min             =   2
         TabIndex        =   58
         Top             =   300
         Value           =   8
         Width           =   3750
      End
      Begin VB.HScrollBar TileY 
         Height          =   300
         Left            =   6300
         Max             =   20
         Min             =   2
         TabIndex        =   57
         Top             =   300
         Value           =   8
         Width           =   3750
      End
      Begin VB.HScrollBar CutLimit 
         Height          =   300
         Left            =   11310
         Max             =   100
         Min             =   1
         TabIndex        =   56
         Top             =   300
         Value           =   10
         Width           =   3750
      End
      Begin VB.Label LblAdapt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "对比度控制"
         Height          =   180
         Index           =   6
         Left            =   270
         TabIndex        =   67
         Top             =   810
         Width           =   900
      End
      Begin VB.Label LblAdapt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.5"
         Height          =   180
         Index           =   7
         Left            =   5100
         TabIndex        =   66
         Top             =   840
         Width           =   270
      End
      Begin VB.Label LblAdapt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "水平分块数"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   64
         Top             =   390
         Width           =   900
      End
      Begin VB.Label LblAdapt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         Height          =   180
         Index           =   1
         Left            =   5070
         TabIndex        =   63
         Top             =   360
         Width           =   90
      End
      Begin VB.Label LblAdapt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "垂直分块数"
         Height          =   180
         Index           =   2
         Left            =   5370
         TabIndex        =   62
         Top             =   360
         Width           =   900
      End
      Begin VB.Label LblAdapt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         Height          =   180
         Index           =   3
         Left            =   10140
         TabIndex        =   61
         Top             =   360
         Width           =   90
      End
      Begin VB.Label LblAdapt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "裁剪限幅"
         Height          =   180
         Index           =   4
         Left            =   10500
         TabIndex        =   60
         Top             =   360
         Width           =   720
      End
      Begin VB.Label LblAdapt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.01"
         Height          =   180
         Index           =   5
         Left            =   15150
         TabIndex        =   59
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "基于Retinex的去雾"
      Height          =   765
      Index           =   3
      Left            =   3030
      TabIndex        =   45
      Top             =   8910
      Visible         =   0   'False
      Width           =   15645
      Begin VB.HScrollBar ScaleAmount 
         Height          =   300
         Left            =   1140
         Max             =   8
         Min             =   1
         TabIndex        =   48
         Top             =   330
         Value           =   3
         Width           =   3750
      End
      Begin VB.HScrollBar MaxScale 
         Height          =   300
         Left            =   6180
         Max             =   300
         Min             =   10
         TabIndex        =   47
         Top             =   330
         Value           =   200
         Width           =   3750
      End
      Begin VB.HScrollBar Dynamic 
         Height          =   300
         Left            =   11340
         Max             =   500
         Min             =   100
         TabIndex        =   46
         Top             =   300
         Value           =   200
         Width           =   3750
      End
      Begin VB.Label LblRetinex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "尺度数量"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   54
         Top             =   360
         Width           =   720
      End
      Begin VB.Label LblRetinex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         Height          =   180
         Index           =   1
         Left            =   4950
         TabIndex        =   53
         Top             =   390
         Width           =   90
      End
      Begin VB.Label LblRetinex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最大尺度"
         Height          =   180
         Index           =   2
         Left            =   5310
         TabIndex        =   52
         Top             =   360
         Width           =   720
      End
      Begin VB.Label LblRetinex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "200"
         Height          =   180
         Index           =   3
         Left            =   10080
         TabIndex        =   51
         Top             =   360
         Width           =   270
      End
      Begin VB.Label LblRetinex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "对比度"
         Height          =   180
         Index           =   4
         Left            =   10500
         TabIndex        =   50
         Top             =   360
         Width           =   540
      End
      Begin VB.Label LblRetinex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   180
         Index           =   5
         Left            =   15240
         TabIndex        =   49
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "实时去雾"
      Height          =   735
      Index           =   2
      Left            =   3030
      TabIndex        =   38
      Top             =   8070
      Visible         =   0   'False
      Width           =   15615
      Begin VB.HScrollBar SampleRadius 
         Height          =   300
         Left            =   1080
         Max             =   200
         Min             =   5
         TabIndex        =   40
         Top             =   270
         Value           =   50
         Width           =   3750
      End
      Begin VB.HScrollBar Rho 
         Height          =   300
         Left            =   6180
         Max             =   200
         Min             =   50
         TabIndex        =   39
         Top             =   270
         Value           =   150
         Width           =   3750
      End
      Begin VB.Label LblRealTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         Height          =   180
         Index           =   1
         Left            =   4980
         TabIndex        =   44
         Top             =   330
         Width           =   180
      End
      Begin VB.Label LblRealTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "取样半径"
         Height          =   180
         Index           =   7
         Left            =   270
         TabIndex        =   43
         Top             =   330
         Width           =   720
      End
      Begin VB.Label LblRealTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.5"
         Height          =   180
         Index           =   3
         Left            =   10080
         TabIndex        =   42
         Top             =   330
         Width           =   270
      End
      Begin VB.Label LblRealTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "清晰程度"
         Height          =   180
         Index           =   2
         Left            =   5370
         TabIndex        =   41
         Top             =   330
         Width           =   720
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "模糊去雾"
      Height          =   795
      Index           =   1
      Left            =   3030
      TabIndex        =   27
      Top             =   7200
      Visible         =   0   'False
      Width           =   15675
      Begin VB.OptionButton OptBlurMethod 
         Caption         =   "均值模糊"
         Height          =   225
         Index           =   2
         Left            =   14100
         TabIndex        =   37
         Top             =   390
         Width           =   1155
      End
      Begin VB.OptionButton OptBlurMethod 
         Caption         =   "高斯模糊"
         Height          =   225
         Index           =   1
         Left            =   12870
         TabIndex        =   36
         Top             =   390
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton OptBlurMethod 
         Caption         =   "中值模糊"
         Height          =   225
         Index           =   0
         Left            =   11580
         TabIndex        =   35
         Top             =   390
         Width           =   1155
      End
      Begin VB.HScrollBar Percent 
         Height          =   300
         Left            =   6150
         Max             =   100
         Min             =   50
         TabIndex        =   31
         Top             =   330
         Value           =   80
         Width           =   3750
      End
      Begin VB.HScrollBar BlurRadius 
         Height          =   300
         Left            =   1050
         Max             =   200
         Min             =   5
         TabIndex        =   28
         Top             =   330
         Value           =   50
         Width           =   3750
      End
      Begin VB.Label LblBlur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "模糊方法"
         Height          =   180
         Index           =   4
         Left            =   10650
         TabIndex        =   34
         Top             =   390
         Width           =   720
      End
      Begin VB.Label LblBlur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "清晰程度"
         Height          =   180
         Index           =   2
         Left            =   5340
         TabIndex        =   33
         Top             =   390
         Width           =   720
      End
      Begin VB.Label LblBlur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "80%"
         Height          =   180
         Index           =   3
         Left            =   10050
         TabIndex        =   32
         Top             =   390
         Width           =   270
      End
      Begin VB.Label LblBlur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "模糊半径"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   390
         Width           =   720
      End
      Begin VB.Label LblBlur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         Height          =   180
         Index           =   1
         Left            =   4950
         TabIndex        =   29
         Top             =   390
         Width           =   180
      End
   End
   Begin VB.ComboBox CmbMethod 
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   7650
      Width           =   2115
   End
   Begin VB.Frame Frame 
      Caption         =   "暗通道去雾"
      Height          =   1455
      Index           =   0
      Left            =   3030
      TabIndex        =   5
      Top             =   5670
      Visible         =   0   'False
      Width           =   15645
      Begin VB.HScrollBar SubSample 
         Height          =   300
         Left            =   11280
         Max             =   100
         Min             =   20
         TabIndex        =   22
         Top             =   900
         Value           =   50
         Width           =   3750
      End
      Begin VB.HScrollBar Epsilon 
         Height          =   300
         Left            =   6270
         Max             =   500
         Min             =   10
         TabIndex        =   19
         Top             =   900
         Value           =   100
         Width           =   3750
      End
      Begin VB.HScrollBar MaxAtom 
         Height          =   300
         Left            =   1170
         Max             =   255
         Min             =   200
         TabIndex        =   16
         Top             =   900
         Value           =   240
         Width           =   3750
      End
      Begin VB.HScrollBar Omega 
         Height          =   300
         Left            =   11280
         Max             =   100
         Min             =   50
         TabIndex        =   13
         Top             =   360
         Value           =   95
         Width           =   3750
      End
      Begin VB.HScrollBar GuideRadius 
         Height          =   300
         Left            =   6270
         Max             =   200
         Min             =   5
         TabIndex        =   10
         Top             =   360
         Value           =   60
         Width           =   3750
      End
      Begin VB.HScrollBar DKMinRadius 
         Height          =   300
         Left            =   1200
         Max             =   100
         Min             =   5
         TabIndex        =   7
         Top             =   360
         Value           =   15
         Width           =   3750
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.5"
         Height          =   180
         Index           =   11
         Left            =   15180
         TabIndex        =   23
         Top             =   960
         Width           =   270
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下采样率"
         Height          =   180
         Index           =   10
         Left            =   10530
         TabIndex        =   21
         Top             =   960
         Width           =   720
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.01"
         Height          =   180
         Index           =   9
         Left            =   10110
         TabIndex        =   20
         Top             =   960
         Width           =   360
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Epsilon"
         Height          =   180
         Index           =   8
         Left            =   5460
         TabIndex        =   18
         Top             =   960
         Width           =   630
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "220"
         Height          =   180
         Index           =   7
         Left            =   5040
         TabIndex        =   17
         Top             =   960
         Width           =   270
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最大大气光"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   900
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "95%"
         Height          =   180
         Index           =   5
         Left            =   15180
         TabIndex        =   14
         Top             =   420
         Width           =   270
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "去雾程度"
         Height          =   180
         Index           =   4
         Left            =   10470
         TabIndex        =   12
         Top             =   420
         Width           =   720
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "60"
         Height          =   180
         Index           =   3
         Left            =   10110
         TabIndex        =   11
         Top             =   420
         Width           =   180
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "导向半径"
         Height          =   180
         Index           =   2
         Left            =   5460
         TabIndex        =   9
         Top             =   420
         Width           =   720
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         Height          =   180
         Index           =   1
         Left            =   5040
         TabIndex        =   8
         Top             =   420
         Width           =   180
      End
      Begin VB.Label LblDarkChannel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最小值半径"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   900
      End
   End
   Begin VB.PictureBox PicTemp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6810
      Left            =   14160
      Picture         =   "FrmTest.frx":1CFA
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   4
      Top             =   8160
      Visible         =   0   'False
      Width           =   9060
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存图像"
      Height          =   375
      Left            =   1500
      TabIndex        =   3
      Top             =   7140
      Width           =   1275
   End
   Begin VB.PictureBox PicDest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6750
      Left            =   9480
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   2
      Top             =   240
      Width           =   9000
   End
   Begin VB.PictureBox PicSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6750
      Left            =   240
      ScaleHeight     =   450
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   240
      Width           =   9000
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "选择图像"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   7140
      Width           =   1275
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "信息："
      Height          =   180
      Left            =   240
      TabIndex        =   26
      Top             =   8130
      Width           =   540
   End
   Begin VB.Label Lblmethod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "算法"
      Height          =   180
      Left            =   210
      TabIndex        =   24
      Top             =   7710
      Width           =   360
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'**    开发日期 ：  2009-7-21
'**    作    者 ：  laviewpbt
'**    联系方式：   33184777
'**    修改日期 ：   2009-7-21
'**    版    本 ：  Version 1.3.1
'**    转载请不要删除以上信息
'****************************************************************************************



Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)

Private Declare Sub HazeRemovalBasedOnImageBlur Lib "ImageProcessing.dll" (ByVal Scan0 As Long, ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal DarkRadius As Long, ByVal Radius As Long, ByVal Percent As Long, ByVal BlurFunction As Long)
Private Declare Sub RealTimeHazeRemoval Lib "ImageProcessing.dll" (ByVal Scan0 As Long, ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal Radius As Long, ByVal Rho As Single)
Private Declare Sub HazeRemovalUseDarkChannelPrior Lib "ImageProcessing.dll" (ByVal Scan0 As Long, ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal Radius As Long, ByVal GuideRadius As Long, ByVal MaxAtom As Long, ByVal Omega As Single, ByVal Epsilon As Single, ByVal T0 As Single, ByVal SubSample As Single)
Private Declare Sub AdaptAutoLevelOrContrast Lib "ImageProcessing.dll" (ByVal Scan0 As Long, ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal TileX As Long, ByVal TileY As Long, ByVal CutLimit As Single, ByVal Contrast As Single, ByVal SeparateChannel As Integer)
Private Declare Sub AdaptHistEqualize Lib "ImageProcessing.dll" (ByVal Scan0 As Long, ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal TileX As Long, ByVal TileY As Long, ByVal CutLimit As Single, ByVal SeparateChannel As Integer)
Private Declare Sub MSRCR Lib "ImageProcessing.dll" (ByVal Scan0 As Long, ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal MaxScale As Single, ByVal ScaleAmount As Long, ByVal Dynamic As Single)

Private Declare Sub minn Lib "ImageProcessing.dll" (ByVal Scan0 As Long)


Private Img         As New Cimage
Attribute Img.VB_VarHelpID = -1
Private ImgC        As New Cimage
Private FitX        As Long
Private FitY        As Long
Private FitWidth    As Long
Private FitHeight   As Long
Private Init        As Boolean
 

Private BlurMethod  As Long





'Private Sub Command1_Click()
''     Dim i As Long
''     For i = 100 To 240
''        Img.LoadPictureFromStdPicture LoadPicture("c:\2\0" & i & ".jpg")
''        'RealTimeHazeRemoval Img.Pointer, Img.Width, Img.Height, Img.Stride, SampleRadius.Value, Rho.Value * 0.01
''        'HazeRemovalBasedOnImageBlur Img.Pointer, Img.Width, Img.Height, Img.Stride, IIf(Img.Width > Img.Height, IIf(Img.Width * 0.02 > 5, Img.Width * 0.02, 5), IIf(Img.Height * 0.02 > 5, Img.Height * 0.02, 5)), BlurRadius.Value, Percent.Value, BlurMethod
''
''        HazeRemovalUseDarkChannelPrior Img.Pointer, Img.Width, Img.Height, Img.Stride, DKMinRadius.Value, GuideRadius.Value, MaxAtom.Value, Omega.Value * 0.01, Epsilon.Value * 0.0001, T0.Value * 0.01
''
''        PicTemp.Move 0, 0, Img.Width, Img.Height
''        Img.OutPut PicTemp.hdc
''        SavePicture PicTemp.Image, "c:\22\" & i & ".bmp"
''        Img.DisposeResource
''    Next
''
'End Sub

Private Sub Form_Load()
    Img.LoadPictureFromStdPicture PicTemp.Picture
    Img.Render PicSrc.hdc, FitX, FitY, FitWidth, FitHeight, 0, 0, Img.Width, Img.Height
    PicSrc.Refresh
    Set ImgC = Img.Clone
    BlurMethod = 1
    CmbMethod.AddItem "暗通道去雾"
    CmbMethod.AddItem "模糊去雾"
    CmbMethod.AddItem "实时去雾"
    CmbMethod.AddItem "多尺度Retinex"
    CmbMethod.AddItem "自适应直方图均衡化"
    CmbMethod.AddItem "自适应色阶和对比度"
    Init = True
    DoEvents
    CmbMethod.ListIndex = 0
    CmbMethod_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Img.DisposeResource
    ImgC.DisposeResource
    Set Img = Nothing
    Set ImgC = Nothing
End Sub

Private Sub CmbMethod_Click()
    If Init = True Then
        Dim Index As Long, Y As Long
        For Y = Frame.LBound To Frame.UBound
            Frame(Y).Visible = False
        Next
        If CmbMethod.ListIndex >= 4 Then
            Frame(4).Visible = True
            Frame(4).Move 192, 478, Frame(4).Width, Frame(4).Height
            If CmbMethod.ListIndex = 4 Then
                LblAdapt(6).Enabled = False
                LblAdapt(7).Enabled = False
                Contrast.Enabled = False
            Else
                LblAdapt(6).Enabled = True
                LblAdapt(7).Enabled = True
                Contrast.Enabled = True
            End If
        Else
            Frame(CmbMethod.ListIndex).Visible = True
            Frame(CmbMethod.ListIndex).Move 192, 478, Frame(CmbMethod.ListIndex).Width, Frame(CmbMethod.ListIndex).Height
        End If
        UpdateImage
    End If
End Sub

Private Sub CmdOpen_Click()
    On Error GoTo ErrHandle:
    Dim FileName            As String
    FileName = API.ShowOpen(Me.Hwnd, "All Suppported Images |*.bmp;*.jpg;|BMP Images|*.bmp|JPG Images|*.jpg", "选择图像")
    If FileName <> "" Then
        If Img.handle <> 0 Then
            Img.DisposeResource
            ImgC.DisposeResource
        End If
        Img.LoadPictureFromStdPicture LoadPicture(FileName)
        Set ImgC = Img.Clone
        API.GetBestFitInfoEx Img.Width, Img.Height, 600, 450, FitX, FitY, FitWidth, FitHeight
        PicSrc.Cls
        Img.Render PicSrc.hdc, FitX, FitY, FitWidth, FitHeight, 0, 0, Img.Width, Img.Height
        PicSrc.Refresh
        UpdateImage
    End If
    Exit Sub
ErrHandle:
    MsgBox "不支持的图像格式。", vbCritical
End Sub

Private Sub CmdSave_Click()
    Dim FileName            As String
    FileName = API.ShowSave(Me.Hwnd, "BMP Images|*.bmp", "保存图像", 0)
    If FileName <> "" Then
        PicTemp.Move 0, 0, Img.Width, Img.Height
        ImgC.OutPut PicTemp.hdc
        SavePicture PicTemp.Image, FileName & ".bmp"
    End If
End Sub

Private Sub DKMinRadius_Change()
    DKMinRadius_Scroll
End Sub

Private Sub DKMinRadius_Scroll()
    LblDarkChannel(1).Caption = DKMinRadius.Value
    UpdateImage
End Sub

Private Sub GuideRadius_Change()
    GuideRadius_Scroll
End Sub
Private Sub GuideRadius_Scroll()
    LblDarkChannel(3).Caption = GuideRadius.Value
    UpdateImage
End Sub

Private Sub Omega_Change()
    Omega_Scroll
End Sub
Private Sub Omega_Scroll()
    LblDarkChannel(5).Caption = Omega.Value & "%"
    UpdateImage
End Sub

Private Sub MaxAtom_Change()
    MaxAtom_Scroll
End Sub
Private Sub MaxAtom_Scroll()
    LblDarkChannel(7).Caption = MaxAtom.Value
    DoEvents
    UpdateImage
End Sub

Private Sub SubSample_Change()
     LblDarkChannel(11).Caption = Format(SubSample.Value * 0.01, "0.00")
     UpdateImage
End Sub

Private Sub SubSample_Validate(Cancel As Boolean)
    SubSample_Change
End Sub


Private Sub Epsilon_Change()
    Epsilon_Scroll
End Sub

Private Sub Epsilon_Scroll()
    LblDarkChannel(9).Caption = Format(Epsilon.Value * 0.0001, "0.0000")
    UpdateImage
End Sub

Private Sub BlurRadius_Change()
    BlurRadius_Scroll
End Sub

Private Sub BlurRadius_Scroll()
    LblBlur(1).Caption = BlurRadius.Value
    UpdateImage
End Sub

Private Sub Percent_Change()
    Percent_Scroll
End Sub

Private Sub Percent_Scroll()
    LblBlur(3).Caption = Percent.Value & "%"
    UpdateImage
End Sub

Private Sub OptBlurMethod_Click(Index As Integer)
    BlurMethod = Index
    UpdateImage
End Sub

Private Sub SampleRadius_Change()
    SampleRadius_Scroll
End Sub

Private Sub SampleRadius_Scroll()
    LblRealTime(1).Caption = SampleRadius.Value
    UpdateImage
End Sub

Private Sub Rho_Change()
    Rho_Scroll
End Sub

Private Sub Rho_Scroll()
    LblRealTime(3).Caption = Rho.Value * 0.01
    UpdateImage
End Sub

Private Sub ScaleAmount_Change()
    ScaleAmount_Scroll
End Sub
Private Sub ScaleAmount_Scroll()
    LblRetinex(1).Caption = ScaleAmount.Value
    UpdateImage
End Sub

Private Sub MaxScale_Change()
    MaxScale_Scroll
End Sub

Private Sub MaxScale_Scroll()
    LblRetinex(3).Caption = MaxScale.Value
    UpdateImage
End Sub

Private Sub Dynamic_Change()
    Dynamic_Scroll
End Sub

Private Sub Dynamic_Scroll()
    LblRetinex(5).Caption = Dynamic.Value * 0.01
    UpdateImage
End Sub


Private Sub TileX_Change()
    TileX_Scroll
End Sub

Private Sub TileX_Scroll()
    LblAdapt(1).Caption = TileX.Value
    UpdateImage
End Sub

Private Sub TileY_Change()
    TileY_Scroll
End Sub

Private Sub TileY_Scroll()
    LblAdapt(3).Caption = TileY.Value
    UpdateImage
End Sub

Private Sub CutLimit_Change()
    CutLimit_Scroll
End Sub
Private Sub CutLimit_Scroll()
    LblAdapt(5).Caption = Format(CutLimit.Value * 0.001, "0.000")
    UpdateImage
End Sub

Private Sub Contrast_Change()
    Contrast_Scroll
End Sub
Private Sub Contrast_Scroll()
    LblAdapt(7).Caption = Format(Contrast.Value * 0.01, "0.00")
    UpdateImage
End Sub

Private Sub ChkSepChannel_Click()
    UpdateImage
End Sub

Private Sub UpdateImage()
    CopyMemory ByVal ImgC.Pointer, ByVal Img.Pointer, Img.Height * Img.Stride
    Dim TimeElpase As Currency
    TimeElpase = API.GetCurrentTime
    Select Case CmbMethod.ListIndex
    Case 0
        HazeRemovalUseDarkChannelPrior ImgC.Pointer, ImgC.Width, ImgC.Height, ImgC.Stride, DKMinRadius.Value, GuideRadius.Value, MaxAtom.Value, Omega.Value * 0.01, Epsilon.Value * 0.0001, 0.1, SubSample.Value * 0.01
    Case 1
        HazeRemovalBasedOnImageBlur ImgC.Pointer, ImgC.Width, ImgC.Height, ImgC.Stride, IIf(Img.Width > Img.Height, IIf(Img.Width * 0.02 > 5, Img.Width * 0.02, 5), IIf(Img.Height * 0.02 > 5, Img.Height * 0.02, 5)), BlurRadius.Value, Percent.Value, BlurMethod
    Case 2
        RealTimeHazeRemoval ImgC.Pointer, ImgC.Width, ImgC.Height, ImgC.Stride, SampleRadius.Value, Rho.Value * 0.01
    Case 3
        MSRCR ImgC.Pointer, ImgC.Width, ImgC.Height, ImgC.Stride, MaxScale.Value, ScaleAmount.Value, Dynamic.Value * 0.01
    Case 4
        AdaptHistEqualize ImgC.Pointer, ImgC.Width, ImgC.Height, ImgC.Stride, TileX.Value, TileY.Value, CutLimit.Value * 0.001, IIf(ChkSepChannel.Value = vbChecked, 1, 0)
    Case 5
        AdaptAutoLevelOrContrast ImgC.Pointer, ImgC.Width, ImgC.Height, ImgC.Stride, TileX.Value, TileY.Value, CutLimit.Value * 0.001, Contrast.Value * 0.01, IIf(ChkSepChannel.Value = vbChecked, 1, 0)
    End Select
    lblinfo.Caption = "用时 " & API.GetCurrentTime - TimeElpase & " ms"
    PicDest.Cls
    ImgC.Render PicDest.hdc, FitX, FitY, FitWidth, FitHeight, 0, 0, Img.Width, Img.Height
    PicDest.Refresh
End Sub


