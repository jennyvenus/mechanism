VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form dlgPara 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "机构学演示程序 - 参数输入"
   ClientHeight    =   5610
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   8
      Left            =   2550
      TabIndex        =   36
      Text            =   "1"
      Top             =   2820
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   7
      Left            =   2550
      TabIndex        =   34
      Text            =   "15"
      Top             =   2505
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   6
      Left            =   2550
      TabIndex        =   33
      Text            =   "126"
      Top             =   2190
      Width           =   705
   End
   Begin VB.ComboBox cbo1 
      Height          =   300
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   1290
      Width           =   1155
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      ItemData        =   "dlgParaMent.frx":0000
      Left            =   945
      List            =   "dlgParaMent.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   285
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   5
      Left            =   2550
      TabIndex        =   27
      Text            =   "56"
      Top             =   1860
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   4
      Left            =   2550
      TabIndex        =   26
      Text            =   "63"
      Top             =   1530
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   3
      Left            =   2550
      TabIndex        =   21
      Text            =   "50"
      Top             =   1200
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   2
      Left            =   2550
      TabIndex        =   20
      Text            =   "75"
      Top             =   885
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   1
      Left            =   2550
      TabIndex        =   19
      Text            =   "100"
      Top             =   570
      Width           =   705
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   0
      Left            =   2550
      TabIndex        =   18
      Text            =   "50"
      Top             =   270
      Width           =   705
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   17
      Top             =   5220
      Width           =   1125
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   15
      Top             =   4515
      Width           =   1125
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   13
      Top             =   3825
      Width           =   1125
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   165
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   11
      Top             =   3255
      Width           =   1125
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   9
      Top             =   2610
      Width           =   1125
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   1095
      TabIndex        =   7
      Top             =   1950
      Width           =   1125
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5940
      TabIndex        =   4
      Top             =   5175
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   930
      ScaleHeight     =   240
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   615
      Width           =   645
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6660
      Top             =   5670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   5175
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "A'E"
      Height          =   240
      Left            =   2055
      TabIndex        =   37
      Top             =   2175
      Width           =   360
   End
   Begin VB.Label Label17 
      Caption         =   "线条宽度"
      Height          =   210
      Left            =   1665
      TabIndex        =   35
      Top             =   2850
      Width           =   720
   End
   Begin VB.Label Label16 
      Caption         =   "放大倍数"
      Height          =   330
      Left            =   1665
      TabIndex        =   32
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "BC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2115
      TabIndex        =   29
      Top             =   1875
      Width           =   480
   End
   Begin VB.Label Label14 
      Caption         =   "BA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2115
      TabIndex        =   28
      Top             =   1545
      Width           =   420
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "NB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2115
      TabIndex        =   25
      Top             =   1230
      Width           =   180
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "O1N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2070
      TabIndex        =   24
      Top             =   915
      Width           =   270
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "MN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2100
      TabIndex        =   23
      Top             =   630
      Width           =   180
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "OM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   2085
      TabIndex        =   22
      Top             =   330
      Width           =   180
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "纸盒的颜色"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   16
      Top             =   4965
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "传送带的颜色"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   14
      Top             =   4275
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "杆的颜色"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   12
      Top             =   3585
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "凸轮2边界颜色"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   165
      TabIndex        =   10
      Top             =   3000
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "凸轮1边界颜色"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   8
      Top             =   2325
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "凸轮1填充颜色"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   6
      Top             =   1665
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "凸轮1填充模式"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   5
      Top             =   990
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "背景颜色"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   2
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "增长角度"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   150
      TabIndex        =   1
      Top             =   285
      Width           =   720
   End
End
Attribute VB_Name = "dlgPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Me.Visible = False
End Sub

Private Sub Form_Activate()
    Text1(0).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim M As msg
    If KeyCode = vbKeyReturn Then
        PeekMessage M, Me.hwnd, 0, 0, PM_REMOVE
    End If
End Sub

Private Sub Form_Load()
    cbo.AddItem "3"
    cbo.AddItem "4"
    cbo.AddItem "5"
    cbo.AddItem "6"
    cbo.AddItem "8"
    cbo.AddItem "9"
    cbo.ListIndex = 5
    cbo1.AddItem "实心"
    cbo1.AddItem "空心"
    cbo1.AddItem "横"
    cbo1.AddItem "竖"
    cbo1.AddItem "反斜"
    cbo1.AddItem "正斜"
    cbo1.AddItem "正十字"
    cbo1.AddItem "斜十字"
    
    cbo1.ListIndex = 0
End Sub

Private Sub OKButton_Click()
    ZoomInMultiple = Val(Text1(7).Text)
    Dim i As Integer
    For i = 0 To 120
        frmmain.pic(i).DrawWidth = Val(Text1(8).Text)
    Next i
    For i = 1 To 125
        Pos(i).x = 0
        Pos(i).y = 0
    Next i
    OO1.A = 100                            'OO 类是杆类
    OO1.T = 318 * PI / 180                'A是杆长
    OO2.A = 149                            'T是幅角
    OO2.T = 35 * PI / 180                  'X，Y是起点坐标
    OO3.A = Sqr(120 ^ 2 + 45 ^ 2)
    OO4.A = 100
    O.CenterX = CamCenter.x          'O类是铰支类
    O.CenterY = CamCenter.y          'CENTER 是中心
    OO3.T = Atn(45 / 120)
    O3.CenterX = O.CenterX + 120 * ZoomInMultiple
    O3.CenterY = O.CenterY - 45 * ZoomInMultiple
    O4.CenterX = O.CenterX - 100 * ZoomInMultiple
    O4.CenterY = O.CenterY
    O5.CenterX = O.CenterX - 278 * ZoomInMultiple
    O5.CenterY = O.CenterY + 138 * ZoomInMultiple
    O1.CenterX = O.CenterX + OO1.A * ZoomInMultiple * Cos(OO1.T)
    O1.CenterY = O.CenterY - OO1.A * ZoomInMultiple * Sin(OO1.T)
    O2.CenterX = O.CenterX + OO2.A * ZoomInMultiple * Cos(OO2.T)
    O2.CenterY = O.CenterY - OO2.A * ZoomInMultiple * Sin(OO2.T)
    
    
    O.Angle = 270 * PI / 180
    O1.Angle = 270 * PI / 180
    O2.Angle = 270 * PI / 180
    O3.Angle = 0
    O4.Angle = 270 * PI / 180
    O5.Angle = 270 * PI / 180
    
    OM.A = Val(Text1(0).Text)
    OM.x = O.CenterX
    OM.y = O.CenterY
    
    MN.A = Val(Text1(1).Text)
    
    O1N.A = Val(Text1(2).Text)
    O1N.x = O1.CenterX
    O1N.y = O1.CenterY
    
    NB.A = Val(Text1(3).Text)
    
    AB.A = Val(Text1(4).Text)
    BC.A = Val(Text1(5).Text)
    
    O4F.A = 108
    O4F.x = O4.CenterX
    O4F.y = O4.CenterY
    
    O4J.A = 65
    O4J.x = O4.CenterX
    O4J.y = O4.CenterY
    
    O5G.A = 104.5
    O5G.x = O5.CenterX
    O5G.y = O5.CenterY
    
    O5H.A = O5G.A * 2
    O5H.x = O5.CenterX
    O5H.y = O5.CenterY
    
    GF.A = 193
    
    MB.A = Sqr(MN.A ^ 2 + NB.A ^ 2)
    
    OB.x = O.CenterX
    OB.y = O.CenterY
    
    O1M.x = O1.CenterX
    O1M.y = O1.CenterY
    
    O2B.x = O2.CenterX
    O2B.y = O2.CenterY
    
    O5O4.A = Sqr(((O5.CenterX - O4.CenterX) / ZoomInMultiple) ^ 2 + ((O5.CenterY - O4.CenterY) / ZoomInMultiple) ^ 2)
    O5O4.T = -Atn((O5.CenterY - O4.CenterY) / (O5.CenterX - O4.CenterX))
    O5F.x = O5.CenterX
    O5F.y = O5.CenterY
    
    O2A.x = O2.CenterX
    O2A.y = O2.CenterY
    O2A.A = BC.A
    
    O2C.A = AB.A
    O2C.x = O2.CenterX
    O2C.y = O2.CenterY
    O2CP.x = O2.CenterX
    O2CP.y = O2.CenterY
    O2AP.x = O2.CenterX
    O2AP.y = O2.CenterY
    O2CP.A = AB.A
    O2AP.A = O2A.A * 2
    APE.A = Val(Text1(6).Text)
    CPBP.A = BC.A * 2
    
    
    
    Ground3.Aclinic = True
    Ground3.Length = 100
    Ground3.Near_Left = True
    Ground3.UpOrDown = True
    Ground3.SmallS = 5
    Ground3.StartX = O3.CenterX - ZoomInMultiple * (190 + 22 + Ground3.Length)
    Ground3.StartY = O3.CenterY - ZoomInMultiple * (30)
    
    Ground4.Aclinic = False
    Ground4.Length = 70
    Ground4.Near_Left = False
    Ground4.UpOrDown = True
    Ground4.SmallS = 5
    Ground4.StartX = Ground3.StartX + Ground3.Length * ZoomInMultiple
    Ground4.StartY = Ground3.StartY
    
    Handspike.Length = 255
    Handspike.HandleLength = 80
    Handspike.XWidth = 10
    Handspike.XHeight = 20
    Handspike.CenterY = Ground3.StartY - ZoomInMultiple * Handspike.XHeight
    
    SpinBlock1.XHeight = 40
    SpinBlock1.XWidth = 25
    Ground1.Aclinic = True
    Ground1.Length = 70
    Ground1.Near_Left = True
    Ground1.UpOrDown = False
    Ground1.SmallS = 5
    Ground1.StartX = Ground3.StartX + 10 * ZoomInMultiple
    Ground1.StartY = Ground3.StartY - 6 * Handspike.XHeight * ZoomInMultiple / 5
    
    Ground2.Aclinic = True
    Ground2.Length = 70
    Ground2.Near_Left = True
    Ground2.UpOrDown = True
    Ground2.SmallS = 5
    Ground2.StartX = Ground1.StartX
    Ground2.StartY = Ground3.StartY - 4 * Handspike.XHeight * ZoomInMultiple / 5
    
    BOX1.BoxHeight = 66
    BOX1.BoxWidth = 87
    BOX1.BoxLength = 2500
    BOX1.Angle = PI / 2
    BOX1.StartX = -5000
    BOX2.BoxHeight = 66
    BOX2.BoxWidth = 87
    BOX2.BoxLength = 2500
    BOX2.Angle = 0
    BOX2.StartX = -5000
    BOX3.BoxHeight = 66
    BOX3.BoxWidth = 87
    BOX3.BoxLength = 2500
    BOX3.Angle = PI / 2
    BOX3.StartX = -5000
    
    Ground5.Aclinic = True
    Ground5.Length = 30
    Ground5.Near_Left = True
    Ground5.SmallS = 5
    Ground5.UpOrDown = True
    Ground5.StartX = Ground4.StartX + ZoomInMultiple * (10 + BOX1.BoxWidth + 10)
    Ground5.StartY = Ground3.StartY
    
    cam1.CenterX = CamCenter.x
    cam1.CenterY = CamCenter.y
    cam1.CT1 = 90
    cam1.CT2 = 120
    cam1.CT3 = 80
    cam1.CT4 = 70
    cam1.C = OO3.A
    cam1.GM = Atn(45 / 120)
    cam1.TT0 = 189 * PI / 180
    cam1.TTH = PI / 10
    cam1.RF = 80
    cam1.RR = 15
    
    cam2.CenterX = CamCenter.x
    cam2.CenterY = CamCenter.y
    cam2.CT1 = 120
    cam2.CT2 = 90
    cam2.CT3 = 130
    cam2.CT4 = 20
    cam2.C = OO4.A
    cam2.GM = PI
    cam2.TT0 = (270 + AngleFO4J + 37.65 / 2) * PI / 180
    cam2.TTH = 37.65 * PI / 180
    cam2.RF = 65
    cam2.RR = 10
    
    O3K.x = O3.CenterX
    O3K.y = O3.CenterY
    'cam1.RF = Val(Text1(7).Text)
    O3K.A = cam1.RF
    O3D.A = 190
    O3D.x = O3.CenterX
    O3D.y = O3.CenterY
    CT33 = 110 * PI / 180
    
    TB1.x = Ground4.StartX + ZoomInMultiple * 5
    TB1.XColor = vbBlue
    TB1.XWidth = BOX1.BoxWidth + 10
    TB1.XHeight = 50
    
    CSD.Radius = 20
    CSD.LeftCenterX = Ground5.StartX + ZoomInMultiple * (Ground5.Length + 10)
    CSD.LeftCenterY = Ground5.StartY + ZoomInMultiple * CSD.Radius
    CSD.XColor = Picture6.BackColor
    CSD.Length = 300
    
    BOX1.StartY = Ground3.StartY - ZoomInMultiple * BOX1.BoxHeight
    OM.P = PI
    OO1.V = 0
    OO1.P = 0
    O1O2.A = Sqr((149 * Sin(35 * PI / 180) - 100 * Sin(318 * PI / 180)) ^ 2 + (149 * Cos(35 * PI / 180) - 100 * Cos(318 * PI / 180)) ^ 2)
    O1O2.T = Atn((149 * Sin(35 * PI / 180) - 100 * Sin(318 * PI / 180)) / (149 * Cos(35 * PI / 180) - 100 * Cos(318 * PI / 180))) + PI
    
    OO1.XColor = Picture5.BackColor
    OO2.XColor = Picture5.BackColor
    OO3.XColor = Picture5.BackColor
    OO4.XColor = Picture5.BackColor
    OO5.XColor = Picture5.BackColor
    O4F.XColor = Picture5.BackColor
    O4J.XColor = Picture5.BackColor
    O5G.XColor = Picture5.BackColor
    O5H.XColor = Picture5.BackColor
    GF.XColor = Picture5.BackColor
    'OM.XColor = Picture5.BackColor
    OM.XColor = RGB(255, 0, 0)
    'MN.XColor = Picture5.BackColor
    MN.XColor = RGB(0, 0, 255)
    O1M.XColor = Picture5.BackColor
    O1N.XColor = Picture5.BackColor
    'NB.XColor = Picture5.BackColor
    NB.XColor = MN.XColor
    MB.XColor = Picture5.BackColor
    OB.XColor = Picture5.BackColor
    O2B.XColor = Picture5.BackColor
    O2A.XColor = Picture5.BackColor
    AB.XColor = Picture5.BackColor
    BC.XColor = Picture5.BackColor
    O2C.XColor = Picture5.BackColor
    O2CP.XColor = Picture5.BackColor
    O2AP.XColor = Picture5.BackColor
    CPBP.XColor = Picture5.BackColor
    APE.XColor = Picture5.BackColor
    O3K.XColor = Picture5.BackColor
    O3D.XColor = Picture5.BackColor
    O1B.XColor = Picture5.BackColor
    O1O2.XColor = Picture5.BackColor
    O2E.XColor = Picture5.BackColor
    OE.XColor = Picture5.BackColor
    cam1.XColor = Picture3.BackColor
    cam2.XColor = Picture4.BackColor
    BOX1.XColor = Picture7.BackColor
    BOX2.XColor = Picture7.BackColor
    BOX3.XColor = Picture7.BackColor

    Me.Hide
    Load frmSplash
    frmSplash.Show
    MaxFrame = 360 / Val(cbo.Text)
    For i = 1 To 360 Step Val(cbo.Text)
        DoEvents
        frmmain.pic(Int(i / Val(cbo.Text)) + 1).Cls
        RotateDemo frmmain.pic(Int(i / Val(cbo.Text)) + 1), i * PI / 180
    Next i
    frmSplash.Hide
    DrawWhichPic = 1
    Set frmmain.pic(0).Picture = frmmain.pic(DrawWhichPic).Image
End Sub

Private Sub Picture1_Click()
    On Error GoTo errHandle
    Dim i As Integer
    cd.ShowColor
    For i = 1 To 120
        frmmain.pic(i).BackColor = cd.Color
    Next i
    Picture1.BackColor = cd.Color
errHandle:
    Err.Clear
End Sub

Private Sub Picture2_Click()
    On Error GoTo errHandle
    cd.ShowColor
    Picture2.BackColor = cd.Color
errHandle:
    Err.Clear
End Sub

Private Sub Picture3_Click()
    On Error GoTo errHandle
    cd.ShowColor
    Picture3.BackColor = cd.Color
errHandle:
    Err.Clear
    
End Sub

Private Sub Picture4_Click()
    On Error GoTo errHandle
    cd.ShowColor
    Picture4.BackColor = cd.Color
errHandle:
    Err.Clear

End Sub

Private Sub Picture5_Click()
    On Error GoTo errHandle
    cd.ShowColor
    Picture5.BackColor = cd.Color
errHandle:
    Err.Clear

End Sub

Private Sub Picture6_Click()
    On Error GoTo errHandle
    cd.ShowColor
    Picture6.BackColor = cd.Color
errHandle:
    Err.Clear

End Sub

Private Sub Picture7_Click()
    On Error GoTo errHandle
    cd.ShowColor
    Picture7.BackColor = cd.Color
errHandle:
    Err.Clear

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandle
    Select Case KeyCode
        Case vbKeyUp
            Text1(Index - 1).SetFocus
            Text1(Index - 1).SelStart = 0
            Text1(Index - 1).SelLength = Len(Text1(Index - 1).Text)
        Case vbKeyReturn
            Text1(Index + 1).SetFocus
            Text1(Index + 1).SelStart = 0
            Text1(Index + 1).SelLength = Len(Text1(Index + 1).Text)
        Case vbKeyDown
            Text1(Index + 1).SetFocus
            Text1(Index + 1).SelStart = 0
            Text1(Index + 1).SelLength = Len(Text1(Index + 1).Text)
    End Select
    Exit Sub
errHandle:
    OKButton.SetFocus
    Err.Clear
End Sub
