VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "机构学课程设计 JennyVenus@msn.com"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   1085
      ButtonWidth     =   1455
      ButtonHeight    =   926
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "运行"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "后退"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前进"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "照像"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "输入参数"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   4680
      Top             =   2490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4890
      Top             =   3825
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   1155
      Index           =   0
      Left            =   1860
      ScaleHeight     =   1155
      ScaleWidth      =   1740
      TabIndex        =   0
      Top             =   2370
      Width           =   1740
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    MaxSized
    Center Me
    CamCenter.x = Me.ScaleWidth / 2
    CamCenter.y = Me.ScaleHeight * 0.6
    Dim i As Integer
    For i = 1 To 120
        Load pic(i)
    Next i
    Load dlgPara
    Center dlgPara
    dlgPara.Show 1
    PowerOn = False
    Timer1.Enabled = False
    DrawWhichPic = 1
    Set pic(0).Picture = pic(DrawWhichPic).Image
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 1 To 120
        Unload pic(i)
    Next i
    Timer1.Enabled = False
    End
End Sub

Private Sub Timer1_Timer()
    If PowerOn Then
        DrawWhichPic = DrawWhichPic + 1
        If DrawWhichPic > MaxFrame Then
            DrawWhichPic = 1
        End If
        Set pic(0).Picture = pic(DrawWhichPic).Image
        #If 0 Then
        Dim S As String
        S = Trim(Str(DrawWhichPic))
        If DrawWhichPic < 10 Then
           S = "0" & S
        End If
        If DrawWhichPic < 100 Then
           S = "0" & S
        End If
        S = "机构学演示程序" & S & ".BMP"
        SavePicture pic(0).Picture, App.Path & "\" & S
        #End If
    End If
End Sub

Public Sub MaxSized()
    With Me
        .Width = 800 * 15 + Me.Width - Me.ScaleWidth 'Screen.Width
        .Height = 600 * 15 + Me.Height - Me.ScaleHeight 'Screen.Height - 500
        .pic(0).Left = 0
        .pic(0).Top = .Height - .ScaleHeight
        .pic(0).Width = .ScaleWidth
        'MsgBox .pic(0).Width / 15
        .pic(0).Height = .ScaleHeight ' - 500
        'msgbox
    End With
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
        Dim S As String
    Select Case Button.Index
        Case 1
            PowerOn = Not PowerOn
            If PowerOn Then
                Timer1.Enabled = True
                Button.ToolTipText = "停止"
                Button.Caption = "停止"
                Toolbar1.Buttons(2).Enabled = False
                Toolbar1.Buttons(3).Enabled = False
                Toolbar1.Buttons(4).Enabled = False
                Toolbar1.Buttons(5).Enabled = False
                
           Else
                Timer1.Enabled = False
                Button.ToolTipText = "运行"
                Button.Caption = "运行"
                Toolbar1.Buttons(2).Enabled = True
                Toolbar1.Buttons(3).Enabled = True
                Toolbar1.Buttons(4).Enabled = True
                Toolbar1.Buttons(5).Enabled = True
            End If
        Case 2
            DrawWhichPic = DrawWhichPic - 1
            If DrawWhichPic < 1 Then
                DrawWhichPic = MaxFrame
            End If
            Set pic(0).Picture = pic(DrawWhichPic).Image
        Case 3
            DrawWhichPic = DrawWhichPic + 1
            If DrawWhichPic > MaxFrame Then
                DrawWhichPic = 1
            End If
            Set pic(0).Picture = pic(DrawWhichPic).Image
        Case 4
            cmndlg.CancelError = True
            On Error GoTo errHandle
            cmndlg.DialogTitle = "保存当前图像"
            cmndlg.Filter = "*.BMP|24位真彩色位图"
            S = Trim(Str(DrawWhichPic))
            If DrawWhichPic < 10 Then
                S = "0" & S
            End If
            If DrawWhichPic < 100 Then
                S = "0" & S
            End If
            cmndlg.FileName = "机构学演示程序" & S & ".BMP"
            cmndlg.ShowSave
            SavePicture pic(0).Picture, cmndlg.FileName
        Case 5
            Center dlgPara
            dlgPara.Show 1
            
    End Select
    Exit Sub
errHandle:
    'MsgBox Err.Description
    Err.Clear

End Sub
