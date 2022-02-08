Attribute VB_Name = "mdlMathAndDemo"
Option Explicit
'这里定义了系统的一些常量和类的实例
Type CPoint                            '点的结构体，有X，Y两个属性
    X As Integer
    Y As Integer
End Type
Public PowerOn As Boolean
Public Angle1 As Double
Public Pause As Boolean
Public Const PI = 3.14159265358979     '数学中的π
Global Const AngleFO4J = 58.384               'FO4J的夹角
Public ZoomInMultiple As Double        '放大倍数
Public cam1 As New Cam     '凸轮类
Public cam2 As New Cam
Public CamCenter As CPoint '凸轮中心点
Public O5 As New Hinge     '铰支类
Public O4 As New Hinge
Public O As New Hinge
Public O2 As New Hinge
Public O3 As New Hinge
Public O1 As New Hinge
Public OO1 As New Staff    '杆类
Public OO2 As New Staff
Public OO3 As New Staff
Public OO4 As New Staff
Public OO5 As New Staff
Public O4F As New Staff
Public O4J As New Staff
Public O5G As New Staff
Public O5H As New Staff
Public GF As New Staff
Public OM As New Staff
Public MN As New Staff
Public O1M As New Staff
Public O1N As New Staff
Public NB As New Staff
Public MB As New Staff
Public OB As New Staff
Public O2B As New Staff
Public O2A As New Staff
Public AB As New Staff
Public BC As New Staff
Public O2C As New Staff
Public O2CP  As New Staff
Public O2AP As New Staff
Public CPBP As New Staff
Public APE As New Staff
Public O3K As New Staff
Public O3D As New Staff
Public O1B As New Staff
Public O1O2 As New Staff
Public O2E As New Staff
Public OE As New Staff
Public Ground1 As New UnderProp      '支撑面类
Public Ground2 As New UnderProp
Public Ground3 As New UnderProp
Public Ground4 As New UnderProp
Public Ground5 As New UnderProp

Public O5O4 As New Auxiliary         '辅助杆类
Public O5F As New Auxiliary

Public Handspike As New TuiBan       '推板类
Public SpinBlock1 As New SlipBlock   '滑块类
Public CSD As New ChuanSongDai       '传送带类
Public CT11 As Double                '初使角
Public CT22 As Double
Public CT33 As Double
Public TB1 As New TuoBan             '托板类
Public BOX1 As New Boxes             '纸盒类
Public BOX2 As New Boxes
Public BOX3 As New Boxes
Public Type XY
    X As Double
    Y As Double
End Type
Public Pos(1 To 125) As XY
Public DrawWhichPic As Integer
Public MaxFrame As Integer
Public maxy As Double
Public Sub Center(Wnd As Object)
    Wnd.Left = (Screen.Width - Wnd.Width) / 2
    Wnd.Top = (Screen.Height - Wnd.Height) / 2
End Sub
Public Sub RotateDemo(Wnd As Object, Angle As Double)
    OM.T = Angle + 2.2
    O1M.T = CA3T_1(OO1.A, OM.A, OO1.T + PI, OM.T)
    O1M.A = CA3A_1(OO1.A, OM.A, OO1.T + PI, OM.T, O1M.T)
    MN.T = CA5T_2(O1N.A, MN.A, O1M.A, O1M.T, False)
    O1N.T = CA6T_2(O1M.A, O1M.T, MN.A, MN.T)
    MN.X = O.CenterX + OM.A * ZoomInMultiple * Cos(OM.T)
    MN.Y = O.CenterY - OM.A * ZoomInMultiple * Sin(OM.T)
    NB.X = O1.CenterX + ZoomInMultiple * O1N.A * Cos(O1N.T)
    NB.Y = O1.CenterY - ZoomInMultiple * O1N.A * Sin(O1N.T)
    NB.T = MN.T + PI / 2
    MB.X = O.CenterX + OM.A * ZoomInMultiple * Cos(OM.T)
    MB.Y = O.CenterY - OM.A * ZoomInMultiple * Sin(OM.T)
    MB.T = MN.T + Atn(NB.A / MN.A)
    OB.T = CA3T_1(OM.A, MB.A, OM.T, MB.T)
    OB.A = CA3A_1(OM.A, MB.A, OM.T, MB.T, OB.T)
    O2B.T = CA3T_1(OO2.A, OB.A, OO2.T + PI, OB.T)
    O2B.A = CA3A_1(OO2.A, OB.A, OO2.T + PI, OB.T, O2B.T)
    AB.T = CA5T_2(O2A.A, AB.A, O2B.A, O2B.T, False)
    O2A.T = CA6T_2(O2B.A, O2B.T, AB.A, AB.T)
    AB.X = O.CenterX + OB.A * ZoomInMultiple * Cos(OB.T)
    AB.Y = O.CenterY - OB.A * ZoomInMultiple * Sin(OB.T)
    BC.X = O.CenterX + OB.A * ZoomInMultiple * Cos(OB.T)
    BC.Y = O.CenterY - OB.A * ZoomInMultiple * Sin(OB.T)
     
    BC.T = CA5T_2(O2C.A, BC.A, O2B.A, O2B.T, True)
    O2C.T = CA6T_2(O2B.A, O2B.T, BC.A, BC.T)
    O2CP.T = O2C.T + PI
    O2AP.T = O2A.T + PI
    CPBP.T = BC.T
    CPBP.X = O2.CenterX + ZoomInMultiple * O2CP.A * Cos(O2CP.T)
    CPBP.Y = O2.CenterY - ZoomInMultiple * O2CP.A * Sin(O2CP.T)
    O2AP.T = O2A.T + PI
    APE.T = AB.T
    APE.X = O2.CenterX + ZoomInMultiple * O2AP.A * Cos(O2AP.T)
    APE.Y = O2.CenterY - ZoomInMultiple * O2AP.A * Sin(O2AP.T)
    
    OM.V = 0
    O1M.V = CA3V_1(OO1.A, OM.A, O1M.A, OO1.T, OM.T, O1M.T, OO1.A, OM.A, OO1.P, OM.P)
    O1M.P = CA3P_1(OO1.A, OM.A, O1M.A, OO1.T, OM.T, O1M.T, OO1.A, OM.A, OO1.P, OM.P)
    MN.V = 0
    O1N.V = 0
    NB.V = 0
    MN.P = CA5P_2(O1M.A, MN.A, O1N.A, O1M.T, MN.T, O1N.T, O1M.V, MN.V, O1N.V, O1M.P)
    O1N.P = CA5P_2(O1M.A, MN.A, O1N.A, O1M.T, MN.T, O1N.T, O1M.V, MN.V, O1N.V, O1M.P)
    O1B.T = CA3T_1(O1N.A, NB.A, O1N.T, NB.T)
    O1B.A = CA3A_1(O1N.A, NB.A, O1N.T, NB.T, O1B.T)
    
    O1B.V = CA3V_1(O1N.A, NB.A, O1B.A, O1N.T, NB.T, O1B.T, O1N.V, NB.V, O1N.P, NB.P)
    O1B.P = CA3P_1(O1N.A, NB.A, O1B.A, O1N.T, NB.T, O1B.T, O1N.V, NB.V, O1N.P, NB.P)
    
    O1O2.V = 0
    O1O2.P = 0
    O2B.V = CA3V_1(O1O2.A, O1B.A, O2B.A, O1O2.T, O1B.T, O2B.T, O1O2.V, O1B.V, O1O2.P, O1B.P)
    O2B.P = CA3P_1(O1O2.A, O1B.A, O2B.A, O1O2.T, O1B.T, O2B.T, O1O2.V, O1B.V, O1O2.P, O1B.P)
    O2E.V = O2B.V * -2
    O2E.P = O2B.P
    O2E.A = O2B.A * 2
    O2E.T = O2B.T + PI
    OE.T = CA3T_1(OO2.A, O2E.A, OO2.T, O2E.T)
    OE.A = CA3A_1(OO2.A, O2E.A, OO2.T, O2E.T, OE.T)
    
    OE.V = CA3V_1(OO2.A, O2E.A, OE.A, OO2.T, O2E.T, OE.T, 0, O2E.V, 0, O2E.P)
    OE.P = CA3P_1(OO2.A, O2E.A, OE.A, OO2.T, O2E.T, OE.T, 0, O2E.V, 0, O2E.P)
    
    cam1.CT = Angle
    Angle1 = Angle + CT33
    cam2.CT = Angle1
    
    If Angle >= 0 And Angle < (cam1.CT1 * PI / 180) / 8 Then
        O3K.T = cam1.TT0 - S1(Angle / (cam1.CT1 * PI / 180)) * cam1.TTH
    End If
    If Angle >= (cam1.CT1 * PI / 180) / 8 And Angle < 7 * (cam1.CT1 * PI / 180) / 8 Then
        O3K.T = cam1.TT0 - S2(Angle / (cam1.CT1 * PI / 180)) * cam1.TTH
    End If
    If Angle >= 7 * (cam1.CT1 * PI / 180) / 8 And Angle <= (cam1.CT1 * PI / 180) Then
        O3K.T = cam1.TT0 - S3(Angle / (cam1.CT1 * PI / 180)) * cam1.TTH
    End If
    
    
    If Angle >= ((cam1.CT1 + cam1.CT2) * PI / 180) And Angle < ((cam1.CT1 + cam1.CT2) * PI / 180 + cam1.CT3 * PI / 180 / 8) Then
        O3K.T = cam1.TT0 - cam1.TTH + S1((Angle - (cam1.CT1 + cam1.CT2) * PI / 180) / (cam1.CT3 * PI / 180)) * cam1.TTH
    End If
    If Angle >= ((cam1.CT1 + cam1.CT2) * PI / 180 + cam1.CT3 * PI / 180 / 8) And Angle < ((cam1.CT1 + cam1.CT2) * PI / 180 + 7 * cam1.CT3 * PI / 180 / 8) Then
        O3K.T = cam1.TT0 - cam1.TTH + S2((Angle - (cam1.CT1 + cam1.CT2) * PI / 180) / (cam1.CT3 * PI / 180)) * cam1.TTH
    End If
    If Angle >= ((cam1.CT1 + cam1.CT2) * PI / 180 + 7 * cam1.CT3 * PI / 180 / 8) And Angle < ((cam1.CT1 + cam1.CT2 + cam1.CT3) * PI / 180) Then
        O3K.T = cam1.TT0 - cam1.TTH + S3((Angle - (cam1.CT1 + cam1.CT2) * PI / 180) / (cam1.CT3 * PI / 180)) * cam1.TTH
    End If
    
    If Angle1 <= 0 Then
        Angle1 = Angle1 + 2 * PI
    End If
    If Angle1 > 2 * PI Then
        Angle1 = Angle1 - 2 * PI
    End If
    
    If Angle1 >= 0 And Angle1 < (cam2.CT1 * PI / 180) / 8 Then
        O4J.T = cam2.TT0 - S1(Angle1 / (cam2.CT1 * PI / 180)) * cam2.TTH
    End If
    If Angle1 >= ((cam2.CT1 * PI / 180) / 8) And Angle1 < (7 * (cam2.CT1 * PI / 180) / 8) Then
        O4J.T = cam2.TT0 - S2(Angle1 / (cam2.CT1 * PI / 180)) * cam2.TTH
    End If
    If Angle1 >= 7 * (cam2.CT1 * PI / 180) / 8 And Angle1 <= (cam2.CT1 * PI / 180) Then
        O4J.T = cam2.TT0 - S3(Angle1 / (cam2.CT1 * PI / 180)) * cam2.TTH
    End If
    
    
    If Angle1 >= ((cam2.CT1 + cam2.CT2) * PI / 180) And Angle1 < ((cam2.CT1 + cam2.CT2) * PI / 180 + cam2.CT3 * PI / 180 / 8) Then
        O4J.T = cam2.TT0 - cam2.TTH + S1((Angle1 - (cam2.CT1 + cam2.CT2) * PI / 180) / (cam2.CT3 * PI / 180)) * cam2.TTH
    End If
    If Angle1 >= ((cam2.CT1 + cam2.CT2) * PI / 180 + cam2.CT3 * PI / 180 / 8) And Angle1 < ((cam2.CT1 + cam2.CT2) * PI / 180 + 7 * cam2.CT3 * PI / 180 / 8) Then
        O4J.T = cam2.TT0 - cam2.TTH + S2((Angle1 - (cam2.CT1 + cam2.CT2) * PI / 180) / (cam2.CT3 * PI / 180)) * cam2.TTH
    End If
    If Angle1 >= ((cam2.CT1 + cam2.CT2) * PI / 180 + 7 * cam2.CT3 * PI / 180 / 8) And Angle1 < ((cam2.CT1 + cam2.CT2 + cam2.CT3) * PI / 180) Then
        O4J.T = cam2.TT0 - cam2.TTH + S3((Angle1 - (cam2.CT1 + cam2.CT2) * PI / 180) / (cam2.CT3 * PI / 180)) * cam2.TTH
    End If
    
    O4F.T = O4J.T - (AngleFO4J * PI / 180)
    GF.X = O4F.X + ZoomInMultiple * O4F.A * Cos(O4F.T)
    GF.Y = O4F.Y - ZoomInMultiple * O4F.A * Sin(O4F.T)
    O5F.T = CA3T_1(O5O4.A, O4F.A, O5O4.T, O4F.T)
    O5F.A = CA3A_1(O5O4.A, O4F.A, O5O4.T, O4F.T, O5F.T)
    GF.T = CA5T_2(O5G.A, GF.A, O5F.A, O5F.T, True)
    O5G.T = CA6T_2(O5F.A, O5F.T, GF.A, GF.T)
    O5H.T = O5G.T
    Handspike.CenterX = O5.CenterX + O5H.A * Cos(O5H.T) * ZoomInMultiple + ZoomInMultiple * (Handspike.Length + Handspike.XWidth)
    SpinBlock1.CenterX = O5.CenterX + O5H.A * Cos(O5H.T) * ZoomInMultiple
    SpinBlock1.CenterY = O5.CenterY - O5H.A * Sin(O5H.T) * ZoomInMultiple
    
    O3D.T = O3K.T
    
    O4J.A = cam2.RF
    TB1.Y = O3.CenterY - ZoomInMultiple * (O3D.A * Sin(O3D.T))
    If Angle >= 0 And Angle <= cam1.CT1 * PI / 180 Then
        BOX1.StartX = TB1.X + ZoomInMultiple * 5
        BOX1.Angle = PI / 2
        BOX1.StartY = TB1.Y - ZoomInMultiple * BOX1.BoxHeight
        If TB1.Y > maxy Then
            maxy = BOX2.StartY
            'Debug.Print maxy
        End If
    End If
    If Angle1 > (cam2.CT1 + cam2.CT2 + 23) * PI / 180 And Angle1 <= (cam2.CT1 + cam1.CT2 + cam2.CT3 - 40) * PI / 180 Then
        BOX1.StartY = Ground3.StartY - ZoomInMultiple * BOX1.BoxHeight
        BOX1.StartX = Handspike.CenterX
        'If BOX1.StartY > maxy Then
        '    maxy = BOX2.StartY
        '    'Debug.Print maxy
        'End If
    End If
    If BOX1.StartX > (CSD.LeftCenterX - ZoomInMultiple * 8) Then
        BOX1.StartX = BOX1.StartX + 2 * 1.72 * Val(dlgPara.cbo.Text) / 5 * ZoomInMultiple
        BOX1.Angle = BOX1.Angle - 2 * 0.0265 * Val(dlgPara.cbo.Text) / 5
    End If
    If Angle >= 0 And Angle <= (cam1.CT1 + cam1.CT2) * PI / 180 Then
        BOX2.StartX = frmmain.pic(0).Width / 2 + ZoomInMultiple * 150 + ZoomInMultiple * Angle * 30
''''''        BOX2.StartX = 4800 + ZoomInMultiple * Angle * 30
        BOX2.Angle = 0
        BOX2.StartY = Ground3.StartY - ZoomInMultiple * BOX1.BoxHeight
        'If BOX2.StartY > maxy Then
        '    maxy = BOX2.StartY
        '    'Debug.Print maxy
        'End If
    Else
        BOX2.StartX = -5000
    End If
    If Angle > (cam1.CT1 + cam1.CT2 + cam1.CT3) * PI / 180 And Angle < PI * 2 Then
        BOX3.StartX = TB1.X + ZoomInMultiple * 5
        'BOX3.StartY = (Angle - (cam1.CT1 + cam1.CT2 + cam1.CT3) * PI / 180) * ZoomInMultiple * 130
        BOX3.StartY = ((maxy + frmmain.pic(0).Height * 0.6) / 60) * Int((Angle * 180 / PI) - 292 + 0.5) * 0.45
        'Debug.Print ((maxy + BOX1.BoxHeight + frmmain.pic(0).Height / 2) / 60) * Int((Angle * 180 / PI) - 292 + 0.5)
''''''        BOX3.StartY = (Angle - (cam1.CT1 + cam1.CT2 + cam1.CT3) * PI / 180) * (2200 / (cam1.CT4 * PI / 180))
    Else
        BOX3.StartY = -5000
    End If
    'Debug.Print maxy
    Wnd.FillStyle = vbSolid
    Wnd.FillColor = BOX1.XColor
    BOX1.Create Wnd
    Wnd.CurrentX = BOX1.StartX + 100
    Wnd.CurrentY = BOX1.StartY + 100
    Wnd.Print "正在封盖"
    BOX2.Create Wnd
    Wnd.CurrentX = BOX2.StartX + 100
    Wnd.CurrentY = BOX2.StartY + 100
    Wnd.Print "封盖完毕"
    BOX3.Create Wnd
    Wnd.CurrentX = BOX3.StartX + 100
    Wnd.CurrentY = BOX3.StartY + 100
    Wnd.Print "尚未封盖"
    Wnd.FillStyle = vbTransparent
    cam1.Draw Wnd, False, True
    Wnd.FillColor = dlgPara.Picture2.BackColor
    Wnd.FillStyle = dlgPara.cbo1.ListIndex
    ExtFloodFill Wnd.hdc, CLng(CamCenter.X / 15 + 1), CLng(CamCenter.Y / 15 + 1), cam1.XColor, CLng(0)
    cam2.Draw Wnd, True, True
    Wnd.FillStyle = vbSolid
    Wnd.FillColor = vbBlue
    Wnd.Circle (O4.CenterX + ZoomInMultiple * cam2.RF * Cos(O4J.T), O4.CenterY - ZoomInMultiple * cam2.RF * Sin(O4J.T)), ZoomInMultiple * cam2.RR, vbBlue
    Wnd.FillColor = vbGreen
    Wnd.Circle (O3.CenterX + ZoomInMultiple * O3K.A * Cos(O3K.T), O3.CenterY - ZoomInMultiple * O3K.A * Sin(O3K.T)), ZoomInMultiple * cam1.RR, cam1.XColor
    Wnd.Circle (O3.CenterX + ZoomInMultiple * O3D.A * Cos(O3D.T), O3.CenterY - ZoomInMultiple * O3D.A * Sin(O3D.T)), ZoomInMultiple * 10, cam1.XColor
    Wnd.FillStyle = vbTransparent
    Handspike.Draw Wnd, vbBlack
    SpinBlock1.Draw Wnd, vbBlack
    Ground1.Draw Wnd
    Ground2.Draw Wnd
    Ground3.Draw Wnd
    Ground4.Draw Wnd
    Ground5.Draw Wnd
    CSD.Draw Wnd
    
    TB1.Draw Wnd
    O3D.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "D"
    OM.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "M"
    MN.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "N"
    O1N.Draw Wnd
    NB.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "B"
    O2A.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "A"
    AB.Draw Wnd
    O2C.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "C"
    BC.Draw Wnd
    O2CP.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "C'"
    CPBP.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "B'"
    O2AP.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "A'"
    APE.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "E"
    O4J.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "J"
    O4F.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "F"
    GF.Draw Wnd
    O5H.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "H"
    O3K.Draw Wnd
    Wnd.Circle (Wnd.CurrentX, Wnd.CurrentY), 45, vbRed
    Wnd.CurrentX = Wnd.CurrentX + 100
    Wnd.Print "K"
    O.Draw Wnd
    Wnd.Print "O"
    O1.Draw Wnd
    Wnd.Print "O1"
    O2.Draw Wnd
    Wnd.Print "O2"
    O3.Draw Wnd
    Wnd.Print "O3"
    O4.Draw Wnd
    Wnd.Print "O4"
    O5.Draw Wnd
    Wnd.Print "O5"
    
    Wnd.PSet (O2.CenterX + ZoomInMultiple * O2AP.A * Cos(O2AP.T) + ZoomInMultiple * APE.A * Cos(APE.T), O2.CenterY - ZoomInMultiple * O2AP.A * Sin(O2AP.T) - ZoomInMultiple * APE.A * Sin(APE.T))
    Pos(Int(Angle * 180 / PI / Val(dlgPara.cbo.Text) + 1)).X = O2.CenterX + ZoomInMultiple * O2AP.A * Cos(O2AP.T) + ZoomInMultiple * APE.A * Cos(APE.T)
    Pos(Int(Angle * 180 / PI / Val(dlgPara.cbo.Text) + 1)).Y = O2.CenterY - ZoomInMultiple * O2AP.A * Sin(O2AP.T) - ZoomInMultiple * APE.A * Sin(APE.T)
    
    Dim i As Integer
    Wnd.CurrentX = CSng(Pos(1).X)
    Wnd.CurrentY = CSng(Pos(1).Y)
    Wnd.PSet (Wnd.CurrentX, Wnd.CurrentY), vbRed
    i = 1
    Do Until i > 120 Or Pos(i).X = 0
        Wnd.Line -(Pos(i).X, Pos(i).Y), vbRed
        i = i + 1
    Loop
    If Pos(MaxFrame).X <> 0 Then
        Wnd.Line -(Pos(1).X, Pos(1).Y), vbRed
    End If
End Sub
Public Function CA3A_1(ByVal A1_A As Single, ByVal A2_A As Single, ByVal A1_T As Single, ByVal A2_T As Single, ByVal A3_T As Single) As Single
    CA3A_1 = A1_A * Cos(A1_T - A3_T) + A2_A * Cos(A2_T - A3_T)
End Function
Public Function CA3T_1(ByVal A1_A As Single, ByVal A2_A As Single, ByVal A1_T As Single, ByVal A2_T As Single) As Single
    Dim temp As Single
    Dim Sine As Single
    Dim Cosine As Single
    Sine = A1_A * Sin(A1_T) + A2_A * Sin(A2_T)
    Cosine = A1_A * Cos(A1_T) + A2_A * Cos(A2_T)
    If Sine > 0 And Cosine > 0 Then
        temp = Atn(Sine / Cosine)
    End If
    If Sine > 0 And Cosine < 0 Then
        temp = (Atn(Sine / Cosine) + PI)
    End If
    If Sine < 0 And Cosine < 0 Then
        temp = (Atn(Sine / Cosine) + PI)
    End If
    If Sine < 0 And Cosine > 0 Then
        temp = (Atn(Sine / Cosine) + 2 * PI)
    End If
    CA3T_1 = temp
End Function
Public Function CA5T_2(ByVal a6m As Single, ByVal a5m As Single, ByVal a4m As Single, ByVal a4f As Single, ByVal shang As Boolean) As Single
    Dim temp As Single
    Dim sine1 As Single
    Dim cosine1 As Single
    cosine1 = (a6m ^ 2 - a4m ^ 2 - a5m ^ 2) / (2 * a4m * a5m)
    sine1 = Sqr(1 - cosine1 ^ 2)
    If Not shang Then
        sine1 = -sine1
    End If
    If sine1 > 0 And cosine1 > 0 Then
        temp = a4f + (Atn(sine1 / cosine1))
    End If
    If sine1 > 0 And cosine1 < 0 Then
        temp = a4f + ((Atn(sine1 / cosine1) + PI))
    End If
    If sine1 < 0 And cosine1 < 0 Then
        temp = a4f + ((Atn(sine1 / cosine1) + PI))
    End If
    If sine1 < 0 And cosine1 > 0 Then
        temp = a4f + ((Atn(sine1 / cosine1) + 2 * PI))
    End If
    CA5T_2 = temp
End Function
Public Function CA6T_2(ByVal a4m As Single, ByVal a4f As Single, ByVal a5m As Single, ByVal a5f As Single) As Single
    Dim temp As Single
    Dim sine1 As Single
    Dim cosine1 As Single
    sine1 = a4m * Sin(a4f) + a5m * Sin(a5f)
    cosine1 = a4m * Cos(a4f) + a5m * Cos(a5f)
    
    If sine1 > 0 And cosine1 > 0 Then
        temp = (Atn(sine1 / cosine1))
    End If
    If sine1 > 0 And cosine1 < 0 Then
        temp = ((Atn(sine1 / cosine1) + PI))
    End If
    If sine1 < 0 And cosine1 < 0 Then
        temp = ((Atn(sine1 / cosine1) + PI))
    End If
    If sine1 < 0 And cosine1 > 0 Then
        temp = ((Atn(sine1 / cosine1) + 2 * PI))
    End If
    CA6T_2 = temp
End Function
Public Function S1(ByVal T As Double) As Double
    Dim temp As Double
    temp = (4 * PI * T - Sin(4 * PI * T)) / (4 * (4 + PI))
    S1 = temp
End Function
Public Function S2(ByVal T As Double) As Double
    Dim temp As Double
    temp = (4 * PI * T - 9 * Sin(4 / 3 * PI * T + PI / 3) + 8) / (4 * (4 + PI))
    S2 = temp
End Function
Public Function S3(ByVal T As Double) As Double
    Dim temp As Double
    temp = (4 * PI * T - Sin(4 * PI * T - 2 * PI) + 16) / (4 * (4 + PI))
    S3 = temp
End Function
Public Function V1(ByVal T As Double) As Double
    Dim temp As Double
    temp = (PI / (4 + PI)) * (1 - Cos(4 * PI * T))
    V1 = temp
End Function
Public Function V2(ByVal T As Double) As Double
    Dim temp As Double
    temp = (PI / (4 + PI)) * (1 - 3 * Cos((4 / 3) * PI * T + PI / 3))
    V2 = temp
End Function
Public Function V3(ByVal T As Double) As Double
    Dim temp As Double
    temp = (PI / (4 + PI)) * (1 - Cos(4 * PI * T - 2 * PI))
    V3 = temp
End Function
Public Function A1(ByVal T As Double) As Double
    Dim temp As Double
    temp = (4 * PI * PI / (4 + PI)) * Sin(4 * PI * T)
    A1 = temp
End Function
Public Function A2(ByVal T As Double) As Double
    Dim temp As Double
    temp = (4 * PI * PI / (4 + PI)) * Sin((4 / 3) * PI * T + PI / 3)
    A2 = temp
End Function
Public Function A3(ByVal T As Double) As Double
    Dim temp As Double
    temp = (4 * PI * PI / (4 + PI)) * Sin(4 * PI * T - PI * 2)
    A3 = temp
End Function
Public Function CA3V_1(ByVal A1_A As Double, ByVal A2_A As Double, ByVal A3_A As Double, ByVal A1_T As Double, ByVal A2_T As Double, ByVal A3_T As Double, ByVal A1_V As Double, ByVal A2_V As Double, ByVal A1_P As Double, ByVal A2_P As Double) As Double
    Dim temp As Double
    temp = A1_V * Cos(A1_T - A3_T) - A1_A * A1_P * Sin(A1_T - A3_T) + A2_V * Cos(A2_T - A3_T) - A2_A * A2_P * Sin(A2_T - A3_T)
    CA3V_1 = temp
End Function
Public Function CA3P_1(ByVal A1_A As Double, ByVal A2_A As Double, ByVal A3_A As Double, ByVal A1_T As Double, ByVal A2_T As Double, ByVal A3_T As Double, ByVal A1_V As Double, ByVal A2_V As Double, ByVal A1_P As Double, ByVal A2_P As Double) As Double
    Dim temp As Double
    temp = (A1_V * Sin(A1_T - A3_T) + A1_A * A1_P * Cos(A1_T - A3_T) + A2_V * Sin(A2_T - A3_T) + A2_A * A2_P * Cos(A2_T - A3_T)) / A3_A
    CA3P_1 = temp
End Function
Public Function CA5P_2(ByVal A4_A As Double, ByVal A5_A As Double, ByVal A6_A As Double, ByVal A4_T As Double, ByVal A5_T As Double, ByVal A6_T As Double, ByVal A4_V As Double, ByVal A5_V As Double, ByVal A6_V As Double, ByVal A4_P As Double) As Double
    Dim temp As Double
    temp = (A4_V * Cos(A4_T - A6_T) - A4_A * A4_P * Sin(A4_T - A6_T) + A5_V * Cos(A5_T - A6_T) - A6_V) / (A5_A * Sin(A5_T - A6_T))
    CA5P_2 = temp
End Function
Public Function CA6P_2(ByVal A4_A As Double, ByVal A5_A As Double, ByVal A6_A As Double, ByVal A4_T As Double, ByVal A5_T As Double, ByVal A6_T As Double, ByVal A4_V As Double, ByVal A5_V As Double, ByVal A6_V As Double, ByVal A4_P As Double) As Double
    Dim temp As Double
    temp = (A4_V * Cos(A4_T - A5_T) - A4_A * A4_P * Sin(A4_T - A5_T) - A6_V * Cos(A5_T - A6_T) + A5_V) / (A6_A * Sin(A5_T - A6_T))
    CA6P_2 = temp
End Function
