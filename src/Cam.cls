VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

 
'local variable(s) to hold property value(s)
Private mvarBT As Double 'local copy
Private mvarR As Double 'local copy
Private mvarTT As Double 'local copy
Private mvarTT0 As Double 'local copy
Private mvarC As Double 'local copy
Private mvarGM As Double 'local copy
Private mvarBP As Double 'local copy
Private mvarTTH As Double 'local copy
Private mvarRF As Double 'local copy
Private mvarCT As Double 'local copy
Private mvarCT1 As Double 'local copy
Private mvarCT2 As Double 'local copy
Private mvarCT3 As Double 'local copy
Private mvarCT4 As Double 'local copy
Private mvarAF As Double 'local copy
Private mvarDT As Double 'local copy
Private mvarPP As Double 'local copy
Private mvarKA As Double 'local copy
Private mvarKR As Double 'local copy
Private mvarKF As Double 'local copy
Private mvarKP As Double 'local copy
Private mvarQL As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarCenterX As Double 'local copy
Private mvarCenterY As Double 'local copy
Private mvarXColor As Double 'local copy
Dim AA1 As New Staff
Dim AA2 As New Staff
Dim AA3 As New Staff
Dim t1 As Double
Dim xyn(1 To 440, 1 To 2) As Double
Dim xyw(1 To 440, 1 To 2) As Double
Dim times As Integer
Dim T As Double
Dim temp As Double
Dim FAI As Double
Dim S As Double
Dim V As Double
Dim A As Double
Dim RPC As Double
Dim BPC As Double
Dim RPC1 As Double
Dim BPC1 As Double
Dim TT01 As Double
'local variable(s) to hold property value(s)
Private mvarRR As Double 'local copy
Public Property Let RR(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RR = 5
    mvarRR = vData
End Property
Public Property Get RR() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RR
    RR = mvarRR
End Property
Public Property Let XColor(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.XColor = 5
    mvarXColor = vData
End Property
Public Property Get XColor() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.XColor
    XColor = mvarXColor
End Property
Public Property Let CenterY(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CenterY = 5
    mvarCenterY = vData
End Property
Public Property Get CenterY() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CenterY
    CenterY = mvarCenterY
End Property
Public Property Let CenterX(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CenterX = 5
    mvarCenterX = vData
End Property
Public Property Get CenterX() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CenterX
    CenterX = mvarCenterX
End Property
Public Sub Draw(ByVal Wnd As Object, WAI As Boolean, BAI As Boolean)
    AA3.X = CamCenter.X
    AA3.Y = CamCenter.Y
    times = 0
    TT01 = mvarTT0 - mvarTTH
    For T = 0 To 1 Step 0.01
        If T >= 0 And T < 1 / 8 Then
            S = S1(T)
            V = V1(T)
            A = A1(T)
        End If
        If T >= 1 / 8 And T < 7 / 8 Then
            S = S2(T)
            V = V2(T)
            A = A2(T)
        End If
        If T >= 7 / 8 And T <= 1 Then
            S = S3(T)
            V = V3(T)
            A = A3(T)
        End If
        mvarTT = mvarTTH * S
        AA1.A = mvarC
        AA1.T = mvarGM
        AA2.A = mvarRF
        AA2.T = mvarTT0 - mvarTT
        AA3.T = CA3T_1(AA1.A, AA2.A, AA1.T, AA2.T)
        AA3.A = CA3A_1(AA1.A, AA2.A, AA1.T, AA2.T, AA3.T)
        mvarR = AA3.A
        mvarBT = AA3.T
        mvarBP = mvarBT - (T * (mvarCT1 * PI / 180))
        times = times + 1
        mvarDT = mvarBT - (mvarTT0 - mvarTT + 3 * PI / 2)
        mvarAF = Atn(((mvarRF * mvarTTH * V) / (mvarR * (mvarCT1 * PI / 180) * Cos(mvarDT))) + Tan(mvarDT))
        FAI = mvarAF + mvarTT0 - mvarTT - (mvarCT1 * PI / 180) * T
        BPC = Atn((mvarR * Sin(mvarBP) + mvarRR * Sin(FAI + 0.5 * PI)) / (mvarR * Cos(mvarBP) + mvarRR * Cos(FAI + 0.5 * PI)))
        RPC = mvarR * Cos(mvarBP - BPC) + mvarRR * Cos(FAI + 0.5 * PI - BPC)
        BPC1 = Atn((mvarR * Sin(mvarBP) + mvarRR * Sin(FAI + 1.5 * PI)) / (mvarR * Cos(mvarBP) + mvarRR * Cos(FAI + 1.5 * PI)))
        RPC1 = mvarR * Cos(mvarBP - BPC1) + mvarRR * Cos(FAI + 1.5 * PI - BPC1)
        mvarKA = Cos(mvarAF) * Cos(mvarAF) * Cos(mvarAF)
        mvarKF = Sin(mvarAF - mvarDT) * Sin(mvarAF - mvarDT) * Sin(mvarAF)
        mvarKP = Cos(mvarDT) * Cos(mvarDT)
        mvarKR = -(2 * Cos(mvarAF) * Cos(mvarDT) - Cos(mvarAF) * Cos(mvarAF) * Cos(mvarAF - mvarDT))
        hf = mvarRF * mvarTTH
        hc = mvarR * (mvarCT1 * PI / 180)
        mvarPP = hf * A / (hc * hc)
        mvarQL = 1 / (1 / mvarKP * (mvarKF * mvarPP + mvarKR * 1 / mvarR + mvarKF * 1 / mvarRF))
        xyn(times, 1) = mvarCenterX + ZoomInMultiple * Cos(BPC + mvarCT) * RPC
        xyn(times, 2) = mvarCenterY - ZoomInMultiple * Sin(BPC + mvarCT) * RPC
        xyw(times, 1) = mvarCenterX + ZoomInMultiple * Cos(BPC1 + mvarCT) * RPC1
        xyw(times, 2) = mvarCenterY - ZoomInMultiple * Sin(BPC1 + mvarCT) * RPC1
    Next T
    For T = 0 To 1 Step 0.01
        times = times + 1
        xyn(times, 1) = mvarCenterX + ZoomInMultiple * Cos(BPC + mvarCT - T * mvarCT2 * PI / 180) * RPC
        xyn(times, 2) = mvarCenterY - ZoomInMultiple * Sin(BPC + mvarCT - T * mvarCT2 * PI / 180) * RPC
        xyw(times, 1) = mvarCenterX + ZoomInMultiple * Cos(BPC1 + mvarCT - T * mvarCT2 * PI / 180) * RPC1
        xyw(times, 2) = mvarCenterY - ZoomInMultiple * Sin(BPC1 + mvarCT - T * mvarCT2 * PI / 180) * RPC1
    Next T
    For T = 0 To 1 Step 0.01
        If T >= 0 And T < 1 / 8 Then
            S = S1(T)
            V = V1(T)
            A = A1(T)
        End If
        If T >= 1 / 8 And T < 7 / 8 Then
            S = S2(T)
            V = V2(T)
            A = A2(T)
        End If
        If T >= 7 / 8 And T <= 1 Then
            S = S3(T)
            V = V3(T)
            A = A3(T)
        End If
        mvarTT = TTH * S
        mvarBT = Atn((C * Sin(GM) + RF * Sin(TT01 + mvarTT)) / (C * Cos(GM) + RF * Cos(TT01 + mvarTT)))
        R = C * Cos(GM - mvarBT) + RF * Cos(TT01 + mvarTT - mvarBT)
        BP = mvarBT - T * (mvarCT3 * PI / 180): DT = mvarBT - (TT01 + mvarTT + 0.5 * PI)
        AF = Atn(((TTH * RF * V) / (R * (mvarCT3 * PI / 180) * Cos(DT))) + Tan(DT))
        FAI = AF + TT01 + mvarTT - (mvarCT3 * PI / 180) * T
        BPC = Atn((R * Sin(BP) + RR * Sin(FAI + 0.5 * PI)) / (R * Cos(BP) + RR * Cos(FAI + 0.5 * PI)))
        RPC = R * Cos(BP - BPC) + RR * Cos(FAI + 0.5 * PI - BPC)
        BPC1 = Atn((mvarR * Sin(mvarBP) + mvarRR * Sin(FAI + 1.5 * PI)) / (mvarR * Cos(mvarBP) + mvarRR * Cos(FAI + 1.5 * PI)))
        RPC1 = mvarR * Cos(mvarBP - BPC1) + mvarRR * Cos(FAI + 1.5 * PI - BPC1)
        KA = Cos(AF) * Cos(AF) * Cos(AF): KF = Sin(AF - DT) * Sin(AF - DT) * Sin(AF)
        KP = Cos(DT) * Cos(DT): KR = -(2 * Cos(AF) * Cos(DT) - Cos(AF) * Cos(AF) * Cos(AF - DT))
        hf = RF * TTH: hc = R * (mvarCT3 * PI / 180): PP = hf * A / (hc * hc)
        mvarQL = 1 / (1 / mvarKP * (mvarKF * mvarPP + mvarKR * 1 / mvarR + mvarKF * 1 / mvarRF))
        times = times + 1
        xyn(times, 1) = ZoomInMultiple * RPC * Cos(mvarCT + BPC - (mvarCT1 + mvarCT2) * PI / 180) + mvarCenterX
        xyn(times, 2) = -ZoomInMultiple * RPC * Sin(mvarCT + BPC - (mvarCT1 + mvarCT2) * PI / 180) + mvarCenterY
        xyw(times, 1) = ZoomInMultiple * RPC1 * Cos(mvarCT + BPC1 - (mvarCT1 + mvarCT2) * PI / 180) + mvarCenterX
        xyw(times, 2) = -ZoomInMultiple * RPC1 * Sin(mvarCT + BPC1 - (mvarCT1 + mvarCT2) * PI / 180) + mvarCenterY
    Next T
    For T = 0 To 1 Step 0.01
        times = times + 1
        xyn(times, 1) = mvarCenterX + ZoomInMultiple * Cos(BPC + mvarCT - (mvarCT1 + mvarCT2 + T * mvarCT4) * PI / 180) * RPC
        xyn(times, 2) = mvarCenterY - ZoomInMultiple * Sin(BPC + mvarCT - (mvarCT1 + mvarCT2 + T * mvarCT4) * PI / 180) * RPC
        xyw(times, 1) = mvarCenterX + ZoomInMultiple * Cos(BPC1 + mvarCT - (mvarCT1 + mvarCT2 + T * mvarCT4) * PI / 180) * RPC1
        xyw(times, 2) = mvarCenterY - ZoomInMultiple * Sin(BPC1 + mvarCT - (mvarCT1 + mvarCT2 + T * mvarCT4) * PI / 180) * RPC1
    Next T
    For times = 1 To 390
        Wnd.Line (xyn(times, 1), xyn(times, 2))-(xyn(times + 1, 1), xyn(times + 1, 2)), mvarXColor
        If WAI = True Then
            Wnd.Line (xyw(times, 1), xyw(times, 2))-(xyw(times + 1, 1), xyw(times + 1, 2)), mvarXColor
        End If
    Next times
    Wnd.Line (xyn(times, 1), xyn(times, 2))-(xyn(1, 1), xyn(1, 2)), mvarXColor
    If WAI = True Then
        Wnd.Line (xyw(times, 1), xyw(times, 2))-(xyw(1, 1), xyw(1, 2)), mvarXColor
    End If
End Sub
Public Property Let QL(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.QL = 5
    mvarQL = vData
End Property
Public Property Get QL() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.QL
    QL = mvarQL
End Property
Public Property Let KP(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KP = 5
    mvarKP = vData
End Property
Public Property Get KP() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KP
    KP = mvarKP
End Property
Public Property Let KF(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KF = 5
    mvarKF = vData
End Property
Public Property Get KF() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KF
    KF = mvarKF
End Property
Public Property Let KR(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KR = 5
    mvarKR = vData
End Property
Public Property Get KR() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KR
    KR = mvarKR
End Property
Public Property Let KA(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.KA = 5
    mvarKA = vData
End Property
Public Property Get KA() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.KA
    KA = mvarKA
End Property
Public Property Let PP(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PP = 5
    mvarPP = vData
End Property
Public Property Get PP() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PP
    PP = mvarPP
End Property
Public Property Let DT(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DT = 5
    mvarDT = vData
End Property
Public Property Get DT() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DT
    DT = mvarDT
End Property
Public Property Let AF(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AF = 5
    mvarAF = vData
End Property
Public Property Get AF() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AF
    AF = mvarAF
End Property
Public Property Let CT4(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CT4 = 5
    mvarCT4 = vData
End Property
Public Property Get CT4() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CT4
    CT4 = mvarCT4
End Property
Public Property Let CT3(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CT3 = 5
    mvarCT3 = vData
End Property
Public Property Get CT3() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CT3
    CT3 = mvarCT3
End Property
Public Property Let CT2(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CT2 = 5
    mvarCT2 = vData
End Property
Public Property Get CT2() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CT2
    CT2 = mvarCT2
End Property
Public Property Let CT1(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CT1 = 5
    mvarCT1 = vData
End Property
Public Property Get CT1() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CT1
    CT1 = mvarCT1
End Property
Public Property Let CT(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CT = 5
    mvarCT = vData
End Property
Public Property Get CT() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CT
    CT = mvarCT
End Property
Public Property Let RF(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RF = 5
    mvarRF = vData
End Property
Public Property Get RF() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RF
    RF = mvarRF
End Property
Public Property Let TTH(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TTH = 5
    mvarTTH = vData
End Property
Public Property Get TTH() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TTH
    TTH = mvarTTH
End Property
Public Property Let BP(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BP = 5
    mvarBP = vData
End Property
Public Property Get BP() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BP
    BP = mvarBP
End Property
Public Property Let GM(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GM = 5
    mvarGM = vData
End Property
Public Property Get GM() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GM
    GM = mvarGM
End Property
Public Property Let C(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.C = 5
    mvarC = vData
End Property
Public Property Get C() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.C
    C = mvarC
End Property
Public Property Let TT0(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TT0 = 5
    mvarTT0 = vData
End Property
Public Property Get TT0() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TT0
    TT0 = mvarTT0
End Property
Public Property Let TT(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TT = 5
    mvarTT = vData
End Property
Public Property Get TT() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TT
    TT = mvarTT
End Property
Public Property Let R(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.R = 5
    mvarR = vData
End Property
Public Property Get R() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.R
    R = mvarR
End Property
Public Property Let BT(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BT = 5
    mvarBT = vData
End Property
Public Property Get BT() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BT
    BT = mvarBT
End Property
