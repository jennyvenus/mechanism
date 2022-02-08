VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hinge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"

 
Option Explicit

'local variable(s) to hold property value(s)
Private mvarCenterX As Single 'local copy
Private mvarCenterY As Single 'local copy
Private mvarJiaoDu As Double 'local copy
'local variable(s) to hold property value(s)

Public Sub Draw(ByVal Wnd As Object)
    Dim jiaodu1 As Double
    Dim jiaodu2 As Double
    Dim Ground As New UnderProp
    Wnd.Circle (mvarCenterX, mvarCenterY), 50, vbBlack
    Wnd.Circle (mvarCenterX, mvarCenterY), 20, vbBlack
    jiaodu1 = mvarJiaoDu - 25 * PI / 180
    jiaodu2 = mvarJiaoDu + 25 * PI / 180
    Wnd.Line (mvarCenterX + 40 * Cos(jiaodu1), mvarCenterY - 40 * Sin(jiaodu1))-(mvarCenterX + 400 * Cos(jiaodu1), mvarCenterY - 400 * Sin(jiaodu1)), vbBlack
    Wnd.Line (mvarCenterX + 40 * Cos(jiaodu2), mvarCenterY - 40 * Sin(jiaodu2))-(mvarCenterX + 400 * Cos(jiaodu2), mvarCenterY - 400 * Sin(jiaodu2)), vbBlack
    If Abs(mvarJiaoDu - 270 * PI / 180) < 0.1 Then
        Ground.Aclinic = True
        Ground.Length = 500 / ZoomInMultiple
        Ground.Near_Left = True
        Ground.SmallS = 5
        Ground.UpOrDown = True
        Ground.StartX = mvarCenterX - 250
        Ground.StartY = mvarCenterY + 370
        Ground.Draw Wnd
    End If
    If Abs(mvarJiaoDu - 0) < 0.1 Then
        Ground.Aclinic = False
        Ground.Length = 500 / ZoomInMultiple
        Ground.Near_Left = True
        Ground.SmallS = 5
        Ground.UpOrDown = False
        Ground.StartY = mvarCenterY - 250
        Ground.StartX = mvarCenterX + 370
        Ground.Draw Wnd
    End If
    If Abs(mvarJiaoDu - PI) < 0.1 Then
        Ground.Aclinic = False
        Ground.Length = 500 / ZoomInMultiple
        Ground.Near_Left = True
        Ground.SmallS = 5
        Ground.UpOrDown = True
        Ground.StartY = mvarCenterY - 250
        Ground.StartX = mvarCenterX - 370
        Ground.Draw Wnd
    End If
    If Abs(mvarJiaoDu - 90 * PI / 180) < 0.1 Then
        Ground.Aclinic = True
        Ground.Length = 500 / ZoomInMultiple
        Ground.Near_Left = True
        Ground.SmallS = 5
        Ground.UpOrDown = False
        Ground.StartX = mvarCenterX - 250
        Ground.StartY = mvarCenterY - 370
        Ground.Draw Wnd
    End If
End Sub

Public Property Let Angle(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.JiaoDu = 5
    mvarJiaoDu = vData
End Property


Public Property Get Angle() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.JiaoDu
    Angle = mvarJiaoDu
End Property



Public Property Let CenterY(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CenterY = 5
    mvarCenterY = vData
End Property


Public Property Get CenterY() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CenterY
    CenterY = mvarCenterY
End Property



Public Property Let CenterX(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CenterX = 5
    mvarCenterX = vData
End Property


Public Property Get CenterX() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CenterX
    CenterX = mvarCenterX
End Property



