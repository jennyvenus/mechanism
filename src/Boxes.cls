VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Boxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarBoxWidth As Integer 'local copy
Private mvarBoxHeight As Integer 'local copy
Private mvarBoxLength As Integer 'local copy
Private mvarDirection As Double 'local copy
Private mvarStartX As Double 'local copy
Private mvarStartY As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarSpeed As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarJiaoDu As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarXColor As Long 'local copy
Public Property Let XColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.XColor = 5
    mvarXColor = vData
End Property


Public Property Get XColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.XColor
    XColor = mvarXColor
End Property
Public Sub Create(ByVal Wnd As Object)
    Wnd.Line (mvarStartX, mvarStartY)-(mvarStartX + ZoomInMultiple * mvarBoxWidth * Cos(mvarJiaoDu), mvarStartY - ZoomInMultiple * mvarBoxWidth * Sin(mvarJiaoDu)), mvarXColor
    Wnd.Line (mvarStartX, mvarStartY)-Step(ZoomInMultiple * mvarBoxWidth, ZoomInMultiple * mvarBoxHeight), mvarXColor, B
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



Public Property Let Speed(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Speed = 5
    mvarSpeed = vData
End Property


Public Property Get Speed() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Speed
    Speed = mvarSpeed
End Property



Public Property Let StartY(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartY = 5
    mvarStartY = vData
End Property


Public Property Get StartY() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartY
    StartY = mvarStartY
End Property



Public Property Let StartX(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartX = 5
    mvarStartX = vData
End Property


Public Property Get StartX() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartX
    StartX = mvarStartX
End Property



Public Property Let Direction(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Direction = 5
    mvarDirection = vData
End Property


Public Property Get Direction() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Direction
    Direction = mvarDirection
End Property



Public Sub Move(ByVal Wnd As Object)
    If mvarDirection = 0 Then
        Wnd.Line (mvarStartX, mvarStartY)-(mvarStartX + ZoomInMultiple * mvarBoxWidth * Cos(mvarJiaoDu), mvarStartY - ZoomInMultiple * mvarBoxWidth * Sin(mvarJiaoDu)), vbGreen
        Wnd.Line (mvarStartX, mvarStartY)-Step(ZoomInMultiple * mvarBoxWidth, ZoomInMultiple * mvarBoxHeight), vbGreen, B
        mvarStartX = mvarStartX + ZoomInMultiple * mvarSpeed
        Wnd.Line (mvarStartX, mvarStartY)-Step(ZoomInMultiple * mvarBoxWidth, ZoomInMultiple * mvarBoxHeight), vbGreen, B
        Wnd.Line (mvarStartX, mvarStartY)-(mvarStartX + ZoomInMultiple * mvarBoxWidth * Cos(mvarJiaoDu), mvarStartY - ZoomInMultiple * mvarBoxWidth * Sin(mvarJiaoDu)), vbGreen
    End If
    If mvarDirection = 1 Then
        Wnd.Line (mvarStartX, mvarStartY)-(mvarStartX + ZoomInMultiple * mvarBoxWidth * Cos(mvarJiaoDu), mvarStartY - ZoomInMultiple * mvarBoxWidth * Sin(mvarJiaoDu)), vbGreen
        Wnd.Line (mvarStartX, mvarStartY)-Step(ZoomInMultiple * mvarBoxWidth, ZoomInMultiple * mvarBoxHeight), vbGreen, B
        mvarStartY = mvarStartY - ZoomInMultiple * mvarSpeed
        Wnd.Line (mvarStartX, mvarStartY)-Step(ZoomInMultiple * mvarBoxWidth, ZoomInMultiple * mvarBoxHeight), vbGreen, B
        Wnd.Line (mvarStartX, mvarStartY)-(mvarStartX + ZoomInMultiple * mvarBoxWidth * Cos(mvarJiaoDu), mvarStartY - ZoomInMultiple * mvarBoxWidth * Sin(mvarJiaoDu)), vbGreen
    End If
End Sub

Public Property Let BoxLength(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BoxLength = 5
    mvarBoxLength = vData
End Property


Public Property Get BoxLength() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BoxLength
    BoxLength = mvarBoxLength
End Property



Public Property Let BoxHeight(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BoxHeight = 5
    mvarBoxHeight = vData
End Property


Public Property Get BoxHeight() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BoxHeight
    BoxHeight = mvarBoxHeight
End Property



Public Property Let BoxWidth(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BoxWidth = 5
    mvarBoxWidth = vData
End Property


Public Property Get BoxWidth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BoxWidth
    BoxWidth = mvarBoxWidth
End Property



