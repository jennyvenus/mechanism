VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChuanSongDai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

 'local variable(s) to hold property value(s)
Private mvarLength As Double 'local copy
Private mvarRadius As Double 'local copy
Private mvarLeftCenterX As Double 'local copy
Private mvarLeftCenterY As Double 'local copy
Private mvarXColor As Double 'local copy
Public Sub Draw(ByVal Wnd As Object)
    Wnd.Line (mvarLeftCenterX, mvarLeftCenterY - ZoomInMultiple * mvarRadius)-(mvarLeftCenterX + ZoomInMultiple * mvarLength, mvarLeftCenterY - ZoomInMultiple * mvarRadius), mvarXColor
    Wnd.Line (mvarLeftCenterX, mvarLeftCenterY - ZoomInMultiple * (mvarRadius - 10))-(mvarLeftCenterX + ZoomInMultiple * mvarLength, mvarLeftCenterY - ZoomInMultiple * (mvarRadius - 10)), mvarXColor
    Wnd.Line (mvarLeftCenterX, mvarLeftCenterY + ZoomInMultiple * mvarRadius)-(mvarLeftCenterX + ZoomInMultiple * mvarLength, mvarLeftCenterY + ZoomInMultiple * mvarRadius), mvarXColor
    Wnd.Line (mvarLeftCenterX, mvarLeftCenterY + ZoomInMultiple * (mvarRadius - 10))-(mvarLeftCenterX + ZoomInMultiple * mvarLength, mvarLeftCenterY + ZoomInMultiple * (mvarRadius - 10)), mvarXColor
    Wnd.Circle (mvarLeftCenterX, mvarLeftCenterY), ZoomInMultiple * mvarRadius, mvarXColor, PI / 2, 3 * PI / 2
    Wnd.Circle (mvarLeftCenterX, mvarLeftCenterY), ZoomInMultiple * (mvarRadius - 10), mvarXColor, PI / 2, 3 * PI / 2
    Wnd.Circle (mvarLeftCenterX + ZoomInMultiple * mvarLength, mvarLeftCenterY), ZoomInMultiple * mvarRadius, mvarXColor, 3 * PI / 2, PI / 2
    Wnd.Circle (mvarLeftCenterX + ZoomInMultiple * mvarLength, mvarLeftCenterY), ZoomInMultiple * (mvarRadius - 10), mvarXColor, 3 * PI / 2, PI / 2
End Sub

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



Public Property Let LeftCenterY(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LeftCenterY = 5
    mvarLeftCenterY = vData
End Property


Public Property Get LeftCenterY() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LeftCenterY
    LeftCenterY = mvarLeftCenterY
End Property



Public Property Let LeftCenterX(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LeftCenterX = 5
    mvarLeftCenterX = vData
End Property


Public Property Get LeftCenterX() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LeftCenterX
    LeftCenterX = mvarLeftCenterX
End Property



Public Property Let Radius(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Radius = 5
    mvarRadius = vData
End Property


Public Property Get Radius() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Radius
    Radius = mvarRadius
End Property



Public Property Let Length(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Length = 5
    mvarLength = vData
End Property


Public Property Get Length() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Length
    Length = mvarLength
End Property



