VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TuoBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"

 'local variable(s) to hold property value(s)
Private mvarX As Double 'local copy
Private mvarY As Double 'local copy
Private mvarXWidth As Double 'local copy
Private mvarXHeight As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarXColor As Double 'local copy
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



Public Property Let XHeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.XHeight = 5
    mvarXHeight = vData
End Property


Public Property Get XHeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.XHeight
    XHeight = mvarXHeight
End Property



Public Property Let XWidth(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.XWidth = 5
    mvarXWidth = vData
End Property


Public Property Get XWidth() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.XWidth
    XWidth = mvarXWidth
End Property

Public Sub Draw(ByVal Wnd As Object)
    Wnd.Line (mvarX, mvarY)-Step(ZoomInMultiple * mvarXWidth, 0), mvarXColor
    Wnd.Line (mvarX, mvarY)-Step(0, ZoomInMultiple * mvarXHeight), mvarXColor
End Sub

Public Property Let y(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    mvarY = vData
End Property


Public Property Get y() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    y = mvarY
End Property



Public Property Let x(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    mvarX = vData
End Property


Public Property Get x() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    x = mvarX
End Property



