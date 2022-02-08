VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlipBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

 
  
Option Explicit
'局部变量保存属性值
Private mvarCenterX As Double '本地复本
Private mvarCenterY As Double '本地复本
Private mvarXWidth As Double '本地复本
Private mvarXHeight As Double '本地复本
Public Property Let XHeight(ByVal vData As Double)
'当给属性赋值时在参数左边使用。
'Syntax: X.XHeight = 5
    mvarXHeight = vData
End Property


Public Property Get XHeight() As Double
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.XHeight
    XHeight = mvarXHeight
End Property



Public Property Let XWidth(ByVal vData As Double)
'当给属性赋值时在参数左边使用。
'Syntax: X.XWidth = 5
    mvarXWidth = vData
End Property


Public Property Get XWidth() As Double
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.XWidth
    XWidth = mvarXWidth
End Property



Public Property Let CenterY(ByVal vData As Double)
'当给属性赋值时在参数左边使用。
'Syntax: X.CenterY = 5
    mvarCenterY = vData
End Property


Public Property Get CenterY() As Double
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.CenterY
    CenterY = mvarCenterY
End Property



Public Property Let CenterX(ByVal vData As Double)
'当给属性赋值时在参数左边使用。
'Syntax: X.CenterX = 5
    mvarCenterX = vData
End Property


Public Property Get CenterX() As Double
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.CenterX
    CenterX = mvarCenterX
End Property




'局部变量保存属性值

Public Sub Draw(ByVal Wnd As Object, ByVal XColor As Long)
    Wnd.Line (mvarCenterX - ZoomInMultiple * mvarXWidth / 2, mvarCenterY - ZoomInMultiple * mvarXHeight / 2)-Step(ZoomInMultiple * mvarXWidth, ZoomInMultiple * mvarXHeight), XColor, B
    Wnd.Circle (mvarCenterX, mvarCenterY), 40, XColor
       
End Sub
