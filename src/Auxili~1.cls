VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Auxiliary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

 '局部变量保存属性值
Private mvarA As Double '本地复本
Private mvarT As Double '本地复本
Private mvarX As Double '本地复本
Private mvarY As Double '本地复本
'保持属性值的局部变量
Private mvarV As Double '局部复制
Private mvarP As Double '局部复制
Public Property Let P(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.P = 5
    mvarP = vData
End Property
Public Property Get P() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.P
    P = mvarP
End Property
Public Property Let V(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.V = 5
    mvarV = vData
End Property
Public Property Get V() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.V
    V = mvarV
End Property
Public Property Let Y(ByVal vData As Double)
'当给属性赋值时在参数左边使用。
'Syntax: X.Y = 5
    mvarY = vData
End Property
Public Property Get Y() As Double
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.Y
    Y = mvarY
End Property
Public Property Let X(ByVal vData As Double)
'当给属性赋值时在参数左边使用。
'Syntax: X.X = 5
    mvarX = vData
End Property
Public Property Get X() As Double
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.X
    X = mvarX
End Property
Public Property Let T(ByVal vData As Double)
'当给属性赋值时在参数左边使用。
'Syntax: X.T = 5
    mvarT = vData
End Property
Public Property Get T() As Double
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.T
    T = mvarT
End Property
Public Property Let A(ByVal vData As Double)
'当给属性赋值时在参数左边使用。
'Syntax: X.A = 5
    mvarA = vData
End Property
Public Property Get A() As Double
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.A
    A = mvarA
End Property



