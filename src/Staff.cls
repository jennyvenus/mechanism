VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"

    



Option Explicit

'local variable(s) to hold property value(s)
Private mvarX As Single 'local copy
Private mvarY As Single 'local copy
Private mvarA As Double 'local copy
Private mvarT As Double 'local copy
Private mvarXColor As Long 'local copy
'��������ֵ�ľֲ�����
Private mvarV As Double '�ֲ�����
Private mvarP As Double '�ֲ�����
Public Property Let P(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.P = 5
    mvarP = vData
End Property


Public Property Get P() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.P
    P = mvarP
End Property



Public Property Let V(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.V = 5
    mvarV = vData
End Property


Public Property Get V() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.V
    V = mvarV
End Property



Public Sub Draw(ByVal Wnd As Object)
    Dim x1 As Single
    Dim y1 As Single
    x1 = mvarX + ZoomInMultiple * mvarA * Cos(mvarT)
    y1 = mvarY - ZoomInMultiple * mvarA * Sin(mvarT)
    Wnd.Line (mvarX, mvarY)-(x1, y1), mvarXColor
End Sub

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



Public Property Let T(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.T = 5
    mvarT = vData
End Property


Public Property Get T() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.T
    T = mvarT
End Property



Public Property Let A(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.A = 5
    mvarA = vData
End Property


Public Property Get A() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.A
    A = mvarA
End Property



Public Property Let y(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Y = 5
    mvarY = vData
End Property


Public Property Get y() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Y
    y = mvarY
End Property



Public Property Let x(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.X = 5
    mvarX = vData
End Property


Public Property Get x() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.X
    x = mvarX
End Property


