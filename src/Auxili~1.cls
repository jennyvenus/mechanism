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

 '�ֲ�������������ֵ
Private mvarA As Double '���ظ���
Private mvarT As Double '���ظ���
Private mvarX As Double '���ظ���
Private mvarY As Double '���ظ���
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
Public Property Let Y(ByVal vData As Double)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.Y = 5
    mvarY = vData
End Property
Public Property Get Y() As Double
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.Y
    Y = mvarY
End Property
Public Property Let X(ByVal vData As Double)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.X = 5
    mvarX = vData
End Property
Public Property Get X() As Double
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.X
    X = mvarX
End Property
Public Property Let T(ByVal vData As Double)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.T = 5
    mvarT = vData
End Property
Public Property Get T() As Double
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.T
    T = mvarT
End Property
Public Property Let A(ByVal vData As Double)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.A = 5
    mvarA = vData
End Property
Public Property Get A() As Double
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.A
    A = mvarA
End Property


