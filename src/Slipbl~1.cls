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
'�ֲ�������������ֵ
Private mvarCenterX As Double '���ظ���
Private mvarCenterY As Double '���ظ���
Private mvarXWidth As Double '���ظ���
Private mvarXHeight As Double '���ظ���
Public Property Let XHeight(ByVal vData As Double)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.XHeight = 5
    mvarXHeight = vData
End Property


Public Property Get XHeight() As Double
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.XHeight
    XHeight = mvarXHeight
End Property



Public Property Let XWidth(ByVal vData As Double)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.XWidth = 5
    mvarXWidth = vData
End Property


Public Property Get XWidth() As Double
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.XWidth
    XWidth = mvarXWidth
End Property



Public Property Let CenterY(ByVal vData As Double)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.CenterY = 5
    mvarCenterY = vData
End Property


Public Property Get CenterY() As Double
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.CenterY
    CenterY = mvarCenterY
End Property



Public Property Let CenterX(ByVal vData As Double)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.CenterX = 5
    mvarCenterX = vData
End Property


Public Property Get CenterX() As Double
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.CenterX
    CenterX = mvarCenterX
End Property




'�ֲ�������������ֵ

Public Sub Draw(ByVal Wnd As Object, ByVal XColor As Long)
    Wnd.Line (mvarCenterX - ZoomInMultiple * mvarXWidth / 2, mvarCenterY - ZoomInMultiple * mvarXHeight / 2)-Step(ZoomInMultiple * mvarXWidth, ZoomInMultiple * mvarXHeight), XColor, B
    Wnd.Circle (mvarCenterX, mvarCenterY), 40, XColor
       
End Sub