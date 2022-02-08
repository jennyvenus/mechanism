VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TuiBan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"

 
  

 

'local variable(s) to hold property value(s)
Private mvarCenterX As Double 'local copy
Private mvarCenterY As Double 'local copy
Private mvarLength As Variant 'local copy
Private mvarXWidth As Variant 'local copy
Private mvarXHeight As Variant 'local copy
'局部变量保存属性值
Private mvarHangleLength As Double '本地复本
Public Property Let HandleLength(ByVal vData As Double)
'当给属性赋值时在参数左边使用。
'Syntax: X.HangleLength = 5
    mvarHangleLength = vData
End Property


Public Property Get HandleLength() As Double
'当检索属性值时在参数右边使用。
'Syntax: Debug.Print X.HangleLength
    HandleLength = mvarHangleLength
End Property



Public Sub Draw(ByVal Wnd As Object, XColor As Long)
    Wnd.Line (mvarCenterX - ZoomInMultiple * mvarXWidth, mvarCenterY)-(mvarCenterX - ZoomInMultiple * (mvarXWidth + mvarLength), mvarCenterY), XColor
    Wnd.Line (mvarCenterX - ZoomInMultiple * mvarXWidth, mvarCenterY - ZoomInMultiple * mvarXHeight)-(mvarCenterX, mvarCenterY + ZoomInMultiple * mvarXHeight), XColor, B
    Wnd.Line (mvarCenterX - ZoomInMultiple * (mvarXWidth + mvarLength), mvarCenterY)-(mvarCenterX - ZoomInMultiple * (mvarXWidth + mvarLength), mvarCenterY + ZoomInMultiple * mvarHangleLength)
End Sub

Public Property Let XHeight(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.XHeight = 5
    mvarXHeight = vData
End Property


Public Property Set XHeight(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.XHeight = Form1
    Set mvarXHeight = vData
End Property


Public Property Get XHeight() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.XHeight
    If IsObject(mvarXHeight) Then
        Set XHeight = mvarXHeight
    Else
        XHeight = mvarXHeight
    End If
End Property



Public Property Let XWidth(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.XWidth = 5
    mvarXWidth = vData
End Property


Public Property Set XWidth(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.XWidth = Form1
    Set mvarXWidth = vData
End Property


Public Property Get XWidth() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.XWidth
    If IsObject(mvarXWidth) Then
        Set XWidth = mvarXWidth
    Else
        XWidth = mvarXWidth
    End If
End Property



Public Property Let Length(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Length = 5
    mvarLength = vData
End Property


Public Property Set Length(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Length = Form1
    Set mvarLength = vData
End Property


Public Property Get Length() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Length
    If IsObject(mvarLength) Then
        Set Length = mvarLength
    Else
        Length = mvarLength
    End If
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



