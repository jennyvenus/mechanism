VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnderProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
 
'local variable(s) to hold property value(s)
Private mvarStartX As Single 'local copy
Private mvarStartY As Single 'local copy
Private mvarLength As Single 'local copy
Private mvarNear_Left As Boolean 'local copy
Private mvarAclinic As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarSmallS As Byte 'local copy
'local variable(s) to hold property value(s)
Private mvarUpOrDown As Boolean 'local copy
Public Property Let UpOrDown(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UpOrDown = 5
    mvarUpOrDown = vData
End Property


Public Property Get UpOrDown() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UpOrDown
    UpOrDown = mvarUpOrDown
End Property



Public Property Let SmallS(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SmallS = 5
    mvarSmallS = vData
End Property


Public Property Get SmallS() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SmallS
    SmallS = mvarSmallS
End Property



Public Sub Draw(ByVal Wnd As Object)
    Dim i As Integer
    If mvarSmallS <= 1 Then
        Exit Sub
    End If
    If mvarUpOrDown = True Then
        If mvarAclinic = True Then
            Wnd.Line (mvarStartX, mvarStartY)-(mvarStartX + ZoomInMultiple * mvarLength, mvarStartY), vbBlack
            If Near_Left = True Then
                For i = mvarStartX + ZoomInMultiple * mvarSmallS / 2 To mvarStartX + ZoomInMultiple * mvarLength - ZoomInMultiple * mvarSmallS / 2 Step ZoomInMultiple * mvarSmallS
                    Wnd.Line (i, mvarStartY)-(i - 65, mvarStartY + 65), vbBlack
                Next i
            Else
                For i = mvarStartX + ZoomInMultiple * mvarSmallS / 2 To mvarStartX + ZoomInMultiple * mvarLength - ZoomInMultiple * mvarSmallS / 2 Step ZoomInMultiple * mvarSmallS
                    Wnd.Line (i, mvarStartY)-(i + 65, mvarStartY + 65), vbBlack
                Next i
            End If
        Else
            Wnd.Line (mvarStartX, mvarStartY)-(mvarStartX, mvarStartY + ZoomInMultiple * mvarLength), vbBlack
            If Near_Left = True Then
                For i = mvarStartY + ZoomInMultiple * mvarSmallS / 2 To mvarStartY + ZoomInMultiple * mvarLength - ZoomInMultiple * mvarSmallS / 2 Step ZoomInMultiple * mvarSmallS
                    Wnd.Line (mvarStartX, i)-(mvarStartX - 65, i - 65), vbBlack
                Next i
            Else
                For i = mvarStartY + ZoomInMultiple * mvarSmallS / 2 To mvarStartY + ZoomInMultiple * mvarLength - ZoomInMultiple * mvarSmallS / 2 Step ZoomInMultiple * mvarSmallS
                    Wnd.Line (mvarStartX, i)-(mvarStartX - 65, i + 65), vbBlack
                Next i
            End If
        End If
    Else
        If mvarAclinic = True Then
            Wnd.Line (mvarStartX, mvarStartY)-(mvarStartX + ZoomInMultiple * mvarLength, mvarStartY), vbBlack
            If Near_Left = True Then
                For i = mvarStartX + ZoomInMultiple * mvarSmallS / 2 To mvarStartX + ZoomInMultiple * mvarLength - ZoomInMultiple * mvarSmallS / 2 Step ZoomInMultiple * mvarSmallS
                    Wnd.Line (i, mvarStartY)-(i + 65, mvarStartY - 65), vbBlack
                Next i
            Else
                For i = mvarStartX + ZoomInMultiple * mvarSmallS / 2 To mvarStartX + ZoomInMultiple * mvarLength - ZoomInMultiple * mvarSmallS / 2 Step ZoomInMultiple * mvarSmallS
                    Wnd.Line (i, mvarStartY)-(i - 65, mvarStartY - 65), vbBlack
                Next i
            End If
        Else
            Wnd.Line (mvarStartX, mvarStartY)-(mvarStartX, mvarStartY + ZoomInMultiple * mvarLength), vbBlack
            If Near_Left = True Then
                For i = mvarStartY + ZoomInMultiple * mvarSmallS / 2 To mvarStartY + ZoomInMultiple * mvarLength - ZoomInMultiple * mvarSmallS / 2 Step ZoomInMultiple * mvarSmallS
                    Wnd.Line (mvarStartX, i)-(mvarStartX + 65, i + 65), vbBlack
                Next i
            Else
                For i = mvarStartY + ZoomInMultiple * mvarSmallS / 2 To mvarStartY + ZoomInMultiple * mvarLength - ZoomInMultiple * mvarSmallS / 2 Step ZoomInMultiple * mvarSmallS
                    Wnd.Line (mvarStartX, i)-(mvarStartX + 65, i - 65), vbBlack
                Next i
            End If
        End If
    End If
End Sub

Public Property Let Aclinic(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Aclinic = 5
    mvarAclinic = vData
End Property


Public Property Get Aclinic() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Aclinic
    Aclinic = mvarAclinic
End Property



Public Property Let Near_Left(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Near_Left = 5
    mvarNear_Left = vData
End Property


Public Property Get Near_Left() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Near_Left
    Near_Left = mvarNear_Left
End Property



Public Property Let Length(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Length = 5
    mvarLength = vData
End Property


Public Property Get Length() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Length
    Length = mvarLength
End Property



Public Property Let StartY(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartY = 5
    mvarStartY = vData
End Property


Public Property Get StartY() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartY
    StartY = mvarStartY
End Property



Public Property Let StartX(ByVal vData As Single)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StartX = 5
    mvarStartX = vData
End Property


Public Property Get StartX() As Single
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartX
    StartX = mvarStartX
End Property



