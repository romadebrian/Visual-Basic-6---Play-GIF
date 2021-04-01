VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rbgColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarRed As Byte
Private mvarGreen As Byte
Private mvarBlue As Byte
Private mvarVBcolor As Long
Private mvarColor As Long

Public Function Color() As Long
'returns long color value for API functions
    Color = mvarRed * &H10000 + mvarGreen * &H100& + CLng(mvarBlue)
End Function

Public Function VBColor() As Long
'returns long color for VB purposes
    VBColor = RGB(mvarRed, mvarGreen, mvarBlue)
End Function

Public Property Let Blue(ByVal vData As Byte)
    mvarBlue = vData
End Property

Public Property Get Blue() As Byte
    Blue = mvarBlue
End Property

Public Property Let Green(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Green = 5
    mvarGreen = vData
End Property

Public Property Get Green() As Byte
    Green = mvarGreen
End Property



Public Property Let Red(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Red = 5
    mvarRed = vData
End Property


Public Property Get Red() As Byte
    Red = mvarRed
End Property



