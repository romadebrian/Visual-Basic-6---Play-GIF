VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarhDc As Long
Private mvarTransparentColor As Long
Private mvarPicture As IPicture
Private mvarxLeft As Long
Private mvaryTop As Long
Private mvarxWidth As Long
Private mvaryHeight As Long
Private mvarLCTFlag As Boolean
Private mvarInterlaceFlag As Boolean
Private mvarSortFlag As Boolean
Private mvarLCTSize As Byte
Private mvarTCI As Byte
Private mvarDelayTime As Integer
Private mvarDisposalMethod As DisposalMethods
Private mvarUserInputFlag As Boolean
Private mvarTCFlag As Boolean

Private mvarColorMap As clsColorMap

Public Enum DisposalMethods
    dmNONE = 0 'No disposal specified. The decoder is not required
    '       to take any action.
    dmDONOT = 1 'Do not dispose. The graphic is to be left in place.
    dmRESTORE_BKGND = 2 'Restore to background color. The area used
    '      by the graphic must be restored to the background color.
    dmRESTORE_PREVIOUS = 3 'Restore to previous. The decoder is
    '      required to restore the area overwritten by the graphic
    '      with what was there prior to rendering the graphic.
End Enum
'local variable(s) to hold property value(s)
Private mvarImage As IPicture 'local copy
Public Property Set Image(ByVal vData As IPicture)
    Set mvarImage = vData
End Property


Public Property Get Image() As IPicture
    Set Image = mvarImage
End Property




Public Property Set ColorMap(ByVal vData As clsColorMap)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ColorMap = Form1
    Set mvarColorMap = vData
End Property

Public Property Get ColorMap() As clsColorMap
    Set ColorMap = mvarColorMap
End Property

Public Property Let TCFlag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TCFlag = 5
    mvarTCFlag = vData
End Property

Public Property Get TCFlag() As Boolean
    TCFlag = mvarTCFlag
End Property

Public Property Let UserInputFlag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UserInputFlag = 5
    mvarUserInputFlag = vData
End Property


Public Property Get UserInputFlag() As Boolean
    UserInputFlag = mvarUserInputFlag
End Property

Public Property Let DisposalMethod(ByVal vData As DisposalMethods)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.DisposalMethods = Form1
    mvarDisposalMethod = vData
End Property

Public Property Get DisposalMethod() As DisposalMethods
    DisposalMethod = mvarDisposalMethod
End Property

Public Property Let DelayTime(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DelayTime = 5
    mvarDelayTime = vData
End Property

Public Property Get DelayTime() As Integer
    DelayTime = mvarDelayTime
End Property

Public Property Let TCI(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TCI = 5
    mvarTCI = vData
End Property

Public Property Get TCI() As Byte
    TCI = mvarTCI
End Property

Public Property Let LCTSize(ByVal vData As Byte)
    mvarLCTSize = vData
End Property

Public Property Get LCTSize() As Byte
    LCTSize = mvarLCTSize
End Property

Public Property Let SortFlag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SortFlag = 5
    mvarSortFlag = vData
End Property

Public Property Get SortFlag() As Boolean
    SortFlag = mvarSortFlag
End Property

Public Property Let InterlaceFlag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InterlaceFlag = 5
    mvarInterlaceFlag = vData
End Property

Public Property Get InterlaceFlag() As Boolean
    InterlaceFlag = mvarInterlaceFlag
End Property

Public Property Let LCTFlag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LCTFlag = 5
    mvarLCTFlag = vData
End Property

Public Property Get LCTFlag() As Boolean
    LCTFlag = mvarLCTFlag
End Property

Public Property Let yHeight(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.yHeight = 5
    mvaryHeight = vData
End Property

Public Property Get yHeight() As Long
    yHeight = mvaryHeight
End Property

Public Property Let xWidth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xWidth = 5
    mvarxWidth = vData
End Property

Public Property Get xWidth() As Long
    xWidth = mvarxWidth
End Property

Public Property Let yTop(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.yTop = 5
    mvaryTop = vData
End Property

Public Property Get yTop() As Long
    yTop = mvaryTop
End Property

Public Property Let xLeft(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xLeft = 5
    mvarxLeft = vData
End Property

Public Property Get xLeft() As Long
    xLeft = mvarxLeft
End Property

Public Property Set Picture(ByVal vData As IPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Picture = Form1
    Set mvarPicture = vData
End Property

Public Property Get Picture() As IPicture
    Set Picture = mvarPicture
End Property

Public Property Let TransparentColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TransparentColor = 5
    mvarTransparentColor = vData
End Property

Public Property Get TransparentColor() As Long
    TransparentColor = mvarTransparentColor
End Property

Public Property Let hdc(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.hDc = 5
    mvarhDc = vData
End Property

Public Property Get hdc() As Long
    hdc = mvarhDc
End Property

Private Sub Class_Initialize()
    Set mvarColorMap = New clsColorMap
End Sub

Private Sub Class_Terminate()
    Set mvarColorMap = Nothing
End Sub
