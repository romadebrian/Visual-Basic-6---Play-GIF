VERSION 5.00
Begin VB.UserControl ctlGIF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   ScaleHeight     =   2220
   ScaleWidth      =   4350
   Begin VB.Timer Timer1 
      Left            =   3840
      Top             =   360
   End
End
Attribute VB_Name = "ctlGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim WithEvents GIF As clsGIF
Attribute GIF.VB_VarHelpID = -1
'Event Declarations:
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event GIFError(ByVal Description As String)
Public Event FrameChange(ByVal Index As Integer)


'Default Property Values:
Const m_def_Speed = 10
Const m_def_Animate = False
Const m_def_BackColor = &H8000000F
Const m_def_AutoSize = False
Const m_def_BackStyle = 1
Const m_def_BorderStyle = 1

'Property Variables:
Dim m_Speed As Integer
Dim m_Animate As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_AutoSize As Boolean
Dim m_BackStyle As ucBackStyle
Dim m_BorderStyle As ucBorderStyle

Private isGIFloaded As Boolean

Public Enum ucBorderStyle
    vbBSNone
    vbFixedSingle
End Enum

Public Enum ucAppearance
    ucFlat = 0
    uc3D = 1
End Enum

Public Enum ucBackStyle
    ucTransparent = 0
    ucOpaque = 1
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function FillRect Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
) As Long

Private Declare Function SetRect Lib "user32" ( _
    lpRect As RECT, _
    ByVal xLeft As Long, _
    ByVal yTop As Long, _
    ByVal xRight As Long, _
    ByVal yBottom As Long _
) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" ( _
    ByVal crColor As Long _
) As Long

Private Sub EraseBackground()
Static initialized As Boolean
Static hbrBkgnd As Long
Static lpRect As RECT
    Dim mBackColor As Long
    If Not initialized Then
        mBackColor = ReverseRGB(GIF.Frames(1).TransparentColor)
        hbrBkgnd = CreateSolidBrush(mBackColor)
        SetRect lpRect, 0, 0, GIF.xWidth, GIF.yHeight
        initialized = True
    End If
    FillRect UserControl.hdc, lpRect, hbrBkgnd
End Sub


Public Function FrameCount() As Integer
    FrameCount = GIF.Frames.Count
End Function

Public Function LoadGIF(GIFfile As Variant) As Boolean
    If GIF.LoadGIF(GIFfile) Then
        isGIFloaded = True
        SetBackColor
        If m_AutoSize Then SizeMe
        Timer1.Interval = 10
        If m_BackStyle = ucTransparent Then
            UserControl.MaskColor = GIF.Frames(1).TransparentColor
            Set UserControl.MaskPicture = GIF.Frames(1).Picture
        End If
        Call GIF.CopyFrame(1, UserControl.hdc, 0, 0)
        LoadGIF = True
    Else
        LoadGIF = False
    End If

End Function

Private Function ReverseRGB(lColor As Long) As Long
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    Red = lColor And &HFF * &H10000
    Green = lColor And &HFF00
    Blue = lColor \ &H10000
    ReverseRGB = Red + Green + Blue
End Function

Private Sub SetBackColor()

    If m_BackStyle = ucOpaque Then
        UserControl.BackColor = m_BackColor
    ElseIf isGIFloaded Then
        UserControl.MaskColor = GIF.Frames(1).TransparentColor
    End If

End Sub


Private Sub SizeMe()
If Not isGIFloaded Then Exit Sub
    If Not m_AutoSize Then Exit Sub
    Dim BorderWidth As Integer, BorderHeight As Integer
    With UserControl
        If m_BorderStyle = vbFixedSingle Then
            BorderWidth = Extender.Width - .ScaleX(.ScaleWidth, vbPixels, UserControl.Parent.ScaleMode)
            BorderHeight = Extender.Height - .ScaleY(.ScaleHeight, vbPixels, UserControl.Parent.ScaleMode)
        End If
        Extender.Width = .ScaleX(GIF.xWidth, vbPixels, UserControl.Parent.ScaleMode) + BorderWidth
        Extender.Height = .ScaleY(GIF.yHeight, vbPixels, UserControl.Parent.ScaleMode) + BorderHeight
    End With
End Sub


Private Sub GIF_GIFError(ByVal Description As String)
    RaiseEvent GIFError(Description)
End Sub

Private Sub Timer1_Timer()
Static iFrame As Integer
    iFrame = iFrame + 1
    If iFrame > GIF.Frames.Count Then iFrame = 1
    'UserControl.MaskPicture = LoadPicture
    UserControl.Picture = LoadPicture
'    EraseBackground 'same as loadpicture
    UserControl.MaskColor = GIF.Frames(iFrame).TransparentColor
    UserControl.MaskPicture = GIF.Frames(iFrame).Picture
    Call GIF.CopyFrame(iFrame, UserControl.hdc, 0, 0)
    Timer1.Interval = GIF.Frames(iFrame).DelayTime
    RaiseEvent FrameChange(iFrame)
End Sub

Private Sub UserControl_Initialize()
    Set GIF = New clsGIF
    UserControl.ScaleMode = vbPixels
    UserControl.AutoRedraw = True
    Timer1.Enabled = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get Appearance() As ucAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As ucAppearance)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    If m_AutoSize Then SizeMe
End Property

Public Property Get BackStyle() As ucBackStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As ucBackStyle)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
    UserControl.BackStyle = New_BackStyle
End Property

Public Property Get BorderStyle() As ucBorderStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As ucBorderStyle)
    m_BorderStyle = New_BorderStyle
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_InitProperties()
    m_AutoSize = m_def_AutoSize
    m_BackStyle = m_def_BackStyle
    UserControl.BackStyle = m_BackStyle
    m_BackColor = m_def_BackColor
    m_Animate = m_def_Animate
    m_Speed = m_def_Speed
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 0)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    UserControl.BackStyle = m_BackStyle
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.BorderStyle = m_BorderStyle
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    UserControl.BackColor = m_BackColor
    m_Animate = PropBag.ReadProperty("Animate", m_def_Animate)
    m_Speed = PropBag.ReadProperty("Speed", m_def_Speed)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 0)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Animate", m_Animate, m_def_Animate)
    Call PropBag.WriteProperty("Speed", m_Speed, m_def_Speed)
End Sub

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    SetBackColor
End Property

Public Property Get Animate() As Boolean
    Animate = m_Animate
End Property

Public Property Let Animate(ByVal New_Animate As Boolean)
    m_Animate = New_Animate
    Timer1.Enabled = New_Animate
    PropertyChanged "Animate"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,10
Public Property Get Speed() As Integer
    Speed = m_Speed
End Property

Public Property Let Speed(ByVal New_Speed As Integer)
    m_Speed = New_Speed
    PropertyChanged "Speed"
    GIF.Speed = New_Speed
End Property

