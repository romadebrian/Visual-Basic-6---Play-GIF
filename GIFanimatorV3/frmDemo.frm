VERSION 5.00
Object = "*\ActlGIF\GIFanimatorV3.vbp"
Begin VB.Form frmDemo 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   6015
   ClientTop       =   4680
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   Begin VB.Timer SmilieTimer 
      Left            =   720
      Top             =   3600
   End
   Begin GIFanimatorV3.ctlGIF ctlGIF 
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
   End
   Begin VB.Timer FlightTimer 
      Left            =   1200
      Top             =   2640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   612
      Left            =   3480
      TabIndex        =   3
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Timer ButtonTimer 
      Left            =   480
      Top             =   2760
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1332
      Left            =   480
      ScaleHeight     =   1275
      ScaleWidth      =   1890
      TabIndex        =   1
      Top             =   240
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Take Off"
      Height          =   492
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Smilies"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1950
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const twoPI = 6.28318530717959
Private hDcButton As Long
Private WithEvents GIF As clsGIF
Attribute GIF.VB_VarHelpID = -1
Private Smilies(1 To 5) As clsGIF

    Dim xLeft
    Dim yTop

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Sub Lissajou(X As Long, Y As Long)
Static t As Single
Static theta
Dim W As Single
Dim H As Single

If theta = 0 Then theta = twoPI / 8
H = Me.ScaleHeight / 3
W = Me.ScaleWidth / 2.4
X = (W * Sin(t)) + W * 1.05

Y = Int(H * Sin(3 * t + theta)) + H * 1.1

t = t + 0.02
If t > twoPI Then
    t = t - twoPI
    theta = theta + 0.3
End If

End Sub

Private Sub Command1_Click()
        
    ctlGIF.Animate = True

    FlightTimer.Interval = 100
Exit Sub


'draw flight path
Dim X As Long, Y As Long
Dim x2 As Long, y2 As Long
Dim i As Integer
Me.ScaleMode = vbPixels
Lissajou X, Y
Lissajou x2, y2

Me.Line (X, Y)-(x2, y2)
For i = 1 To 2000
    Lissajou X, Y
    Me.Line -(X, Y)
Next i


End Sub


Private Sub ctlGIF_GIFError(ByVal Description As String)
    MsgBox Description
End Sub

Private Sub ctlGIF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Static mousePosX As Single, mousePosY As Single
    Dim newLeft As Long
    Dim newTop As Long
    If Button = 0 Then
        mousePosX = X
        mousePosY = Y
      ElseIf Button = 1 Then
        If (X <> mousePosX) Or (Y <> mousePosY) Then
            newLeft = ctlGIF.Left + X - mousePosX
            newTop = ctlGIF.Top + Y - mousePosY
            ctlGIF.Move newLeft, newTop
        End If
    End If

End Sub


Private Sub Form_Load()
Dim i As Integer
     Set GIF = New clsGIF
    
'    Me.ScaleMode = vbTwips
'    Set Me.Picture = LoadResPicture(101, vbResBitmap)
'    'get border sizes in twips for proper picture display
'    Dim xMargins As Long
'    Dim yMargins As Long
'    xMargins = Me.Width - Me.ScaleWidth
'    yMargins = Me.Height - Me.ScaleHeight
'    Me.Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbTwips) + yMargins
'    Me.Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbTwips) + xMargins
    Me.ScaleMode = vbPixels

'animate button
  Command2.Caption = ""
  With GIF
    .LoadGIF LoadResData(102, "CUSTOM")
    'have to guess at border widths
    Command2.Height = GIF.yHeight + 12
    Command2.Width = GIF.xWidth + 12
    ButtonTimer.Interval = 10
  End With

  hDcButton = GetDC(Command2.hWnd)
'fly helicopter
    With ctlGIF
        Dim X As Long, Y As Long
        Lissajou X, Y
        .Left = X
        .Top = Y
        .Appearance = ucFlat
        .AutoSize = True
        .BackStyle = ucTransparent
        '.BackColor = vbRed
        .BorderStyle = vbBSNone
        .Speed = 10
        .LoadGIF LoadResData(101, "CUSTOM")
    End With
    Picture1.ScaleMode = vbPixels
    Picture1.AutoRedraw = True

'put smilies into picturebox
    xLeft = Array(10, 10, 60, 90, 10, 50)
    yTop = Array(0, 10, 10, 10, 50, 50)
    For i = 1 To 5
        Set Smilies(i) = New clsGIF
        With Smilies(i)
        Debug.Print "Loading "; 100 + i
            .LoadGIF LoadResData(100 + i, "SMILIE")
            .xLeft = xLeft(i)
            .yTop = yTop(i)
        End With
    Next i
    SmilieTimer.Interval = 10

End Sub


Private Sub GIF_GIFError(ByVal Description As String)
    MsgBox "Error: " & Description
End Sub


'Timer for use in animating button
Private Sub ButtonTimer_Timer()
Static iFrame As Integer
    iFrame = iFrame + 1
    If iFrame > GIF.Frames.Count Then iFrame = 1
    
    GIF.Clear hDcButton
    GIF.CopyFrame iFrame, hDcButton, 5, 5
    ButtonTimer.Interval = GIF.Frames(iFrame).DelayTime
    
End Sub


Private Sub FlightTimer_Timer()
Dim X As Long, Y As Long
    'move control around the form
    Lissajou X, Y
    ctlGIF.Left = X
    ctlGIF.Top = Y
    
End Sub


'timer used for animating smilies
Private Sub SmilieTimer_Timer()
Static iFrame(1 To 5) As Integer
Dim i As Integer

    Picture1.Picture = LoadPicture
For i = 1 To 5
    iFrame(i) = iFrame(i) + 1
    If iFrame(i) > Smilies(i).Frames.Count Then iFrame(i) = 1
    Smilies(i).CopyFrame iFrame(i), Picture1.hdc, CLng(xLeft(i)), CLng(yTop(i))
Next i
    Picture1.Refresh
    SmilieTimer.Interval = Smilies(1).Frames(1).DelayTime
 
End Sub


