VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   7776
   ClientTop       =   1656
   ClientWidth     =   5604
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   5604
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   972
      Left            =   360
      ScaleHeight     =   924
      ScaleWidth      =   1284
      TabIndex        =   4
      Top             =   480
      Width           =   1332
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   4320
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   492
      Left            =   3600
      TabIndex        =   3
      Top             =   2640
      Width           =   1212
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   372
      Left            =   1440
      TabIndex        =   2
      Top             =   3480
      Width           =   1692
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   492
      Left            =   1560
      TabIndex        =   1
      Top             =   2640
      Width           =   1452
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1212
      Left            =   2280
      ScaleHeight     =   1164
      ScaleWidth      =   1644
      TabIndex        =   0
      Top             =   240
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents GIF As clsGIF
Attribute GIF.VB_VarHelpID = -1
Dim iFrame As Integer

Private Sub Command1_Click()
    Dim itemp As Integer
    
    With CommonDialog1
        .InitDir = "C:\Documents and Settings\All Users\Documents\Smilies\icons\Extended"
        
        .Filter = "GIF89a (*.GIF)|*.GIF"
        .CancelError = True
        On Error GoTo UserCancelled
        .ShowOpen
        On Error GoTo 0
'Set Picture1.Picture = LoadPicture(.filename)
'itemp = GIF.TestFrame(.filename)
'Set Picture2.Picture = GIF.Frames(itemp).Image

    End With
'Exit Sub

    With GIF
        .LoadGIF CommonDialog1.filename
'        Picture1.Picture = LoadPicture
'        .CopyFrame 1, Picture1.hdc, 0, 0
'        Picture1.Refresh
    End With
Set Picture2.Picture = GIF.Frames(1).Image
Set Picture1.Picture = GIF.Frames(1).Picture
UserCancelled:
End Sub


Private Sub Command2_Click()
    Timer1.Interval = 10
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Command3_Click()
Timer1.Enabled = False
        With GIF
        '.LoadGIF "C:\Documents and Settings\All Users\Documents\Smilies\icons\Extended\ukliam3.gif"
'        .LoadGIF "C:\Documents and Settings\All Users\Documents\AnimatedGIFs\heli.gif"
'        .LoadGIF "C:\Documents and Settings\All Users\Documents\AnimatedGIFs\drip2.gif"
        '.LoadGIF "C:\Documents and Settings\All Users\Documents\Smilies\icons\wave.gif"
        '.LoadGIF "C:\Documents and Settings\All Users\Documents\Smilies\icons\wave.gif"
        .LoadGIF "C:\Documents and Settings\All Users\Documents\Smilies\icons\Extended\coolgleama.gif"
        .CopyFrame iFrame, Picture1.hdc, 10, 10
        iFrame = iFrame + 1
    End With

End Sub

Private Sub Form_Load()
iFrame = 1
    Set GIF = New clsGIF
'        With GIF
'        '.LoadGIF "C:\Documents and Settings\All Users\Documents\AnimatedGIFs\heli.gif"
'        .LoadGIF "C:\Documents and Settings\All Users\Documents\Smilies\icons\Extended\coolgleama.gif"
'        .CopyFrame iFrame, Picture1.hdc, 10, 10
'        Picture1.Refresh
'        iFrame = iFrame + 1
'        If iFrame > .Frames.Count Then iFrame = 1
'    End With
Timer1.Enabled = False
End Sub


Private Sub GIF_GIFError(ByVal Description As String)
    MsgBox Description
    End
End Sub

Private Sub Timer1_Timer()
    With GIF
        Picture1.Picture = LoadPicture
        .CopyFrame iFrame, Picture1.hdc, 10, 10
        Timer1.Interval = .Frames(iFrame).DelayTime
        iFrame = iFrame + 1
        If iFrame > .Frames.Count Then iFrame = 1
    End With

End Sub


