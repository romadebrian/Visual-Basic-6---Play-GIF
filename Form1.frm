VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "AniGIF.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form PlayGIF 
   Caption         =   "Play GIF"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   960
      Top             =   3480
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   3255
      Left            =   3600
      TabIndex        =   4
      Top             =   3600
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   5741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin AniGIFCtrl.AniGIF AniGIF1 
      Height          =   2520
      Left            =   10080
      TabIndex        =   3
      Top             =   240
      Width           =   4080
      BackColor       =   16761024
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   -1  'True
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "Form1.frx":0000
      ExtendWidth     =   7197
      ExtendHeight    =   4445
      Loop            =   0
      AutoRewind      =   0   'False
      Synchronized    =   -1  'True
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   3255
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
      _Version        =   393216
      AutoPlay        =   -1  'True
      BackColor       =   -2147483638
      FullWidth       =   289
      FullHeight      =   217
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
      ExtentX         =   8705
      ExtentY         =   5530
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   5880
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   8880
      Top             =   3600
      Width           =   4815
   End
End
Attribute VB_Name = "PlayGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pic As Integer

Private Sub Command1_Click()
If Command1.Caption = "Play" Then
    Animation1.Open App.Path & "\Nyan.avi"
    Animation1.Play
    Command1.Caption = "Stop"
Else
    Animation1.Stop
    Command1.Caption = "Play"
End If
End Sub

Private Sub Form_Load()
Dim loc As String

WebBrowser1.Navigate (App.Path & "\" & "b7a.gif")
'AniGIF1.GIF = (App.Path & "\" & "b7a.gif")
loc = App.Path & "\b7a.gif"
WebBrowser2.Navigate "about:" & "<html>" & "<body leftMargin=0 topMargin=0 marginheight=0 marginwidth=0 scroll=no>" _
& "<img src=""" & loc & """></img></body></html>"
End Sub

Private Sub Timer1_Timer()
pic = pic + 1
Label1.Caption = pic

Select Case pic
    Case 1
        Image1.Picture = LoadPicture(App.Path & "\file\frame006.gif")
    Case 2
        Image1.Picture = LoadPicture(App.Path & "\file\frame007.gif")
    Case 3
        Image1.Picture = LoadPicture(App.Path & "\file\frame008.gif")
    Case 4
        Image1.Picture = LoadPicture(App.Path & "\file\frame009.gif")
    Case 5
        Image1.Picture = LoadPicture(App.Path & "\file\frame010.gif")
    Case 6
        Image1.Picture = LoadPicture(App.Path & "\file\frame011.gif")
        pic = 1
End Select
End Sub
