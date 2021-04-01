VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsFrame"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarxWidth As Long
Private mvaryHeight As Long
Private mvarTransparentColor As Long 'Transparent color
Private mvarxLeft As Long 'Left Position within graphic device
Private mvaryTop As Long 'Top Position within graphic device
Private mSignature As String
Private mVersion As String
Private mBackColor As Long
Private mvarComments As String
Private mRepeattimes As Long
Private mvarspeed As Integer
Private mvarFrames As clsFrames
Private mvarGlobalColorMap As clsColorMap
Private mvarpixel As Byte
Private mvarSortFlag As Boolean
Private mvarCR As Byte
Private mvarM As Byte
Private mvarAspectRatio As Byte
Private BkgndDC As Long
Private BkgndBM As Long

Public Event GIFError(ByVal Description As String)

'----------- Local Constants and Enumerations ---------------
'introducer
Private Const EXTENSION As Byte = &H21
'labels
Private Const APPLICATION As Byte = &HFF
Private Const COMMENT As Byte = &HFE
Private Const CONTROL As Byte = &HF9
Private Const PLAINTEXT As Byte = &H1
Private Const TRAILER As Byte = &H3B
Private Const IMAGE_SEPARATOR As Byte = &H44
Private Const IMAGE_DESCRIPTOR As Byte = &H2C
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long _
)

'------ GDI API Routines and Structures -----------------
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long _
) As Long

Private Declare Function SelectObject Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal hObject As Long _
) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" ( _
    ByVal hObject As Long, _
    ByVal nCount As Long, _
    lpObject As Any _
) As Long

Private Declare Function FillRect Lib "user32" ( _
    ByVal hdc As Long, _
    lpRect As RECT, _
    ByVal hBrush As Long _
) As Long

Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal dwRop As Long _
) As Long

Private Declare Function TransparentBlt Lib "Msimg32" ( _
    ByVal hDestDC As Long, _
    ByVal xDest As Long, _
    ByVal yDest As Long, _
    ByVal nWidthDest As Long, _
    ByVal nHeightDest As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal nWidthSrc As Long, _
    ByVal nHeightSrc As Long, _
    ByVal crTransparent As Long _
) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" ( _
    ByVal crColor As Long _
) As Long

Private Declare Function DeleteDC Lib "gdi32" ( _
    ByVal hdc As Long _
) As Long

Private Declare Function DeleteObject Lib "gdi32" ( _
    ByVal hObject As Long _
) As Long

Private Declare Function SetRect Lib "user32" ( _
    lpRect As RECT, _
    ByVal xLeft As Long, _
    ByVal yTop As Long, _
    ByVal xRight As Long, _
    ByVal yBottom As Long _
) As Long
Private Const HGDI_ERROR = &HFFFF
Const RC_PALETTE As Long = &H100
Const SIZEPALETTE As Long = 104
Const RASTERCAPS As Long = 38
Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPAL As Long
    Reserved As Long
End Type
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Private Sub AddFrame(Buffer() As Byte, ptr As Long)
Dim startptr As Long
Dim stopptr As Long
Dim myframe As clsFrame
Dim nBytes As Byte
Dim fnum As Integer
Dim picbuf As String
Dim localDC As Long
Static fileHeader As String



If mvarFrames.Count = 0 Then _
    fileHeader = BytesToString(Buffer(0), ptr + 1)

    'at this point, buffer(ptr+2) = CONTROL
    startptr = ptr + 1
    Set myframe = mvarFrames.Add
    With myframe
            
            'Graphic Control Exension
            ' may precede each image
            ' 1 - Extension Introducer = &h21
            ' 2- Graphic control label = &hF9
            ' 3 - Block Size
            ' 4(0) - Transparent Color Flag
            ' 4(1) - User Input Flag: 1=user input is expected before continuing
            ' 4(2-4) - Disposal Method
            ' 5/6 - Delay time: if >0 then specifies number of 1/100 seconds to wait before continuing
            '                    after the graphic has been rendered
            ' 7 - Transparent Color Index: 1=Transparent Index is given
            ' 8 - Block Terminator=0: marks end of graphic control extension

            '  Disposal methods
            '   0 -   No disposal specified. The decoder is not required to take any action.
            '   1 -   Do not dispose. The graphic is to be left in place.
            '   2 -   Restore to background color. The area used by the graphic must be restored
            '          to the background color.
            '   3 -   Restore to previous. The decoder is required to restore the area overwritten
            '          by the graphic with what was there prior to rendering the graphic.
            '   4-7 -    To be defined.
                nBytes = Buffer(ptr + 3)
                .TCI = Buffer(ptr + 7) 'transparency color index
                .DelayTime = BytesToInteger(Buffer(ptr + 5)) * mvarspeed
                .DisposalMethod = SplitBits(Buffer(ptr + 4), 2, 3)
                .TCFlag = ((Buffer(ptr + 4) And 1) = 1)
                .UserInputFlag = ((Buffer(ptr + 4) And 2) = 2)
                ptr = ptr + nBytes + 4
        
        Do
            Select Case Buffer(ptr + 1)
        
        'sometimes there is a comment between control block and image data
            Case EXTENSION
                Select Case Buffer(ptr + 2)
                Case COMMENT
                    GetComment Buffer(), ptr
                Case APPLICATION
                    GetApplication Buffer(), ptr
                Case CONTROL
                    Exit Do
                Case Else
                    Stop
                End Select
            Case IMAGE_DESCRIPTOR
            ' IMAGE DESCRIPTOR
            ' 1 - separator = &h2C
            ' 2/3 - Left LSB/MSB
            ' 4/5 - Top LSB/MSB
            ' 6/7 - Width LSB/MSB
            ' 8/9 - Height LSB/MSB
            ' 10 (0-2)Pixel - # bits per pixel for this image
            ' 10 (5) Sort Flag
            ' 10 (6)I - 1=Interlaced, 0=sequential
            ' 10 (7)M - 0=Use global color map, ignore 'pixel': 1=Local color map follows, use 'pixel'
                .xLeft = BytesToInteger(Buffer(ptr + 2))
                .yTop = BytesToInteger(Buffer(ptr + 4))
                .xWidth = BytesToInteger(Buffer(ptr + 6))
                .yHeight = BytesToInteger(Buffer(ptr + 8))
                .LCTSize = SplitBits(Buffer(ptr + 10), 0, 3)
                .SortFlag = ((Buffer(ptr + 10) And 32) = 32)
                .InterlaceFlag = ((Buffer(ptr + 10) And 64) = 64)
                .LCTFlag = ((Buffer(ptr + 10) And 128) = 128)
                
                'LOCAL COLOR TABLE
                If .LCTFlag = True Then
                    ptr = GetColorTable(Buffer, ptr + 11, .LCTSize, .ColorMap)
                Else
                    ptr = ptr + 10
                End If
'***** Problems
                'set transparent color
                If .TCFlag Then 'is there a transparency color?
                    If .LCTFlag Then 'local color table?
                        .TransparentColor = .ColorMap(.TCI).Color
                    Else 'or global color table
                        .TransparentColor = mvarGlobalColorMap(.TCI).Color
                    End If
                '.TransparentColor = mvarGlobalColorMap(1).Color
                'Debug.Print "Transparent color "; Hex(.TransparentColor); .TCI,
                'Debug.Print "BackColor "; Hex(mBackColor); BackgroundIndex
'NOTE: it seems that backcolor is sometimes used as transparent color??????
                    
                    'TransparentColor = mvarGlobalColorMap(1).Color 'mBackColor
                End If

                'IMAGE DATA
                '1 - LZW code Size
                '2 - Block Size
                '3 - n data
                'blocksize
                'data
                '...
                'Block Terminator = 0
        
                'advance pointer to end of image data
                ptr = AdvancePointer(Buffer, ptr + 2)
                stopptr = ptr
        '*** End of image data

                'output frame to temporary file
                fnum = FreeFile
                Open App.Path & "\temp.gif" For Binary As fnum
                picbuf = fileHeader & BytesToString(Buffer(startptr), stopptr - startptr + 1) & TRAILER
                Put #fnum, 1, picbuf
                Close fnum
                'load picture and delete temp file
                Set .Picture = LoadPicture(App.Path & "\temp.gif")
'                Set .Image = .Picture 'LoadPicture(App.Path & "\temp.gif")
                Kill (App.Path & "\temp.gif")
                'GetBMinfo .Picture
                
                'put picture into local DC
                localDC = CreateCompatibleDC(GetDC(0))
                Dim hOldBM As Long
                SelectObject localDC, .Picture.hPAL
                hOldBM = SelectObject(localDC, .Picture.Handle)
                
                'this is the problem, getting the picture into the DC
                Set .Image = hDCToPicture(localDC, 0, 0, mvarxWidth, mvaryHeight)
                
                Call SetDC(myframe, localDC)
                                
                'copy DC to picture
                Set .Picture = hDCToPicture(.hdc, 0, 0, mvarxWidth, mvaryHeight)
            Case Else
                Stop
            End Select
        Call SelectObject(localDC, hOldBM)
        DeleteDC localDC
    Loop While Not Buffer(ptr + 1) = TRAILER
End With
Set myframe = Nothing

End Sub

Public Sub Clear(hdc As Long)

If BkgndDC = 0 Then Exit Sub

BitBlt hdc, mvarxLeft, mvaryTop, mvarxWidth, mvaryHeight, BkgndDC, 0, 0, vbSrcCopy

End Sub


Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPAL As Long) As Picture
    Dim R As Long, pic As PicBmp, IPic As IPicture, IID_IDispatch As GUID

    'Fill GUID info
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With pic
        .Size = Len(pic) ' Length of structure
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPAL = hPAL ' Handle to palette (may be null)
    End With

    'Create the picture
    R = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)

    'Return the new picture
    Set CreateBitmapPicture = IPic
End Function
Sub GetApplication(Buffer() As Byte, ptr As Long)
Dim nBytes As Byte
        nBytes = Buffer(ptr + 3)

        'see if this is the NETSCAPE Loop extension
        If nBytes = &HB Then
            If BytesToString(Buffer(ptr + 4), 11) = "NETSCAPE2.0" Then
            'NETSCAPE LOOP APPLICATION EXTENSION
            ' 1 - Extension Introducer = &h21
            ' 2 - Application Label = &HFF
            ' 3 - Length of block = &H0B
            ' 4-14 - "NETSCAPE2.0"
            ' 15 - length of data subblock &H3
            ' 16 - &H1
            ' 17/18 - Number of iterations (LSB/MSB)
            ' 19 - Terminator &H00
                'get Number of iterations
                mRepeattimes = Buffer(ptr + 17) + Buffer(ptr + 18) * &H100
                ptr = ptr + 19
            Else
            'APPLICATION EXTENSION
            ' 1 - Extension Introducer = &h21
            ' 2 - Application Label = &HFF
            ' Bytes in data block
            'data
            'Block Terminator = 0
            'skip this extension
                ptr = ptr + nBytes + 4
                While Buffer(ptr) > 0
                    ptr = ptr + Buffer(ptr) + 1
                    'Debug.Print "UNKNOWN APPLICATION EXTENSION"
                Wend
            End If
        End If

End Sub

'For debug and development purposes
Public Sub GetBMinfo(pic As StdPicture)
Dim PicBits() As Byte, PicInfo As BITMAP
Dim Cnt As Long, BytesPerLine As Long
Dim i As Long, j As Long
   GetObject pic, Len(PicInfo), PicInfo
    'reallocate storage space
    
    ReDim PicBits(1 To PicInfo.bmWidthBytes, 1 To PicInfo.bmHeight)
    
    'Copy the bitmapbits to the array
    Call GetBitmapBits(pic, PicInfo.bmWidthBytes * PicInfo.bmHeight, PicBits(1, 1))
    
    For j = 1 To PicInfo.bmHeight
      For i = 1 To PicInfo.bmWidthBytes
        Debug.Print Hex(PicBits(i, j)),
     Next i
     Debug.Print
    Next j

    Stop

End Sub

Private Sub GetComment(Buffer() As Byte, ptr As Long)
    Dim nBytes As Byte
    'COMMENT EXTENSION
    ' 1 - Extension Introducer = &h21
    ' 2 - Comment Label = &HFE
    ' 3 - byte count (n)
    'n bytes of data
    'byte count
    'n bytes of data - 1 to 255 characters of comment
    '.....
    '0 byte count - terminator
        ptr = ptr + 3
 '       nbytes = Buffer(ptr - 1)
        'Debug.Print BytesToString(Buffer(ptr + 4), CLng(nbytes))
        
        While Buffer(ptr) > 0
            mvarComments = mvarComments & BytesToString(Buffer(ptr + 1), _
                CLng(Buffer(ptr))) & vbCrLf
            ptr = ptr + Buffer(ptr) + 1
        Wend

'ptr should now be pointing to zero
'ptr+1 is at 2C

End Sub


Sub GetPlainText(Buffer() As Byte, ptr As Long)
    'PLAIN TEXT EXTENSION
    ' 1 - Extension Introducer = &h21
    ' 2 - Plain Text Label = 1
    ' 3 - Number of bytes in data section
    ' .....
    '  data
    ' ....
    ' Block Terminator = 0
        Stop

End Sub


Function hDCToPicture(ByVal hDCSrc As Long, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long, hBmp As Long, hBmpPrev As Long, R As Long
    Dim hPAL As Long, hPalPrev As Long, RasterCapsScrn As Long, HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long, LogPal As LOGPALETTE

    'Create a compatible device context
    hDCMemory = CreateCompatibleDC(hDCSrc)
    'Create a compatible bitmap
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    'Select the compatible bitmap into our compatible device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    'Raster capabilities?
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
    'Does our picture use a palette?
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette
    'What's the size of that palette?
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Set the palette version
        LogPal.palVersion = &H300
        'Number of palette entries
        LogPal.palNumEntries = 256
        'Retrieve the system palette entries
        R = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        'Create the palette
        hPAL = CreatePalette(LogPal)
        'Select the palette
        hPalPrev = SelectPalette(hDCMemory, hPAL, 0)
        'Realize the palette
        R = RealizePalette(hDCMemory)
    End If

    'Copy the source image to our compatible device context
    R = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    'Restore the old bitmap
    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        'Select the palette
        hPAL = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    'Delete our memory DC
    R = DeleteDC(hDCMemory)

    Set hDCToPicture = CreateBitmapPicture(hBmp, hPAL)
End Function

Public Property Let AspectRatio(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AspectRatio = 5
    mvarAspectRatio = vData
End Property


Public Property Get AspectRatio() As Byte
    AspectRatio = mvarAspectRatio
End Property



Public Sub CopyFrame(Index As Integer, hdc As Long, xLeft As Long, yTop As Long)
Static initialized As Boolean
If Not initialized Then
    mvarxLeft = xLeft
    mvaryTop = yTop
    BkgndDC = CreateCompatibleDC(hdc)
    BkgndBM = CreateCompatibleBitmap(hdc, mvarxWidth, mvaryHeight)
    SelectObject BkgndDC, BkgndBM
    BitBlt BkgndDC, 0, 0, mvarxWidth, mvaryHeight, hdc, mvarxLeft, mvaryTop, vbSrcCopy
    initialized = True
End If

'copy frame to hdc
With mvarFrames(Index)
    If .TCFlag Then
        If mvarM = True Then
            Call TransparentBlt(hdc, xLeft, yTop, mvarxWidth, mvaryHeight, _
                .hdc, 0, 0, mvarxWidth, mvaryHeight, .TransparentColor)
        Else
            Call TransparentBlt(hdc, xLeft, yTop, mvarxWidth, mvaryHeight, _
                .hdc, 0, 0, mvarxWidth, mvaryHeight, .TransparentColor)
        End If
    Else
        BitBlt hdc, xLeft, yTop, mvarxWidth, mvaryHeight, .hdc, 0, 0, vbSrcCopy
    End If
End With
End Sub

Public Property Get Image(Index As Integer) As IPicture
    Set Image = mvarFrames(Index).Image
End Property

Private Sub ImageToDC(objPic As IPicture, hdc As Long)
'puts picture into DC

End Sub

Public Property Let M(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.M = 5
    mvarM = vData
End Property


Public Property Get M() As Byte
    M = mvarM
End Property



Public Property Let CR(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CR = 5
    mvarCR = vData
End Property


Public Property Get CR() As Byte
    CR = mvarCR
End Property



Public Sub SetDC(myframe As clsFrame, hdc As Long)
'hdc contains new image to overlay onto frame
    Dim i As Integer
    Dim Index As Integer
    Dim hbrBkgnd As Long
    Dim myRect As RECT
    Dim hBitmap As Long
    Dim hbrtemp As Long

 With myframe
    Index = mvarFrames.Count
   'create a memory DC
    .hdc = CreateCompatibleDC(GetDC(0))
   
   'create a bitmap the size of the logical screen
   hBitmap = CreateCompatibleBitmap(GetDC(0), mvarxWidth, mvaryHeight)
    
    'put bitmap into DC
    Call SelectObject(.hdc, hBitmap)

'for 1st frame
'   set background color
'   copy image in
'for next frames,
'   copy previous image in
'   follow disposal method
'   overlay new image
        'Set .Image = hDCToPicture(hdc, 0, 0, mvarxWidth, mvaryHeight)

    hbrBkgnd = CreateSolidBrush(.TransparentColor)
    If Index = 1 Then
    'set background color
        Call SetRect(myRect, 0, 0, mvarxWidth, mvaryHeight)
        Call FillRect(.hdc, myRect, hbrBkgnd)
     'copy image in
        'Call BitBlt(.hdc, .xLeft, .yTop, .xWidth, .yHeight, hdc, 0, 0, vbSrcCopy)

        Call TransparentBlt(.hdc, .xLeft, .yTop, .xWidth, .yHeight, _
            hdc, 0, 0, .xWidth, .yHeight, .TransparentColor)
    Else
        With mvarFrames(Index - 1)
            'copy previous image in (fullsize)
            Call SetRect(myRect, 0, 0, mvarxWidth, mvaryHeight)
            Call FillRect(myframe.hdc, myRect, hbrBkgnd)
            Call BitBlt(myframe.hdc, 0, 0, mvarxWidth, mvaryHeight, .hdc, 0, 0, vbSrcCopy)
        '    Call TransparentBlt(myFrame.hdc, 0, 0, mvarxWidth, mvaryHeight, _
                .hdc, 0, 0, mvarxWidth, mvaryHeight, .Transparentcolor)
            
            Select Case mvarFrames(Index - 1).DisposalMethod
            Case Is = dmNONE, dmDONOT 'don't dispose of previous image
                
            Case Is = dmRESTORE_BKGND 'dispose of previous image
                Call SetRect(myRect, .xLeft, .yTop, .xWidth + .xLeft, .yHeight + .yTop)
                Call FillRect(myframe.hdc, myRect, hbrBkgnd)
            Case Is = dmRESTORE_PREVIOUS
                If Index > 2 Then
                    Call BitBlt(myframe.hdc, 0, 0, mvarxWidth, mvaryHeight, mvarFrames(Index - 2).hdc, 0, 0, vbSrcCopy)
                Else
                    Stop
                End If
            End Select
        End With
    'overlay new image
        If .TCFlag Then 'if transparent
            Call TransparentBlt(.hdc, .xLeft, .yTop, .xWidth, .yHeight, _
            hdc, 0, 0, .xWidth, .yHeight, .TransparentColor)
        Else
            Call BitBlt(.hdc, .xLeft, .yTop, .xWidth, .yHeight, hdc, 0, 0, vbSrcCopy)
        End If
    End If

 End With
 
ExitSub:
 DeleteObject hBitmap
 DeleteObject hbrBkgnd
 
End Sub

Public Property Let SortFlag(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SortFlag = 5
    mvarSortFlag = vData
End Property


Public Property Get SortFlag() As Boolean
    SortFlag = mvarSortFlag
End Property



Public Property Let pixel(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.pixel = 5
    mvarpixel = vData
End Property


Public Property Get pixel() As Byte
    pixel = mvarpixel
End Property

Public Property Get GlobalColorMap() As clsColorMap
    If mvarGlobalColorMap Is Nothing Then
        Set mvarGlobalColorMap = New clsColorMap
    End If


    Set GlobalColorMap = mvarGlobalColorMap
End Property


Public Property Set GlobalColorMap(vData As clsColorMap)
    Set mvarGlobalColorMap = vData
End Property


Public Property Get Frames() As clsFrames
'    If mvarFrames Is Nothing Then
'        Set mvarFrames = New clsFrames
'    End If
'
'
    Set Frames = mvarFrames
End Property


Public Property Set Frames(vData As clsFrames)
    Set mvarFrames = vData
End Property


Private Sub Class_Initialize()
    Set mvarGlobalColorMap = New clsColorMap
    'Set mvarFrames = New clsFrames
    mvarspeed = 10
End Sub

Private Sub Class_Terminate()
  Set mvarGlobalColorMap = Nothing
  Set mvarFrames = Nothing
End Sub






Public Property Let Speed(ByVal vData As Integer)
    mvarspeed = vData
End Property


Public Property Get Speed() As Integer
    Speed = mvarspeed
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
    mvarxLeft = vData
End Property


Public Property Get xLeft() As Long
    xLeft = mvarxLeft
End Property



Public Function BackColor() As Long
'if there is a global color map then use backcolor

    BackColor = mBackColor

End Function





Public Property Get yHeight() As Long
    yHeight = mvaryHeight
End Property





Public Property Get xWidth() As Long
    xWidth = mvarxWidth
End Property

Public Function LoadGIF(Resource As Variant) As Boolean
Dim ptr As Long
    Dim Buffer() As Byte
    Dim nBytes As Byte
    Dim startptr As Long
    Dim stopptr As Long
    Dim hbrBkgnd As Long
    'Dim myRect As RECT
    'Dim mFrameData As FrameData
    Dim BackgroundIndex As Integer
    
    Dim fnum As Integer
    Dim i As Long
    Dim picbuf As String
    Dim mColorMapSize As Integer
    Dim Index As Integer
    Dim hREGN As Long
    Dim LogPal As LOGPALETTE
    Dim hPAL As Long
    
    LoadGIF = False

    Set mvarFrames = Nothing
    Set mvarFrames = New clsFrames
    mvarComments = ""
    'need to clear all previous classes
'    For i = mvarFrames.Count To 1 Step -1
'        Set mvarFrames(i).Picture = LoadPicture
 '       Set mvarFrames(i).Image = LoadPicture
  '      Set mvarFrames(i).Picture = Nothing
   '     Set mvarFrames(i).Image = Nothing
'        If mvarFrames(i).hdc Then DeleteDC mvarFrames(i).hdc
'        mvarFrames.Remove i
'    Next i
    
    For i = mvarGlobalColorMap.Count To 1 Step -1
        mvarGlobalColorMap.Remove i
    Next i

    If VarType(Resource) = vbString Then 'load gif from file
        'see if file exists
        If Dir$(Resource) = "" Or Resource = "" Then
            RaiseEvent GIFError("File Does Not Exist.")
            Exit Function
        End If
        fnum = FreeFile
        Open Resource For Binary Access Read As fnum
            ReDim Buffer(0 To LOF(fnum) - 1)
            Get #fnum, , Buffer 'Get contents of GIF File into buffer
        Close fnum
    ElseIf VarType(Resource) = vbByte + vbArray Then 'data already in array
        Buffer = Resource
    Else 'try loading from resource file
        Buffer = LoadResData(Resource, "CUSTOM")
    End If


' GIF Signature
'The HEADER block is a small 6-byte (6-character) block. It is the first block in
'every file and contains the GIF version of the file (i.e either GIF87a or GIF89a). '
'GIF image decoders use this information to determine the version of the file.

'look for tag declaring GIF file
    mSignature = BytesToString(Buffer(0), 3)
    If mSignature <> "GIF" Then
'       RaiseEvent GIFError("Not a gif file: File format is not supported.")
       Exit Function
    End If

    mVersion = BytesToString(Buffer(3), 3)
    If mVersion <> "89a" Then
        'RaiseEvent GIFError("GIF version " & mVersion & " is not supported.")
        'Exit Function
    End If
' BLOCK #2 LOGICAL SCREEN DESCRIPTOR

'The LOGICAL SCREEN DESCRIPTOR
'defines an area of pixels which you can think of as a screen.
'The dimensions of this area define the size of your GIF on screen.
'The Logical Screen Block also chooses one of the colors in the Global
'Color Table to be the Background color of the screen.
'The Logical Screen Header can also contain the aspect ratio of the image

'Screen Descriptor - 7 bytes
' 1/2 - Width LSB/MSB
' 3/4 - Height LSB/MSB
' 5 - Pixel (0-2) - pixel+1 = # bits/pixel in image.
'      The value of 'pixel' also defines the maximum number of colors within an image
' 5 - SortFlag (3) - if true then color table is sorted
' 5 - cr (4-6) - cr+1 = # bits of color resolution
' 5 - M (7) - M = 1, Global color map follows Descriptor
' 6 - Background = Color index of screen background color
' 7 - Pixel Aspect Ratio - pixel width/pixel height: Aspect ratio =
'      (Pixel Aspect Ratio + 15) / 64

    mvarxWidth = BytesToInteger(Buffer(6))
    mvaryHeight = BytesToInteger(Buffer(8))

    mvarpixel = SplitBits(Buffer(10), 0, 3)
    mvarSortFlag = ((Buffer(10) And &H8) = &H8)
    mvarCR = SplitBits(Buffer(10), 4, 3)
    mvarM = ((Buffer(10) And &H80) = &H80)
    BackgroundIndex = Buffer(11)
    mvarAspectRatio = Buffer(12)

Debug.Print "Bit color resolution: "; mvarCR
Debug.Print "Bits per Pixel: "; mvarpixel
'GLOBAL COLOR MAP
' Block #3 Global Color Map(optional)M=1
'size of color map is 3*2^bitsperPixel colors

    
    If mvarM = True Then
        ptr = GetColorTable(Buffer(), ptr + 13, mvarpixel, mvarGlobalColorMap)
        mColorMapSize = 2 ^ (mvarpixel + 1)
        If BackgroundIndex < mColorMapSize Then
            mBackColor = mvarGlobalColorMap(BackgroundIndex).Color
        Else
            Stop
        End If
    Else 'no global color map
        ptr = ptr + 12
    End If
        'Set objpic.Picture = LoadPicture

'GIF EXTENSION BLOCKS
' 1 - Introducer(marker) = &H21
' 2 - Function Code
' 3 - byte count (n)
'n bytes of data
'byte count
'n bytes of data
'.....
'0 byte count - terminator

'************************************************************
'get any global extensions that exist
mvarComments = ""
Do

    Select Case Buffer(ptr + 2) ' label
    
    Case APPLICATION
        GetApplication Buffer, ptr
    
    Case COMMENT
        GetComment Buffer, ptr
        
    Case Is = PLAINTEXT
        GetPlainText Buffer, ptr
    
    Case CONTROL
        AddFrame Buffer, ptr
    
    Case TRAILER
    
    Case Else 'no control block
        Stop
        Exit Do
    
    End Select



Loop While Not Buffer(ptr + 1) = TRAILER
'**********************************************************************


    

ExitFunction:

    LoadGIF = True

End Function

Function ImageCount(Buffer() As Byte, ByVal ptr As Long) As Integer
Dim LCTSize As Byte
Dim LCTFlag As Boolean
Dim nBytes As Byte
Dim L As Integer
Dim icount As Integer
    icount = 0
    Do
        If Buffer(ptr + 2) = CONTROL Then
            'Graphic Control Exension
            nBytes = Buffer(ptr + 3)
            ptr = ptr + nBytes + 4
        Else
            'No Control Block
        End If
                Stop
        'check for local comments
        If Buffer(ptr + 2) = COMMENT Then
            nBytes = Buffer(ptr + 3)
            ptr = ptr + nBytes + 4
            While Buffer(ptr) > 0
                ptr = ptr + Buffer(ptr) + 1
            Wend
        
        End If
        
        'after control should be an Image block
        ' IMAGE BLOCK
        If Buffer(ptr + 1) = IMAGE_DESCRIPTOR Then
            LCTSize = SplitBits(Buffer(ptr + 10), 0, 3)
            LCTFlag = ((Buffer(ptr + 10) And 128) = 128)
            
        'LOCAL COLOR TABLE
            If LCTFlag = True Then
                L = 2 ^ (LCTSize + 1)
                ptr = ptr + 3 * L - 1 + 11
            Else
                ptr = ptr + 10
            End If
        
    'IMAGE DATA
        'advance pointer to end of image data
        ptr = AdvancePointer(Buffer, ptr + 2)
        icount = icount + 1
'    Else
 '       If Buffer(ptr + 1) <> TRAILER Then Stop
    End If
    Stop
    Loop While Not Buffer(ptr + 1) = TRAILER
    
    ImageCount = icount

End Function
Private Function AdvancePointer(Buffer() As Byte, ptr As Long) As Long
'returns a pointer to the block terminator
Dim nBytes As Byte
Dim p As Long
    p = ptr
    '1st byte is number of bytes in data block
    nBytes = Buffer(p)
    While nBytes > 0
        'point to next block
        p = p + nBytes + 1
        nBytes = Buffer(p)
    Wend
    AdvancePointer = p
End Function

Private Function BytesToInteger(Buffer As Byte) As Integer
    CopyMemory BytesToInteger, Buffer, 2
End Function

Private Function BytesToString(Buffer As Byte, Length As Long) As String

BytesToString = String(Length, 0)
CopyMemory ByVal BytesToString, Buffer, Length

End Function

'Private Sub EraseBackGround(myRect As RECT, ByVal fillColor As Long)
'Dim hbrBkgnd As Long
'        If fillColor And &H80000000 Then
'            hbrBkgnd = GetSysColorBrush(fillColor And &HFFFFFF)
'        Else
'            hbrBkgnd = CreateSolidBrush(fillColor)
'        End If
'        Call FillRect(myFrame.hDc, myRect, hbrBkgnd)
'        DeleteObject hbrBkgnd
'End Sub

Private Function GetColorTable(Buffer() As Byte, ByVal ptr As Long, ByVal lenTable As Byte, Table As clsColorMap) As Long
    Dim L As Long
    Dim i As Long, j As Long
    L = 2 ^ (lenTable + 1)
    For i = 0 To L - 1
        j = ptr + i * 3
        Table.Add Buffer(j), Buffer(j + 1), Buffer(j + 2)
    Next i
    

    GetColorTable = ptr + 3 * L - 1
End Function
Private Function SplitBits(ByVal X As Byte, Startbit As Byte, nBits As Byte) As Byte

'shift x right startbits times
    X = X \ (2 ^ Startbit)
'mask off top end
    SplitBits = X And (2 ^ nBits - 1)

End Function


