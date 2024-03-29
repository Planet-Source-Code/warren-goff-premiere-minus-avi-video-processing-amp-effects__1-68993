VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ==================================================================================
' cDIBSection256.cls
' Copyright © 1999 Steve McMahon
' Visit vbAccelerator at http://vbaccelerator.com
'
' Creates and manages a 256 colour GDI DibSection.  This is DIB
' in which the bitmap bits are stored in windows memory so can
' be modified.  Also, there are only 256 colours (1 byte/pixel)
' and the colour palette can be modified using GetDIBColorTable
' and SetDIBColorTable.  This means fades etc can be achieved
' by simply manipulating the DIB Color Table, rather than
' modifying the bitmap bits.  By doing this, a fade on a
' 512x512 fade can run much quicker than the equivalent for a
' True Colour DIB.
' The speed you run at depends on how your gfx driver implements
' DIB colour tables.  On a 8Mb ATI Xpert@Work, Win95, this code
' runs at > 300 fps for a 256x256 DIB!  However, on a 4Mb Matrox
' Millenium, NT it runs at less speed.
'
' Note: for best performance, when compiling an executable check
' all the boxes on the Properties-Compile tab Advanced Optimisations
' button, particularly Remove Array Bounds checks.
' ==================================================================================


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO256
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO256, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pRGBQuad As Any) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Private Declare Function GetDIBits256 Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

' Start of structure:
Private Const BITMAPTYPE As Integer = &H4D42
Private Type BITMAPFILEHEADER
   bfType As Integer '- type  ="BM" i.e &H4D42 - 2
   bfSize As Long ' - size in bytes of file - 6
   bfReserved1 As Integer ' - reserved, must be 0 - 8
   bfReserved2 As Integer ' - reserved, must be 0 - 10
   bfOffBits As Long ' offset from this structure to the bitmap bits - 14
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const CREATE_ALWAYS = 2
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_BEGIN = 0
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Long, lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO256, ByVal wUsage As Long) As Long



Private m_hDIb As Long
Private m_hBmpOld As Long
Private m_hDC As Long
Private m_lPtr As Long
Private m_tBI As BITMAPINFO256
' for speed - declare RGB array as global
Private tRGB(0 To 255) As RGBQUAD

Friend Sub SetPalette( _
      cP As cPalette _
   )
   cP.ExtractToRGBQuadArray tRGB
   
   Dim i As Long
   'For i = 0 To 255
   '   Debug.Print Hex(tRGB(i).rgbBlue) & Hex(tRGB(i).rgbGreen) & Hex(tRGB(i).rgbRed)
   'Next i
End Sub

Public Function CreateDIB( _
        ByVal lHDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean
Dim i As Long
   With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 8
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
   End With
   ' Create Gray scale palette as default:
   For i = 0 To 255
      With m_tBI.bmiColors(i)
         .rgbBlue = i
         .rgbGreen = i
         .rgbRed = i
      End With
   Next
   hDib = CreateDIBSection( _
            lHDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)
   CreateDIB = (hDib <> 0)
End Function
Public Function CreateFromPicture( _
        ByRef picThis As StdPicture _
    )
Dim lHDC As Long
Dim lhDCDesktop As Long
Dim lhBmpOld As Long
Dim tBMP As BITMAP
Dim lC As Long
    
   GetObjectAPI picThis.handle, Len(tBMP), tBMP
   If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
      lhDCDesktop = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
      If (lhDCDesktop <> 0) Then
         lHDC = CreateCompatibleDC(lhDCDesktop)
         DeleteDC lhDCDesktop
         If (lHDC <> 0) Then
            ' Select the bitmap into the compatible DC:
            lhBmpOld = SelectObject(lHDC, picThis.handle)
            ' Get the DIB Color Table (according to the docs, GetDIBits should do this, but it
            ' doesn't seen to):
            lC = GetDIBColorTable(lHDC, 0, 256, tRGB(0))
            ' if this assert fails, the picture you're creating from
            ' is not 256 colours:
            Debug.Assert (lC = 256)
            ' Move the bits across:
            GetDIBits256 lHDC, picThis.handle, 0, tBMP.bmHeight, ByVal m_lPtr, m_tBI, DIB_RGB_COLORS
            ' Set the colour table to correct values:
            If (lC > 0) Then
              SetDIBColorTable m_hDC, 0, 256, tRGB(0)
            End If
            ' clear up:
            SelectObject lHDC, lhBmpOld
            DeleteObject lHDC
         End If
      End If
   End If
End Function
Public Function Create( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
Dim lHDCDesk As Long
    ClearUp
    lHDCDesk = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    m_hDC = CreateCompatibleDC(lHDCDesk)
    DeleteDC lHDCDesk
    If (m_hDC <> 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            Create = True
        Else
            DeleteObject m_hDC
            m_hDC = 0
        End If
    End If
End Function
Public Property Get BytesPerScanLine() As Long
    ' Scans must align on dword boundaries:
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property

Public Sub LoadPictureBlt( _
        ByVal lHDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
Dim lC As Long
   
   SetDIBColorTable m_hDC, 0, 256, tRGB(0)
   If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
   If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
   BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lHDC, lSrcLeft, lSrcTop, eRop
    
End Sub


Public Sub PaintPicture( _
        ByVal lHDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    BitBlt lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Property Get hdc() As Long
    hdc = m_hDC
End Property
Public Property Get hDib() As Long
    hDib = m_hDIb
End Property
Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property
Public Sub RandomiseBits()
Dim bDib() As Byte
Dim x As Long, y As Long
Dim lC As Long
Dim tSA As SAFEARRAY2D
Dim xEnd As Long
    
   ' Get the bits in the from DIB section:
   With tSA
       .cbElements = 1
       .cDims = 2
       .Bounds(0).lLbound = 0
       .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
       .Bounds(1).lLbound = 0
       .Bounds(1).cElements = BytesPerScanLine()
       .pvData = m_lPtr
   End With
   CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

   ' random:
   Randomize Timer
    
    xEnd = Width - 1
   For y = 0 To m_tBI.bmiHeader.biHeight - 1
       For x = 0 To xEnd
           lC = Rnd * 255
           bDib(x, y) = lC
       Next
   Next
    
   ' Clear the temporary array descriptor
   CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
End Sub

Public Sub ClearUp()
    If (m_hDC <> 0) Then
        If (m_hDIb <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIb
        End If
        DeleteObject m_hDC
    End If
    m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0
End Sub

Public Sub Fade(ByVal lAmount As Long)
Dim tRGBOut(0 To 255) As RGBQUAD
Dim lR As Long, lG As Long, lB As Long
Dim i As Long
Dim lC As Long
   lC = GetDIBColorTable(m_hDC, 0, 256, tRGB(0))
   Debug.Assert (lC = 256)
   For i = 0 To lC - 1
      With tRGB(i)
         lB = lAmount * .rgbBlue \ 255
         lG = lAmount * .rgbGreen \ 255
         lR = lAmount * .rgbRed \ 255
      End With
      With tRGBOut(i)
         .rgbBlue = lB
         .rgbGreen = lG
         .rgbRed = lR
      End With
   Next
   lC = SetDIBColorTable(m_hDC, 0, 256, tRGBOut(0))
   Debug.Assert (lC = 256)
End Sub
Public Property Get Color(ByVal nIndex As Long) As Long
Dim lC As Long
Dim tRGBItem As RGBQUAD
   lC = GetDIBColorTable(m_hDC, nIndex, 1, tRGBItem)
   Debug.Assert (lC = 1)
   If (lC = 1) Then
      Color = tRGBItem.rgbRed Or (tRGBItem.rgbGreen * &H100) Or (tRGBItem.rgbBlue * &H10000)
   Else
      Color = -1  ' CLR_INVALID
   End If
End Property
Public Property Let Color(ByVal nIndex As Long, ByVal lColor As Long)
Dim lC As Long
Dim tRGBItem As RGBQUAD
Dim lB As Long, lG As Long, lR As Long
   lB = (lColor And &HFF0000) \ &H10000
   lG = (lColor And &HFF00&) \ &H100
   lR = (lColor And &HFF)
   tRGBItem.rgbBlue = lB
   tRGBItem.rgbGreen = lG
   tRGBItem.rgbRed = lR
   lC = SetDIBColorTable(m_hDC, nIndex, 1, tRGBItem)
End Property
Public Sub AdjustLightness(ByVal lAmount As Long)
Dim tRGBOut(0 To 256) As RGBQUAD
Dim lC As Long
Static i As Long
Static fAmount As Single
Static r As Long, g As Long, b As Long
Static h As Single, s As Single, l As Single

   fAmount = lAmount / 100#
   lC = GetDIBColorTable(m_hDC, 0, 256, tRGB(0))
   Debug.Assert (lC = 256)
   If (lC > 0) Then
      For i = 0 To lC
         RGBToHSL tRGB(i).rgbRed, tRGB(i).rgbGreen, tRGB(i).rgbBlue, h, s, l
         l = l * fAmount
         HLSToRGB h, s, l, r, g, b
         If r < 0 Then r = 0
         If r > 255 Then r = 255
         If g < 0 Then g = 0
         If g > 255 Then g = 255
         If b < 0 Then b = 0
         If b > 255 Then b = 255
         tRGBOut(i).rgbBlue = b
         tRGBOut(i).rgbGreen = g
         tRGBOut(i).rgbRed = r
      Next
      lC = SetDIBColorTable(m_hDC, 0, 256, tRGBOut(0))
      Debug.Assert (lC = 256)
   End If
End Sub
Public Sub GrayScale()
Dim tRGBOut(0 To 256) As RGBQUAD
Dim lC As Long
Static i As Long
Static lGS As Long

   lC = GetDIBColorTable(m_hDC, 0, 256, tRGB(0))
   Debug.Assert (lC = 256)
   If (lC > 0) Then
      For i = 0 To lC
         lGS = (222& * tRGB(i).rgbRed + 707& * tRGB(i).rgbGreen + 71& * tRGB(i).rgbBlue) / 1000&
         tRGBOut(i).rgbBlue = lGS
         tRGBOut(i).rgbGreen = lGS
         tRGBOut(i).rgbRed = lGS
      Next
      lC = SetDIBColorTable(m_hDC, 0, 256, tRGBOut(0))
      Debug.Assert (lC = 256)
   End If
End Sub
Public Sub Invert()
Dim tRGBOut(0 To 256) As RGBQUAD
Dim lC As Long, i As Long
   lC = GetDIBColorTable(m_hDC, 0, 256, tRGB(0))
   Debug.Assert (lC = 256)
   For i = 0 To lC
      tRGBOut(i).rgbBlue = (&HFF Xor tRGB(i).rgbBlue)
      tRGBOut(i).rgbGreen = (&HFF Xor tRGB(i).rgbGreen)
      tRGBOut(i).rgbRed = (&HFF Xor tRGB(i).rgbRed)
   Next
   lC = SetDIBColorTable(m_hDC, 0, 256, tRGBOut(0))
   Debug.Assert (lC = 256)
   
End Sub

Public Sub CopyPalette(ByRef cDib As cDIBSection256)
Dim lC As Long
   lC = GetDIBColorTable(cDib.hdc, 0, 256, tRGB(0))
   Debug.Assert (lC = 256)
   If (lC > 0) Then
      lC = SetDIBColorTable(m_hDC, 0, 256, tRGB(0))
      Debug.Assert (lC = 256)
   End If
End Sub
Private Sub RGBToHSL( _
      ByVal r As Long, ByVal g As Long, ByVal b As Long, _
      h As Single, s As Single, l As Single _
   )
Dim Max As Single
Dim Min As Single
Dim delta As Single
Dim rR As Single, rG As Single, rB As Single

   rR = r / 255: rG = g / 255: rB = b / 255

'{Given: rgb each in [0,1].
' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
        Max = Maximum(rR, rG, rB)
        Min = Minimum(rR, rG, rB)
        l = (Max + Min) / 2    '{This is the lightness}
        '{Next calculate saturation}
        If Max = Min Then
            'begin {Acrhomatic case}
            s = 0
            h = 0
           'end {Acrhomatic case}
        Else
           'begin {Chromatic case}
                '{First calculate the saturation.}
           If l <= 0.5 Then
               s = (Max - Min) / (Max + Min)
           Else
               s = (Max - Min) / (2 - Max - Min)
            End If
            '{Next calculate the hue.}
            delta = Max - Min
           If rR = Max Then
                h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
           ElseIf rG = Max Then
                h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
           ElseIf rB = Max Then
                h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
            End If
            'Debug.Print h
            'h = h * 60
           'If h < 0# Then
           '     h = h + 360            '{Make degrees be nonnegative}
           'End If
        'end {Chromatic Case}
      End If
'end {RGB_to_HLS}
End Sub

Private Sub HLSToRGB( _
      h As Single, s As Single, l As Single, _
      r As Long, g As Long, b As Long _
   )
Dim rR As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single

   If s = 0 Then
      ' Achromatic case:
      rR = l: rG = l: rB = l
   Else
      ' Chromatic case:
      ' delta = Max-Min
      If l <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = l * (1 - s)
      Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = l - s * (1 - l)
      End If
      ' Get the Max value:
      Max = 2 * l - Min
      
      ' Now depending on sector we can evaluate the h,l,s:
      If (h < 1) Then
         rR = Max
         If (h < 0) Then
            rG = Min
            rB = rG - h * (Max - Min)
         Else
            rB = Min
            rG = h * (Max - Min) + rB
         End If
      ElseIf (h < 3) Then
         rG = Max
         If (h < 2) Then
            rB = Min
            rR = rB - (h - 2) * (Max - Min)
         Else
            rR = Min
            rB = (h - 2) * (Max - Min) + rR
         End If
      Else
         rB = Max
         If (h < 4) Then
            rR = Min
            rG = rR - (h - 4) * (Max - Min)
         Else
            rG = Min
            rR = (h - 4) * (Max - Min) + rG
         End If
         
      End If
            
   End If
   r = rR * 255: g = rG * 255: b = rB * 255
End Sub
Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function
Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function

Public Function SavePicture(ByVal sFileName As String) As Boolean
Dim lC As Long, i As Long

   ' Fix up the palette to match the current DIB colour table
   lC = GetDIBColorTable(m_hDC, 0, 256, tRGB(0))
   Debug.Assert (lC = 256)
   For i = 0 To lC - 1
      LSet m_tBI.bmiColors(i) = tRGB(i)
   Next
   ' Save to BMP with 256 colour palette:
   SavePicture = SaveToBitmap(m_tBI, m_lPtr, sFileName)

End Function
Private Function SaveToBitmap(ByRef tBI As BITMAPINFO256, ByVal lPtrBits As Long, ByVal sFileName As String)
Dim tBH As BITMAPFILEHEADER
Dim tRGBQ As RGBQUAD
Dim hFile As Long
Dim lBytesWritten As Long
Dim lSize As Long
Dim lR As Long
Dim bErr As Boolean
Dim hMem As Long, lPtr As Long
Dim lErr As Long

   ' Prepare the BITMAPFILEHEADER
   With tBH
      .bfType = BITMAPTYPE
      .bfOffBits = 14 + Len(tBI)
      .bfSize = .bfOffBits + tBI.bmiHeader.biSizeImage
   End With
   hFile = CreateFile(sFileName, _
                 GENERIC_READ Or GENERIC_WRITE, _
                  ByVal 0&, _
                  ByVal 0&, _
                  CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, _
                  0)
   lErr = Err.LastDllError
   If (hFile = INVALID_HANDLE_VALUE) Then
      ' error
      Err.Raise 17, App.EXEName & ".cDIBSection256", ApiError(lErr)
   Else
      
      ' Writing the BITMAPFILEINFOHEADER is somewhat painful
      ' due to non-byte alignment of structure...
      hMem = GlobalAlloc(GPTR, 14)
      lPtr = GlobalLock(hMem)
      CopyMemory ByVal lPtr, tBH.bfType, 2
      CopyMemory ByVal lPtr + 2, tBH.bfSize, 4
      CopyMemory ByVal lPtr + 6, 0&, 4
      CopyMemory ByVal lPtr + 10, tBH.bfOffBits, 4
      lSize = 14
      lR = WriteFile(hFile, ByVal lPtr, lSize, lBytesWritten, ByVal 0&)
      GlobalUnlock hMem
      GlobalFree hMem
      
      ' Add the BITMAPINFOHEADER and colour palette:
      bErr = FileErrHandler(lR, lSize, lBytesWritten)
      If Not bErr Then
         lSize = Len(tBI)
         lR = WriteFile(hFile, tBI, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      If Not bErr Then
         ' Its easy to write the bitmap data, though...
         lSize = tBI.bmiHeader.biSizeImage
         lR = WriteFile(hFile, ByVal lPtrBits, lSize, lBytesWritten, ByVal 0&)
         bErr = FileErrHandler(lR, lSize, lBytesWritten)
      End If
      
      
      CloseHandle hFile
      'SavePicture = Not (bErr)
   End If

End Function
Private Function ApiError(ByVal e As Long) As String
    Dim s As String, c As Long
    s = String(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      0, e, 0&, s, Len(s), ByVal 0)
    If c Then ApiError = Left$(s, c)
End Function

Private Function FileErrHandler(ByVal lR As Long, ByVal lSize As Long, ByVal lBytes As Long) As Boolean
   If (lR = 0) Or Not (lSize = lBytes) Then
      'Err.Raise
      FileErrHandler = True
   End If
End Function




Private Sub Class_Terminate()
    ClearUp
End Sub


