VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAVICreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ERR_BASE As Long = vbObjectError + 1024 + 77561

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Type RECT
   Left As Long
   TOp As Long
   Right As Long
   Bottom As Long
End Type

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

Private Type TAVISTREAMINFO ' this is the ANSI version
    fccType As Long
    fccHandler As Long
    dwFlags As Long         '/* Contains AVITF_* flags */
    dwCaps As Long
    wPriority As Integer
    wLanguage As Integer
    dwScale As Long
    dwRate As Long ' /* dwRate / dwScale == samples/second */
    dwStart As Long
    dwLength As Long '; /* In units above... */
    dwInitialFrames As Long
    dwSuggestedBufferSize As Long
    dwQuality As Long
    dwSampleSize As Long
    rcFrame As RECT
    dwEditCount As Long
    dwFormatChangeCount As Long
    szName(0 To 63) As Byte
End Type

Private Type TAVICOMPRESSOPTIONS
    fccType As Long
    fccHandler As Long
    dwKeyFrameEvery As Long
    dwQuality As Long
    dwBytesPerSecond As Long
    dwFlags As Long
    lpFormat As Long
    cbFormat As Long
    lpParms As Long
    cbParms As Long
    dwInterleaveEvery As Long
End Type

Private Declare Sub AVIFileInit Lib "avifil32.dll" ()
Private Declare Sub AVIFileExit Lib "avifil32.dll" ()

Private Declare Function AVIFileOpen Lib "avifil32.dll" Alias "AVIFileOpenA" ( _
   ppfile As Any, _
   ByVal szFile As String, _
   ByVal uMode As Long, _
   lpHandler As Any) As Long

Private Declare Function AVIFileCreateStream Lib "avifil32.dll" ( _
      pfile As Any, _
      ppavi As Any, _
      psi As TAVISTREAMINFO _
   ) As Long

Private Declare Function AVIMakeCompressedStream Lib "avifil32.dll" ( _
      ppsCompressed As Any, _
      ppsSource As Any, _
      lpOptions As TAVICOMPRESSOPTIONS, _
   pclsidHandler As Any) As Long

Private Declare Function AVIStreamSetFormat Lib "avifil32.dll" ( _
      pavi As Any, _
      ByVal lPos As Long, _
      lpFormat As Any, _
      ByVal cbFormat As Long _
   ) As Long

Private Declare Function AVIStreamWrite Lib "avifil32.dll" ( _
      pavi As Any, _
      ByVal lStart As Long, _
      ByVal lSamples As Long, _
      lpBuffer As Any, _
      ByVal cbBuffer As Long, _
      ByVal dwFlags As Long, _
      plSampWritten As Long, _
      plBytesWritten As Long _
   ) As Long

Private Declare Function AVIStreamRelease Lib "avifil32.dll" (pavi As Any) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Long, lpBI As Any, ByVal wUsage As Long) As Long

Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Const OF_WRITE = &H1
Private Const OF_VERIFY = &H400
Private Const OF_SHARE_EXCLUSIVE = &H10
Private Const OF_SHARE_DENY_WRITE = &H20
Private Const OF_SHARE_DENY_READ = &H30
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OF_SHARE_COMPAT = &H0
Private Const OF_REOPEN = &H8000
Private Const OF_READWRITE = &H2
Private Const OF_READ = &H0
Private Const OF_PROMPT = &H2000
Private Const OF_PARSE = &H100
Private Const OF_EXIST = &H4000
Private Const OF_DELETE = &H200
Private Const OF_CREATE = &H1000
Private Const OF_CANCEL = &H800

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Declare Function DrawDibOpen Lib "MSVFW32.dll" () As Long
Private Declare Function DrawDibClose Lib "MSVFW32.dll" (ByVal hDD As Long) As Long
Private Declare Function DrawDibDraw Lib "MSVFW32.dll" (ByVal hDD As Long, ByVal hdc As Long, _
      ByVal xDst As Long, ByVal yDst As Long, ByVal dxDst As Long, ByVal dyDst As Long, _
      lpBI As Any, lpBits As Any, _
      ByVal xSrc As Long, ByVal ySrc As Long, ByVal dxSrc As Long, ByVal dySrc As Long, _
      ByVal wFlags As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

' Collection of cBmp objects to add to the
' AVI
Private m_colBmp As New Collection
Private m_lWidth As Long
Private m_lHeight As Long
Private m_sName As String
Private m_lHandler As Long
Private m_lType As Long
Private m_lDuration As Long
Private m_lBitsPerPixel As Long
Private m_sFilename As String
Private m_cPal As cPalette

Private m_tAVI As TAVISTREAMINFO
Private m_tACO As TAVICOMPRESSOPTIONS
Private m_hDrawDib As Long
Private m_pAs As Long
Private m_pGF As Long
Private m_pGFCompressed As Long
Private m_tBIH As BITMAPINFOHEADER
Private m_iStreamFrame As Long

' Gets/sets the palette if this is an 8pp AVI
Public Property Let Palette(cPal As cPalette)
   Set m_cPal = cPal
End Property
Public Property Set Palette(cPal As cPalette)
   Set m_cPal = cPal
End Property
Public Property Get Palette() As cPalette
   Set Palette = m_cPal
End Property

' Gets/sets the width of each frame in the AVI
Public Property Get Width() As Long
   Width = m_lWidth
End Property
Public Property Let Width(ByVal lWidth As Long)
   m_lWidth = lWidth
End Property
' Gets/sets the height of each frame in the AVI
Public Property Get Height() As Long
   Height = m_lHeight
End Property
Public Property Let Height(ByVal lHeight As Long)
   m_lHeight = lHeight
End Property
' Gets/sets the name that will eb written out to the AVI header
Public Property Get Name() As String
   Name = m_sName
End Property
Public Property Let Name(ByVal sName As String)
   m_sName = sName
End Property

' Gets/sets the Video Handler that will be used for compression.
' Set to 'mrle' for 8bpp AVIs, otherwise 'cvid'
Public Property Get VideoHandlerFourCC() As Long
   VideoHandlerFourCC = m_lHandler
End Property
Public Property Let VideoHandlerFourCC(ByVal lFourCC As Long)
   m_lHandler = lFourCC
End Property

' Gets/sets the video type.  Defaults to 'vids'
Public Property Get VideoTypeFourCC() As Long
   VideoTypeFourCC = m_lType
End Property
Public Property Let VideoTypeFourCC(ByVal lFourCC As Long)
   m_lType = lFourCC
End Property

' Gets the number of frames in the internal collection
Public Property Get FrameCount() As Long
   FrameCount = m_colBmp.Count
End Property

' Gets the bitmap at the frame with the specified index
Public Property Get Frame(Index As Variant) As cBmp
   Set Frame = m_colBmp.Item(Index)
End Property

' Removes the bitmap at the frame with the specified index
Public Sub RemoveFrame(Index As Variant)
   m_colBmp.Remove Index
End Sub

' Adds a new frame to include in the AVI when it is created.
Public Sub AddFrame(cb As cBmp, Optional key As Variant)
   m_colBmp.Add cb, key
End Sub

' Inserts a new frame to include in the AVI when it is created.
Public Sub InsertFrameBefore(cb As cBmp, keyBefore As Variant, Optional key As Variant)
   m_colBmp.Add cb, key, keyBefore
End Sub


' Inserts a new frame to include in the AVI when it is created.
Public Sub InsertFrameAfter(cb As cBmp, keyAfter As Variant, Optional key As Variant)
   m_colBmp.Add cb, key, , keyAfter
End Sub

' Gets/sets how long a frame will be
Public Property Get FrameDuration() As Long
   FrameDuration = m_lDuration
End Property
Public Property Let FrameDuration(ByVal lDuration As Long)
   m_lDuration = lDuration
End Property

' Gets/sets the number of Bits/Pixel to use when creating
' the AVI.  Set to either 8 (256 colours) or 24 (16 million+ colours)
Public Property Get bitsPerPixel() As Long
   bitsPerPixel = m_lBitsPerPixel
End Property
Public Property Let bitsPerPixel(ByVal lBitsPerPixel As Long)
   m_lBitsPerPixel = lBitsPerPixel
End Property

' Gets/sets filename to save AVI to
Public Property Get Filename() As String
   Filename = m_sFilename
End Property
Public Property Let Filename(ByVal sFileName As String)
   m_sFilename = sFileName
End Property

' Gets long value of FourCC value from string, e.g. 'vids'
Public Function FourCCFromString(ByVal sString As String) As Long
Dim lRet As Long
Dim sChar As String
Dim lChar As Long
   sChar = Mid(sString, 1)
   lRet = lRet Or Asc(sChar)
   sChar = Mid(sString, 2)
   lRet = lRet Or Asc(sChar) * &H100&
   sChar = Mid(sString, 3)
   lRet = lRet Or Asc(sChar) * &H10000
   sChar = Mid(sString, 4)
   lChar = Asc(sChar)
   lRet = lRet Or (lChar And &H7F&) * &H1000000
   If (lChar And &H80&) = &H80& Then
      lRet = lRet Or &H80000000
   End If
   FourCCFromString = lRet
End Function

' Gets the string from a FourCC long value
Public Function FourCCToString(ByVal lFourCC As Long) As String
Dim sRet As String
Dim lUByte As Long
   sRet = Chr(lFourCC And &HFF)
   sRet = sRet & Chr((lFourCC And &HFF00&) \ &H100&)
   sRet = sRet & Chr((lFourCC And &HFF0000) \ &H10000)
   lUByte = (lFourCC And &H7F000000) \ &H1000000
   If (lFourCC And &H80000000) = &H80000000 Then
      lUByte = lUByte Or &H80&
   End If
   sRet = sRet & Chr(lUByte)
   FourCCToString = sRet
End Function

' Creates a new AVI Stream on the FileName, and writes out
' the first image to it.
Public Sub StreamCreate(cFirstImage As cBmp)
Dim hR As Long

   m_iStreamFrame = 0
   m_lWidth = cFirstImage.Width
   m_lHeight = cFirstImage.Height

   hR = AVIFileOpen(m_pAs, m_sFilename, _
      OF_CREATE Or OF_WRITE, ByVal 0&)
   If (FAILED(hR)) Then
      pErr 2
   Else
      If createAVIStream() Then
         If createCompressedAVIStream() Then
            addFrameToAVIStream cFirstImage, m_iStreamFrame
         End If
      End If
   End If
   
End Sub

' Adds an image to the Stream created using StreamCreate
Public Sub StreamAdd(cImage As cBmp)

   If (m_pGFCompressed = 0) Then
      pErr 10
   Else
      m_iStreamFrame = m_iStreamFrame + 1
      addFrameToAVIStream cImage, m_iStreamFrame
   End If
   
End Sub

' Commits the content from the Stream created using StreamCreate
' and clears up any resources
Public Sub StreamClose()
   
   If Not (m_pGFCompressed = 0) Then
      AVIStreamRelease ByVal m_pGFCompressed
      m_pGFCompressed = 0
   End If
   
   If Not (m_pGF = 0) Then
      AVIStreamRelease ByVal m_pGF
      m_pGF = 0
   End If
   
   ' release avi stream
   If Not (m_pAs = 0) Then
      AVIStreamRelease ByVal m_pAs
      m_pAs = 0
   End If

End Sub

' Creates an AVI based on the collection created using the
' Add.. and Insert.. Frame methods
Public Sub CreateFromCollection()
Dim hR As Long

   hR = AVIFileOpen(m_pAs, m_sFilename, _
      OF_CREATE Or OF_WRITE, ByVal 0&)
   If (FAILED(hR)) Then
      pErr 2
   Else
      If createAVIStream() Then
         If createCompressedAVIStream() Then
            Dim cb As cBmp
            Dim Index As Long
            For Each cb In m_colBmp
               addFrameToAVIStream cb, Index
               Index = Index + 1
            Next
         End If
      End If
   End If

   If Not (m_pGFCompressed = 0) Then
      AVIStreamRelease ByVal m_pGFCompressed
      m_pGFCompressed = 0
   End If
   
   If Not (m_pGF = 0) Then
      AVIStreamRelease ByVal m_pGF
      m_pGF = 0
   End If
   
   ' release avi stream
   If Not (m_pAs = 0) Then
      AVIStreamRelease ByVal m_pAs
      m_pAs = 0
   End If

End Sub

Private Function createAVIStream() As Boolean
On Error Resume Next
Dim b() As Byte
Dim i As Long
Dim hR As Long
   
   b = StrConv(m_sName, vbFromUnicode)
   For i = 0 To UBound(b)
      m_tAVI.szName(i) = b(i)
   Next i
   For i = UBound(b) + 1 To 63
      m_tAVI.szName(i) = 0
   Next i
   m_tAVI.fccType = m_lType
   m_tAVI.fccHandler = m_lHandler
   m_tAVI.dwLength = m_colBmp.Count
   m_tAVI.dwScale = 100000
   m_tAVI.dwRate = m_tAVI.dwScale * 1000 / m_lDuration
   m_tAVI.rcFrame.Right = m_lWidth
   m_tAVI.rcFrame.Bottom = m_lHeight
   
   hR = AVIFileCreateStream(ByVal m_pAs, m_pGF, m_tAVI)
   If (FAILED(hR)) Then
      pErr 7
   Else
      createAVIStream = True
   End If
   
End Function

Private Function createCompressedAVIStream() As Boolean
Dim hR As Long

   m_tACO.fccType = m_lType
   m_tACO.fccHandler = m_lHandler
   
   hR = AVIMakeCompressedStream(m_pGFCompressed, _
      ByVal m_pGF, _
      m_tACO, _
      ByVal 0&)
   If (FAILED(hR)) Then
      pErr 8
   Else
      
      m_tBIH.biSize = 40
      m_tBIH.biWidth = m_lWidth
      m_tBIH.biHeight = m_lHeight
      m_tBIH.biPlanes = 1
      m_tBIH.biBitCount = m_lBitsPerPixel
      
      If (m_lBitsPerPixel = 24) Then
         m_tBIH.biSizeImage = BytesPerScanLine24(m_lWidth) * m_lHeight
         hR = AVIStreamSetFormat(ByVal m_pGFCompressed, ByVal 0&, _
            m_tBIH, LenB(m_tBIH))
      ElseIf (m_lBitsPerPixel = 8) Then
         m_tBIH.biSizeImage = BytesPerScanLine8(m_lWidth) * m_lHeight
         Dim tBI256 As BITMAPINFO256
         LSet tBI256.bmiHeader = m_tBIH
         m_cPal.ExtractToRGBQuadArray tBI256.bmiColors
         hR = AVIStreamSetFormat(ByVal m_pGFCompressed, ByVal 0&, _
            tBI256, LenB(tBI256))
      Else
         pErr 10
      End If
                     
      If (FAILED(hR)) Then
         pErr 9
      Else
         createCompressedAVIStream = True
      End If
            
   End If
   
End Function

Private Function addFrameToAVIStream(cb As cBmp, ByVal Index As Long) As Boolean
Dim lHDC As Long
Dim hMem As Long
Dim lPtrBits As Long
Dim lSamplesWritten As Long
Dim lBytesWritten As Long
Dim lR As Long
Dim hR As Long
Dim hPal As Long
Dim hPalOld As Long

   lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
            
   If (m_lBitsPerPixel = 24) Then
      
      hMem = GlobalAlloc(GPTR, m_tBIH.biSizeImage)
      lPtrBits = GlobalLock(hMem)

      lR = GetDIBits(lHDC, cb.hBmp, 0, m_lHeight, _
         ByVal lPtrBits, m_tBIH, DIB_RGB_COLORS)
      
   ElseIf (m_lBitsPerPixel = 8) Then
      
      Dim cD As New cDIBSection256
      cD.Create m_lWidth, m_lHeight
      cD.SetPalette m_cPal
      
      Dim cM As New cMemDC
      cM.Create
      cM.SelectObject cb
      cD.LoadPictureBlt cM.hdc
      cM.UnselectObject
      
      cD.PaintPicture Forms(0).hdc, 640 + m_lWidth
      
      lPtrBits = cD.DIBSectionBitsPtr
               
      'Dim tBI256 As BITMAPINFO256
      'm_cPal.ExtractToRGBQuadArray tBI256.bmiColors
      'LSet tBI256.bmiHeader = m_tBIH
      'hPalOld = SelectObject(lHDC, m_cPal.hPalette)
      'lR = GetDIBits(lHDC, cB.hBmp, 0, m_lHeight, _
      '   ByVal lPtrBits, tBI256, DIB_RGB_COLORS)
      'SelectObject lHDC, hPalOld
   End If
   
   hR = AVIStreamWrite(ByVal m_pGFCompressed, Index, 1, _
      ByVal lPtrBits, m_tBIH.biSizeImage, _
      0&, lSamplesWritten, lBytesWritten)
   If FAILED(hR) Then
      pErr 12
   Else
      addFrameToAVIStream = True
   End If
   
   If Not (hMem = 0) Then
      GlobalUnlock hMem
      GlobalFree hMem
   End If
   
   DeleteDC lHDC
   
   
End Function

Private Function BytesPerScanLine24(ByVal lWidth As Long) As Long
   BytesPerScanLine24 = (lWidth * 3 + 3) And &HFFFFFFFC
End Function
Private Function BytesPerScanLine8(ByVal lWidth As Long) As Long
   BytesPerScanLine8 = (lWidth + 3) And &HFFFFFFFC
End Function

Private Function FAILED(ByVal hR As Long) As Boolean
   FAILED = Not (SUCCEEDED(hR))
End Function

Private Function SUCCEEDED(ByVal hR As Long) As Boolean
   SUCCEEDED = ((hR And &H80000000) = 0)
End Function

Private Sub pErr(ByVal lErr As Long)
Dim sMsg As String
   
   Select Case lErr
   Case 1
      sMsg = "File not found"
   Case 2
      sMsg = "Failed to open AVI file"
   Case 3
      sMsg = "Unable to open AVI frames"
   Case 4
      sMsg = "AVI contains no frames"
   Case 5
      sMsg = "Could not read the format of the AVI"
   Case 6
      sMsg = "Unsupported AVI format"
   Case 7
      sMsg = "Failed to create a new AVI stream"
   Case 8
      sMsg = "Failed to create compressed AVI stream"
   Case 9
      sMsg = "Failed to set compressed AVI stream format"
   Case 10
      sMsg = "Unsupported bits/pixel setting: only 8 and 24 bpp AVIs are currently supported."
   Case 11
      sMsg = "Must have a valid open stream created using StreamCreate first."
   Case 12
      sMsg = "Failed to add image to the AVI Stream."
   Case Else
      sMsg = "Unexpected error " & lErr
   End Select
   
   Err.Raise ERR_BASE + lErr, App.EXEName & ".cAVIFrameExtract", sMsg
   
End Sub


Private Function ApiError(ByVal e As Long) As String
    Dim s As String, c As Long
    s = String(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      0, e, 0&, s, Len(s), ByVal 0)
    If c Then ApiError = Left$(s, c)
End Function


Private Sub Class_Initialize()
   AVIFileInit
   m_hDrawDib = DrawDibOpen()
   m_lType = FourCCFromString("vids")
End Sub

Private Sub Class_Terminate()
   StreamClose
   AVIFileExit
   DrawDibClose m_hDrawDib
End Sub
