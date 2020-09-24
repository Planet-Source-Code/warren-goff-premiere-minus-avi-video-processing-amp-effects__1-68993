VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cVideoHandlers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Private Declare Function ICLocate Lib "MSVFW32.dll" ( _
    ByVal fccType As Long, _
    ByVal fccHandler As Long, _
    lpbiIn As Any, _
    lpbOut As Any, _
    ByVal wFlags As Long _
    ) As Long

Public Enum EBitmapCompressionTypes
   BI_RGB = 0&
   BI_RLE8 = 1&
End Enum

Private Type TICInfo
    dwSize As Long ';                 // sizeof(ICINFO)
    fccType As Long ';                // compressor type     'vidc' 'audc'
    fccHandler As Long ';             // compressor sub-type 'rle ' 'jpeg' 'pcm '
    dwFlags As Long ';                // flags LOWORD is type specific
    dwVersion As Long ';              // version of the driver
    dwVersionICM As Long ';           // version of the ICM used
    '//
    '// under Win32, the driver always returns UNICODE strings.
    '//
    'WCHAR   szName[16];             // short name
    szName(0 To 31) As Byte
    'WCHAR   szDescription[128];     // long name
    szDescription(0 To 255) As Byte
    'WCHAR   szDriver[128];          // driver that contains compressor
    szDriver(0 To 255) As Byte
End Type

Private Declare Function ICInfo Lib "MSVFW32.dll" ( _
    ByVal fccType As Long, _
    ByVal fccHandler As Long, _
    lpicinfo As Any _
    ) As Long

Private Declare Function ICOpen Lib "MSVFW32.dll" ( _
    ByVal fccType As Long, _
    ByVal fccHandler As Long, _
    ByVal wMode As Long _
    ) As Long

Private Declare Function ICClose Lib "MSVFW32.dll" ( _
    ByVal hiC As Long _
    ) As Long

Private Declare Function ICGetInfo Lib "MSVFW32.dll" ( _
    ByVal hiC As Long, _
    picInfo As Any, _
    ByVal cb As Long _
    ) As Long

Private Declare Function ICSendMessage Lib "MSVFW32.dll" ( _
    ByVal hiC As Long, _
    ByVal msg As Long, _
    dw1 As Any, _
    ByVal dw2 As Long _
    ) As Long


Private Const ICMODE_COMPRESS = 1
Private Const ICMODE_DECOMPRESS = 2
Private Const ICMODE_FASTDECOMPRESS = 3
Private Const ICMODE_QUERY = 4
Private Const ICMODE_FASTCOMPRESS = 5
Private Const ICMODE_DRAW = 8

Private Const ICERR_OK = 0
Private Const ICERR_DONTDRAW = 1
Private Const ICERR_NEWPALETTE = 2
Private Const ICERR_GOTOKEYFRAME = 3
Private Const ICERR_STOPDRAWING = 4

Private Const ICERR_UNSUPPORTED = -1&
Private Const ICERR_BADFORMAT = -2&
Private Const ICERR_MEMORY = -3&
Private Const ICERR_INTERNAL = -4&
Private Const ICERR_BADFLAGS = -5&
Private Const ICERR_BADPARAM = -6&
Private Const ICERR_BADSIZE = -7&
Private Const ICERR_BADHANDLE = -8&
Private Const ICERR_CANTUPDATE = -9&
Private Const ICERR_ABORT = -10&
Private Const ICERR_ERROR = -100&
Private Const ICERR_BADBITDEPTH = -200&
Private Const ICERR_BADIMAGESIZE = -201&

Private Const ICERR_CUSTOM = -400&             '// errors less than ICERR_CUSTOM...


Private m_colHandlers As Collection


Public Function SuggestedVideoHandlerFourCC( _
      ByVal bitsPerPixel As Long _
   ) As Long
Dim bIh As BITMAPINFOHEADER
Dim hiC As Long
Dim tICI As TICInfo

   '// Initialize the bitmap structure.
   bIh.biSize = LenB(bIh)
   bIh.biPlanes = 1
   bIh.biCompression = BI_RGB      '// standard RGB bitmap
   If (bitsPerPixel = 8) Then
      bIh.biBitCount = 8              '// 8 bits-per-pixel format
      bIh.biClrUsed = 256
      bIh.biClrImportant = 256
   Else
      bIh.biBitCount = 24
   End If
    
   Dim ICTYPE_VIDEO As Long
   ICTYPE_VIDEO = FourCCFromString("vidc")
    
   hiC = ICLocate(ICTYPE_VIDEO, ByVal 0&, bIh, _
       ByVal 0&, ICMODE_COMPRESS)
       
   If Not (hiC = 0) Then
      ICGetInfo hiC, tICI, LenB(tICI)
      SuggestedVideoHandlerFourCC = tICI.fccHandler
   
      ICClose hiC
   End If

End Function

Private Sub getHandlers()
Dim i As Long
Dim iLast As Long
Dim iNext As Long
Dim tICBlank As TICInfo
Dim tIC As TICInfo
Dim tBIH As BITMAPINFOHEADER
Dim fccType As Long
Dim lPtrTic As Long
Dim hiC As Long
Dim sName As String
Dim sDescription As String
Dim lR As Long
Dim cV As cVideoHandler

   Set m_colHandlers = New Collection

   fccType = FourCCFromString("vidc")
   iLast = 0
   Do
      LSet tIC = tICBlank
      iNext = ICInfo(fccType, iLast, tIC)
      If Not (iNext = 0) Then
         hiC = ICOpen(tIC.fccType, tIC.fccHandler, ICMODE_QUERY)
         If Not (hiC = 0) Then
            ICGetInfo hiC, tIC, Len(tIC)
                        
            sName = tIC.szName
            sDescription = tIC.szDescription
            
            Set cV = New cVideoHandler
            cV.fInit sName, sDescription, tIC.fccHandler
            m_colHandlers.Add cV
            
            ICClose hiC
         End If
         iLast = iLast + 1
      Else
      End If
   Loop While (iNext > 0)
End Sub

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

Public Property Get HandlerCount() As Long
   HandlerCount = m_colHandlers.Count
End Property

Public Property Get Handler(ByVal Index As Long) As cVideoHandler
   Set Handler = m_colHandlers(Index)
End Property

Public Function IndexForFourCC(ByVal lFourCC As Long) As Long
Dim cVH As cVideoHandler
Dim i As Long
   For Each cVH In m_colHandlers
      i = i + 1
      If cVH.FourCC = lFourCC Then
         IndexForFourCC = i
      End If
   Next
End Function

Private Sub Class_Initialize()
   
   getHandlers
   
End Sub
