VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==================================================================================
' cPalette.cls
' Copyright © 1999 Steve McMahon
'
' Visit vbAccelerator at http://vbaccelerator.com
'
' Creates and manages a series of palette entries (note
' this is independent of the Windows palette).
' Functions include Loading and Saving JASC format palette
' files, native palette files and creating various useful
' default palettes:
'  * Mono (black and white)
'  * 16 Colour EGA
'  * 256 colour Halftone (all colour combinations of &H40,&H80,&HC0 and &HFF)
'  * 256 colour Web safe palette (216 actual websafe colours)
'
' Use the palette in combination with cColourReduceDIB to create various
' colour reductions.
'
' Note: for best performance, when compiling an executable check
' all the boxes on the Properties-Compile tab Advanced Optimisations
' button, particularly Remove Array Bounds checks.
' ==================================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal lpvDest As Long, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryRef Lib "kernel32" Alias "RtlMoveMemory" ( _
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

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias _
   "VarPtr" (Ptr() As Any) As Long
Private Type tNode                              'Octree node struct
    bIsLeaf As Boolean                          'Leaf flag
    bAddedReduce As Boolean                     'Linked list flag
    vRed As Long                                'Red Value
    vGreen As Long                              'Green Value
    vBlue As Long                               'Blue Value
    cClrs As Long                               'Pixel count
    iChildren(0 To 1, 0 To 1, 0 To 1) As Long   'Child pointers
    iNext As Long                               'Next reducable node
End Type

Private aNodes() As tNode, cNodes As Long
Private nDepth As Byte, TopGarbage As Long
Private cClr As Long, aReduce(1 To 8) As Long
Private cOps As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
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

' make private storage in compatible
' with windows storage in memory/on disk
Private m_tPal() As RGBQUAD
Private m_iPalette As Long

Public Event InitProgress(ByVal Max As Long)
Public Event Progress(ByVal lPosition As Long)
Public Event Complete(ByVal lTimeMs As Long)


Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   'palPalEntry(1) As PALETTEENTRY
End Type

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As Any) As Long ' LOGPALETTE
Private Declare Function GetNearestPaletteIndex Lib "gdi32" (ByVal HPALETTE As Long, ByVal crColor As Long) As Long
Private Declare Function GetPaletteEntries Lib "gdi32" (ByVal HPALETTE As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long

Private m_hPal As Long

Friend Sub ExtractToRGBQuadArray(ByRef tRGBQuad() As RGBQUAD)
   CopyMemoryRef tRGBQuad(0), m_tPal(0), m_iPalette * LenB(tRGBQuad(0))
End Sub

Friend Sub CreateFromRGBQuadArray(ByRef tRGBQuad() As RGBQUAD)
   '
   ClearUp
   m_iPalette = UBound(tRGBQuad) - LBound(tRGBQuad) + 1
   ReDim m_tPal(0 To m_iPalette - 1) As RGBQUAD
   CopyMemoryRef m_tPal(0), tRGBQuad(0), m_iPalette * LenB(tRGBQuad(0))
   '
End Sub

Public Property Get HPALETTE() As Long
   HPALETTE = m_hPal
End Property

Public Sub CreatehPalette()
Dim tLP As LOGPALETTE
Dim tP As PALETTEENTRY

   ClearUp
   
   tLP.palVersion = &H300
   tLP.palNumEntries = m_iPalette
   ReDim b(0 To LenB(tLP) + m_iPalette * LenB(tP) - 1) As Byte
   CopyMemoryRef b(0), tLP, LenB(tLP)
   CopyMemoryRef b(LenB(tLP)), m_tPal(1), LenB(tP) * m_iPalette
   
   m_hPal = CreatePalette(b(0))
      
End Sub

Public Property Get ClosestIndex( _
      ByVal Red As Long, _
      ByVal Green As Long, _
      ByVal Blue As Long _
   ) As Long
Dim lMinIndex As Long
   
   If m_hPal Then
      
      Dim lRGB As Long
      
      lRGB = (Red And &HFF&) + (Green And &HFF&) * &H100& + (Blue And &HFF&) * &H10000
      lMinIndex = GetNearestPaletteIndex(m_hPal, lRGB)
   
   Else
   
      Dim i As Long
      Dim lER As Long, lEB As Long, lEG As Long
      Dim lMinER As Long, lMinEB As Long, lMinEG As Long
         
         lMinER = 255: lMinEB = 255: lMinEG = 255
         For i = 1 To m_iPalette
            With m_tPal(i)
               If (Red = .rgbRed) And (Blue = .rgbBlue) And (Green = .rgbGreen) Then
                  ClosestIndex = i
                  Exit Property
               Else
                  lER = Abs(Red - .rgbRed)
                  lEB = Abs(Blue - .rgbBlue)
                  lEG = Abs(Green - .rgbGreen)
                  If (lER + lEB + lEG < lMinER + lMinEB + lMinEG) Then
                     lMinER = lER
                     lMinEB = lEB
                     lMinEG = lEG
                     lMinIndex = i
                  End If
               End If
            End With
         Next i
   
   End If
   
   ClosestIndex = lMinIndex
      
End Property

Public Sub Add(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte)
    m_iPalette = m_iPalette + 1
    ReDim Preserve m_tPal(1 To m_iPalette) As RGBQUAD
    m_tPal(m_iPalette).rgbRed = Red
    m_tPal(m_iPalette).rgbGreen = Green
    m_tPal(m_iPalette).rgbBlue = Blue
End Sub

Public Property Get Count() As Long
' Number of colours in this palette:
   Count = m_iPalette
End Property
Public Property Get Red(ByVal iIndex) As Byte
' Red value for palette index
   Red = m_tPal(iIndex).rgbRed
End Property
Public Property Get Green(ByVal iIndex) As Byte
' Green value for palette index
   Green = m_tPal(iIndex).rgbGreen
End Property
Public Property Get Blue(ByVal iIndex) As Byte
' Blue value for palette index
   Blue = m_tPal(iIndex).rgbBlue
End Property

Public Sub CreateWebSafe()
Dim lIndex As Long
Dim r As Long, g As Long, b As Long
Dim l As Long, i As Long

   ' Websafe (IE 216 colour) palette
   m_iPalette = 256
   ReDim m_tPal(0 To 255) As RGBQUAD
   p16ColourLow8 0
   p16ColourHigh8 248
   lIndex = 8
   For b = 0 To &HFF Step &H33
      For g = 0 To &HFF Step &H33
         For r = 0 To &HFF Step &H33
            ' ignore if the output is any combination of 0 and FF
            l = r + g + b
            If l = 0 Or l = &H2FD Then
               ' ignore
            ElseIf l = &H1FE And (r = 0 Or g = 0 Or b = 0) Then
               ' ignore
            ElseIf l = &HFF And ((r = 0 And g = 0) Or (r = 0 And b = 0) Or (g = 0 And b = 0)) Then
               ' ignore
            Else
               ' add
               With m_tPal(lIndex)
                  .rgbRed = r: .rgbGreen = g: .rgbBlue = b
               End With
               lIndex = lIndex + 1
            End If
         Next
      Next
   Next
   ' Fill the remain entries with gray shades:
   r = 8: g = 8: b = 8
   For i = 217 To 247
      With m_tPal(lIndex)
         .rgbRed = r: .rgbGreen = g: .rgbBlue = b
         r = r + 8: g = g + 8: b = b + 8
      End With
   Next i
   
End Sub
Public Sub CreateHalfTone()
Dim lIndex As Long
Dim r As Long, g As Long, b As Long
Dim rA As Long, gA As Long, bA As Long
Dim l As Long, i As Long

   ' Halftone 256 colour palette
   m_iPalette = 256
   ReDim m_tPal(0 To 255) As RGBQUAD
   For b = 0 To &H100 Step &H40
      If b = &H100 Then
         bA = b - 1
      Else
         bA = b
      End If
      For g = 0 To &H100 Step &H40
         If g = &H100 Then
            gA = g - 1
         Else
            gA = g
         End If
         For r = 0 To &H100 Step &H40
            If r = &H100 Then
               rA = r - 1
            Else
               rA = r
            End If
            With m_tPal(lIndex)
               .rgbRed = rA: .rgbGreen = gA: .rgbBlue = bA
            End With
            lIndex = lIndex + 1
         Next r
      Next g
   Next b
   
End Sub
Public Sub CreateMono()
   ' Monochrome palette
   m_iPalette = 2
   ReDim m_tPal(1 To 2) As RGBQUAD
   With m_tPal(2)
      .rgbBlue = 255
      .rgbGreen = 255
      .rgbRed = 255
   End With
End Sub
Private Sub p16ColourLow8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 1
   With m_tPal(lStartIndex + 2)
      .rgbRed = &H80: .rgbGreen = 0: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 3)
      .rgbRed = 0: .rgbGreen = &H80: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 4)
      .rgbRed = &H80: .rgbGreen = &H80: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 5)
      .rgbRed = 0: .rgbGreen = 0: .rgbBlue = &H80
   End With
   With m_tPal(lStartIndex + 6)
      .rgbRed = &H80: .rgbGreen = 0: .rgbBlue = &H80
   End With
   With m_tPal(lStartIndex + 7)
      .rgbRed = 0: .rgbGreen = &H80: .rgbBlue = &H80
   End With
   With m_tPal(lStartIndex + 8)
      .rgbRed = &HC0: .rgbGreen = &HC0: .rgbBlue = &HC0
   End With

End Sub
Private Sub p16ColourHigh8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 9
   With m_tPal(lStartIndex + 9)
      .rgbRed = &H80: .rgbGreen = &H80: .rgbBlue = &H80
   End With
   With m_tPal(lStartIndex + 10)
      .rgbRed = &HFF: .rgbGreen = 0: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 11)
      .rgbRed = 0: .rgbGreen = &HFF: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 12)
      .rgbRed = &HFF: .rgbGreen = &HFF: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 13)
      .rgbRed = 0: .rgbGreen = 0: .rgbBlue = &HFF
   End With
   With m_tPal(lStartIndex + 14)
      .rgbRed = &HFF: .rgbGreen = 0: .rgbBlue = &HFF
   End With
   With m_tPal(lStartIndex + 15)
      .rgbRed = 0: .rgbGreen = &HFF: .rgbBlue = &HFF
   End With
   With m_tPal(lStartIndex + 16)
      .rgbRed = &HFF: .rgbGreen = &HFF: .rgbBlue = &HFF
   End With
End Sub
Public Sub Create16Colour()
   ' Standard EGA style 16 colour palette:
   m_iPalette = 16
   ReDim m_tPal(0 To 15) As RGBQUAD
   p16ColourLow8 0
   p16ColourHigh8 8
End Sub

Public Sub Clear()
   Erase m_tPal
   m_iPalette = 0
End Sub

Public Function Load(ByVal sFileName As String, Optional ByRef lPosition As Long) As Boolean
Dim hFile As Long
Dim lSize As Long
Dim lBytesRead As Long

   hFile = CreateFile(sFileName, _
                  GENERIC_READ, _
                  ByVal 0&, _
                  ByVal 0&, _
                  OPEN_EXISTING, _
                  FILE_ATTRIBUTE_NORMAL, _
                  ByVal 0&)
   If Not hFile = INVALID_HANDLE_VALUE Then
      If lPosition > 0 Then
         SetFilePointer hFile, lPosition, ByVal 0&, FILE_BEGIN
      Else
         lPosition = 0
      End If
      ReadFile hFile, lSize, 4, lBytesRead, ByVal 0&
      If lBytesRead = 4 Then
         If lSize > 0 And lSize <= 4097 Then
            m_iPalette = lSize
            ReDim m_tPal(1 To m_iPalette) As RGBQUAD
            lSize = LenB(m_tPal(1)) * m_iPalette
            ReadFile hFile, m_tPal(1), lSize, lBytesRead, ByVal 0&
            If lSize = lBytesRead Then
               lPosition = lPosition + 4 + lBytesRead
               Load = True
            End If
         End If
      End If
      CloseHandle hFile
   End If

End Function
Public Function Save(ByVal sFileName As String, Optional ByVal lPosition As Long) As Boolean
Dim hFile As Long
Dim lSize As Long
Dim lBytesWritten As Long

   hFile = CreateFile(sFileName, _
                  GENERIC_WRITE, _
                  ByVal 0&, _
                  ByVal 0&, _
                  CREATE_ALWAYS, _
                  FILE_ATTRIBUTE_NORMAL, _
                  ByVal 0&)
   If Not hFile = INVALID_HANDLE_VALUE Then
      If lPosition > 0 Then
         SetFilePointer hFile, lPosition, 0, FILE_BEGIN
      Else
         lPosition = 0
      End If
      WriteFile hFile, m_iPalette, 4, lBytesWritten, ByVal 0&
      If lBytesWritten = 4 Then
         lSize = LenB(m_tPal(1)) * m_iPalette
         WriteFile hFile, m_tPal(1), lSize, lBytesWritten, ByVal 0&
         If lSize = lBytesWritten Then
            lPosition = lPosition + 4 + lBytesWritten
            Save = True
         End If
      End If
      CloseHandle hFile
   End If
   
End Function
Public Function SaveToJASCFile(ByVal sFileName As String) As Boolean
Dim i As Long
Dim sJasc As String
Dim iFile As Long

   On Error Resume Next
   Kill sFileName
   
   sJasc = "JASC" & vbCrLf & "0100" & vbCrLf & m_iPalette
   For i = 1 To m_iPalette
      sJasc = sJasc & vbCrLf & m_tPal(i).rgbRed & " " & m_tPal(i).rgbGreen & " " & m_tPal(i).rgbBlue
   Next i
   
   On Error GoTo ErrorHandler
   iFile = FreeFile
   Open sFileName For Binary Access Write As #iFile
   Put #iFile, , sJasc
   Close #iFile
   iFile = 0
   Exit Function
   
ErrorHandler:
   Err.Raise Err.Number, App.EXEName & ".cPalette", Err.Description
   If (iFile <> 0) Then
      Close #iFile
   End If
   Exit Function
   
End Function

Public Function LoadFromJASCFile(ByVal sFileName As String) As Boolean
' Loads a JASC style .Pal file
' Not quick because it is a validating loader with error handling.
' The Native format is considerably quicker
Dim iFile As Long
Dim sBuf As String
Dim iPos As Long
Dim iNextPos As Long
Dim sLines() As String, iLineCount As Long, iLine As Long
Dim sParts() As String, iPartCount As Long
   
   m_iPalette = 0
   Erase m_tPal

   On Error GoTo LoadFromFileError
   ' Load it:
   iFile = FreeFile
   Open sFileName For Binary Access Read As #iFile
   sBuf = String$(LOF(iFile), 32)
   Get #iFile, , sBuf
   Close #iFile
   iFile = 0
   
   SplitDelimitedString sBuf, vbCrLf, sLines(), iLineCount
   For iLine = 1 To iLineCount
      SplitDelimitedString sLines(iLine), " ", sParts(), iPartCount
      If (iPartCount = 3) Then
         m_iPalette = m_iPalette + 1
         ReDim Preserve m_tPal(0 To m_iPalette - 1) As RGBQUAD
         With m_tPal(m_iPalette - 1)
            .rgbRed = CByte(sParts(1))
            .rgbGreen = CByte(sParts(2))
            .rgbBlue = CByte(sParts(3))
         End With
      End If
   Next iLine
   
   LoadFromJASCFile = True
   
   Exit Function

LoadFromFileError:
   Err.Raise Err.Number, App.EXEName & ".cPalette", Err.Description
   If (iFile <> 0) Then
      Close #iFile
      m_iPalette = 0
      Erase m_tPal
   End If
   Exit Function

End Function

Private Sub SplitDelimitedString( _
        ByVal sString As String, _
        ByVal sDelim As String, _
        ByRef sValues() As String, _
        ByRef iCount As Long _
    )
' ==================================================================
' Splits sString into an array of parts which are
' delimited in the string by sDelim.  The array is
' indexed 1-iCount where iCount is the number of
' items.  If no items found iCount=1 and the array has
' one element, the original string.
'   sString : String to split
'   sDelim  : Delimiter
'   sValues : Return array of values
'   iCount  : Number of items returned in sValues()
' ==================================================================
Dim iPos As Long
Dim iNextPos As Long
Dim iDelimLen As Long

    iCount = 0
    Erase sValues
    iDelimLen = Len(sDelim)
    iPos = 1
    iNextPos = InStr(sString, sDelim)
    Do While iNextPos > 0
        iCount = iCount + 1
        ReDim Preserve sValues(1 To iCount) As String
        sValues(iCount) = Mid$(sString, iPos, (iNextPos - iPos))
        iPos = iNextPos + iDelimLen
        iNextPos = InStr(iPos, sString, sDelim)
    Loop
    iCount = iCount + 1
    ReDim Preserve sValues(1 To iCount) As String
    sValues(iCount) = Mid$(sString, iPos)
End Sub



Public Sub ClearUp()
   If Not (m_hPal = 0) Then
      DeleteObject m_hPal
      m_hPal = 0
   End If
End Sub

Private Sub Class_Terminate()
   ClearUp
End Sub





