VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
'*  VB file:   cDIB.cls... by Ray Mercer
'*  created:   12/1999 by Ray Mercer
'*  uploaded:  2/2000
'*  modified:  2/25/2000 by Ray Mercer
'*             Patrick Pasteels pointed out a bug in my code
'*             -fixed: ReDim m_memBitmapInfo(0 To 39) now correctly equals 40 bytes
'*
'*
'*  Copyright (C) 1999 - 2000 Ray Mercer.  All rights reserved.
'*  Latest version can be downloaded from http://www.shrinkwrapvb.com
'****************************************************************
Option Explicit

Private Const BMP_MAGIC_COOKIE As Integer = 19778 'this is equivalent to ascii string "BM"
'//BITMAP DEFINES (from mmsystem.h)
Private Type BITMAPFILEHEADER '14 bytes
        bfType As Integer '"magic cookie" - must be "BM"
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
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

Private Type RGBQUAD
    Red As Byte
    Green As Byte
    Blue As Byte
    Reserved As Byte
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'/* constants for the biCompression field */
Private Const BI_RGB  As Long = 0&
'#define BI_RLE8       1L
'#define BI_RLE4       2L
'#define BI_BITFIELDS  3L
'for use with AVIFIleInfo

'Private Type AVI_FILE_INFO  '108 bytes?
'    dwMaxBytesPerSecond As Long
'    dwFlags As Long
'    dwCaps As Long
'    dwStreams As Long
'    dwSuggestedBufferSize As Long
'    dwWidth As Long
'    dwHeight As Long
'    dwScale As Long
'    dwRate As Long
'    dwLength As Long
'    dwEditCount As Long
'    szFileType As String * 64
'End Type

'Private Declare Function CreateDIBSection_256 Lib "GDI32.DLL" Alias "CreateDIBSection" (ByVal hdc As Long, _
'                                                                                ByVal pbmi As BITMAPINFO_256, _
'                                                                                ByVal iUsage As Long, _
'                                                                                ByRef ppvBits As Long, _
'                                                                                ByVal hSection As Long, _
'                                                                                ByVal dwOffset As Long) As Long 'hBitmap
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long 'handle
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long 'Pointer to mem
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long 'BOOL
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal dwLen As Long)

Private Const HEAP_ZERO_MEMORY As Long = &H8

Private m_memBits() As Byte
Private m_memBitmapInfo() As Byte
Private m_bih As BITMAPINFOHEADER
Private m_bfh As BITMAPFILEHEADER


Public Function CreateFromFile(ByVal filename As String) As Boolean
    Dim hFile As Long
        
    If Not ExistFile(filename) Then
        MsgBox "File does not exist:" & vbCrLf & filename, vbCritical, App.Title
        Exit Function
    End If
        
    hFile = FreeFile()
    
    '<====ERROR TRAP ON
    On Error Resume Next
    Open filename For Binary Access Read As #hFile
    If err Then
        If err.Number = 70 Then
            MsgBox "File is locked - cannot access:" & vbCrLf & filename, vbCritical, App.Title
        Else
            MsgBox err.Description, vbInformation, App.Title
        End If
        Exit Function 'assume file was not opened
    End If
    On Error GoTo 0
    '====>ERROR TRAP OFF
    
    'OK, file is opened - now for the real algorithm...
    Get #hFile, , m_bfh 'get the BITMAPFILEHEADER this identifies the bitmap

    If m_bfh.bfType <> BMP_MAGIC_COOKIE Then 'this is not a BMP file
        MsgBox "File is not a supported bitmap format:" & vbCrLf & filename, vbInformation, App.Title
        Close #hFile
        Exit Function
    Else
        'now get the info header
        Get #hFile, Len(m_bfh) + 1, m_bih 'start at the 15th byte
        
        'now get the bitmap bits
        ReDim m_memBits(0 To m_bih.biSizeImage - 1)
        Get #hFile, m_bfh.bfOffBits + 1, m_memBits
        
        'and BitmapInfo variable-length UDT
        ReDim m_memBitmapInfo(0 To m_bfh.bfOffBits - 14) 'don't need first 14 bytes (fileinfo)
        Get #hFile, Len(m_bfh) + 1, m_memBitmapInfo
        
        Close #hFile   'Close file
    End If
    
    CreateFromFile = True 'indicate success
    
    
    
'    Debug.Print "BitCount: " & vbTab & vbTab & bih.biBitCount
'    Debug.Print "ClrImportant: " & vbTab & bih.biClrImportant
'    Debug.Print "ClrUsed: " & vbTab & vbTab & bih.biClrUsed
'    Debug.Print "Compression: " & vbTab & "&H" & Hex$(bih.biCompression)
'    Debug.Print "Height: " & vbTab & vbTab & bih.biHeight
'    Debug.Print "Planes: " & vbTab & vbTab & bih.biPlanes 'always 1
'    Debug.Print "Size: " & vbTab & vbTab & vbTab & bih.biSize
'    Debug.Print "SizeImage: " & vbTab & vbTab & bih.biSizeImage
'    Debug.Print "Width: " & vbTab & vbTab & vbTab & bih.biWidth
'    Debug.Print "XPelsPerMeter: " & vbTab & bih.biXPelsPerMeter 'usually 0
'    Debug.Print "YPelsPerMeter: " & vbTab & bih.biYPelsPerMeter 'usually 0

End Function

Public Function CreateFromPackedDIBPointer(ByRef pDIB As Long) As Boolean
On Error Resume Next
Debug.Assert pDIB <> 0
'Creates a full-color (no palette) DIB from a pointer to a full-color memory DIB

'get the BitmapInfoHeader
Call CopyMemory(ByVal VarPtr(m_bih.biSize), ByVal pDIB, Len(m_bih))
If m_bih.biBitCount < 16 Then
    Debug.Print "Error! DIB was less than 16 colors."
    Exit Function 'only supports high-color or full-color dibs
End If

'now get the bitmap bits
If m_bih.biSizeImage < 1 Then Exit Function 'return False
ReDim m_memBits(0 To m_bih.biSizeImage - 1)
Call CopyMemory(m_memBits(0), ByVal pDIB + 40, m_bih.biSizeImage)

'and BitmapInfo variable-length UDT
ReDim m_memBitmapInfo(0 To 39) 'don't need first 14 bytes (fileinfo)
Call CopyMemory(m_memBitmapInfo(0), m_bih, Len(m_bih))

'create a file header
With m_bfh
    .bfType = BMP_MAGIC_COOKIE
    .bfSize = 55 + m_bih.biSizeImage 'size of file as written to disk
    .bfReserved1 = 0&
    .bfReserved2 = 0&
    .bfOffBits = 54 'BitmapInfoHeader + BitmapFileHeader
End With

'and return True
CreateFromPackedDIBPointer = True

End Function

Public Function WriteToFile(ByVal filename As String) As Boolean
Dim hFile As Integer
On Error Resume Next
DoEvents
hFile = FreeFile()
Open filename For Binary As hFile
    Put hFile, 1, m_bfh
    Put hFile, Len(m_bfh) + 1, m_memBitmapInfo
    Put hFile, , m_memBits
Close hFile
WriteToFile = True
End Function

Private Function ExistFile(ByVal sSpec As String) As Boolean
    On Error Resume Next
    Call FileLen(sSpec)
    ExistFile = (err = 0)
End Function

Public Property Get BitCount() As Long
    BitCount = m_bih.biBitCount

End Property

Public Property Get Height() As Long
    Height = m_bih.biHeight
End Property

Public Property Get Width() As Long
    Width = m_bih.biWidth
End Property

Public Property Get Compression() As Long
    Compression = m_bih.biCompression
End Property

Public Property Get SizeInfoHeader() As Long
    SizeInfoHeader = m_bih.biSize
End Property

Public Property Get SizeImage() As Long
    SizeImage = m_bih.biSizeImage
End Property

Public Property Get Planes() As Long
    Planes = m_bih.biPlanes
End Property

Public Property Get ClrImportant() As Long
    ClrImportant = m_bih.biClrImportant
End Property

Public Property Get ClrUsed() As Long
    ClrUsed = m_bih.biClrUsed
End Property

Public Property Get XPPM() As Long
    XPPM = m_bih.biXPelsPerMeter
End Property

Public Property Get YPPM() As Long
    YPPM = m_bih.biYPelsPerMeter
End Property

Public Property Get FileType() As Long
    FileType = m_bfh.bfType
End Property

Public Property Get SizeFileHeader() As Long
    SizeFileHeader = m_bfh.bfSize
End Property

Public Property Get BitOffset() As Long
    BitOffset = m_bfh.bfOffBits
End Property

Public Property Get PointerToBits() As Long
    PointerToBits = VarPtr(m_memBits(0))
End Property

Public Property Get PointerToBitmapInfo() As Long
    PointerToBitmapInfo = VarPtr(m_memBitmapInfo(0))
End Property

Public Property Get SizeBitmapInfo() As Long
    SizeBitmapInfo = UBound(m_memBitmapInfo()) + 1
End Property
