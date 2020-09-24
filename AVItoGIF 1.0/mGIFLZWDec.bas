Attribute VB_Name = "mGIFLZWDec"
'================================================
' Module:        mGIFLZWDec.bas
' Author:        Vlad Vissoultchev (*)
' Dependencies:  cDIB.cls
' Last revision: -
'================================================

' (*)
'
'   From original work:
'
'   VB Gif Library Project (GIF87a/89a reader)
'   Copyright (c) 2003 Vlad Vissoultchev
'
'   Warning! use of this code in commercial applications may
'   fall under patent claims from Unisys which are holding
'   patents on LZW algorithm.


Option Explicit

'-- API:

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

'//

'-- Private Constants:

Private Const MAX_BITS                 As Long = 12
Private Const TABLE_SIZE               As Long = 2 ^ MAX_BITS

'-- Private Variables:

Private m_lInitBits                    As Long
Private m_lClearTable                  As Long
Private m_lInputBitCount               As Long
Private m_lInputBitBuffer              As Long
Private m_lCurrentBits                 As Long
Private m_lMaxCode                     As Long
Private m_lPrefixCode(0 To TABLE_SIZE) As Long
Private m_aAppendChar(0 To TABLE_SIZE) As Byte

Private m_aEncodedBytes()              As Byte
Private m_lByte                        As Long

Private m_aBits()                      As Byte
Private m_lImageSize                   As Long
Private m_lBytesWidth                  As Long
Private m_lImageHeight                 As Long

Private m_lInterlacedGroup             As Long
Private m_aInterlacedStep(0 To 3)      As Byte
Private m_aInterlacedInit(0 To 3)      As Byte

Private m_Pow2(-1 To 31)               As Long



'========================================================================================
' Methods
'========================================================================================

Public Sub InitPowers()

  Dim lPw As Long
    
    '-- Init look-up table for fast 2 ^ x
    m_Pow2(-1) = 0
    m_Pow2(0) = 1
    For lPw = 1 To 30
        m_Pow2(lPw) = 2 * m_Pow2(lPw - 1)
    Next
    m_Pow2(31) = &H80000000
End Sub

Public Sub Decode(oDIB08 As cDIB, ByVal IsInterlaced As Boolean, ByVal LZWCodeSize As Byte, EncodedBytes() As Byte)

  Dim tSA As SAFEARRAY2D
    
    If (oDIB08.BPP = [08_bpp]) Then
    
        '-- Store source encoded data
        m_aEncodedBytes() = EncodedBytes()
        
        '-- Get some image props.
        With oDIB08
            m_lImageSize = .Size
            m_lBytesWidth = .Width
            m_lImageHeight = .Height - 1
        End With
        
        '-- Initialize <Interlaced> mode vars.
        If (IsInterlaced And oDIB08.Height > 4) Then
            CopyMemory m_aInterlacedStep(0), &H2040808, 4
            CopyMemory m_aInterlacedInit(0), &H1020400, 4
          Else
            CopyMemory m_aInterlacedStep(0), &H1010101, 4
            CopyMemory m_aInterlacedInit(0), &H0, 4
        End If
        m_lInterlacedGroup = 0
        
        '-- Initialize LZW decoder vars.
        m_lInitBits = LZWCodeSize + 1
        m_lClearTable = m_Pow2(m_lInitBits - 1)
        m_lInputBitCount = 0
        m_lInputBitBuffer = 0
        m_lCurrentBits = m_lInitBits
        m_lMaxCode = m_Pow2(m_lInitBits) - 1
        m_lByte = 1
        
        '-- Map DIB bits
        pvBuildSA tSA, oDIB08
        CopyMemory ByVal VarPtrArray(m_aBits()), VarPtr(tSA), 4
        
        '-- Expand encoded data
        pvLZWExpand
        
        '-- Unmap DIBs
        CopyMemory ByVal VarPtrArray(m_aBits()), 0&, 4
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function pvLZWReadCode() As Long
    
    Do While m_lInputBitCount < m_lCurrentBits
        m_lInputBitBuffer = m_lInputBitBuffer Or (m_aEncodedBytes(m_lByte) * m_Pow2(m_lInputBitCount))
        m_lByte = m_lByte + 1
        m_lInputBitCount = m_lInputBitCount + 8
    Loop
    
    pvLZWReadCode = m_lInputBitBuffer And (m_Pow2(m_lCurrentBits) - 1)
    m_lInputBitBuffer = m_lInputBitBuffer \ m_Pow2(m_lCurrentBits)
    m_lInputBitCount = m_lInputBitCount - m_lCurrentBits
End Function

Private Function pvLZWDecodeString(aStack() As Byte, ByVal lIdx As Long, ByVal lCode As Long) As Long
    
    Do While lCode >= m_lClearTable
        aStack(lIdx) = m_aAppendChar(lCode)
        lIdx = lIdx + 1
        lCode = m_lPrefixCode(lCode)
    Loop
    
    aStack(lIdx) = lCode
    pvLZWDecodeString = lIdx
End Function

Private Sub pvLZWExpand()
    
    Dim x                 As Long
    Dim y                 As Long
    Dim lNewCode          As Long
    Dim lOldCode          As Long
    Dim lNextCode         As Long
    Dim aCharacter        As Byte
    Dim aStack(0 To 4000) As Byte
    Dim bClearFlag        As Boolean
    Dim lStackIdx         As Long
    
    On Error GoTo ErrH
    
    lNextCode = m_lClearTable + 2 ' First code = m_lClearTable + 2
    bClearFlag = -1
    lNewCode = pvLZWReadCode()
    
    Do: lNewCode = pvLZWReadCode()
        '-- Check for terminator
        If (lNewCode = m_lClearTable + 1) Then ' Terminator = m_lClearTable + 1
            Exit Sub
        End If
        
        If (bClearFlag) Then
        
            bClearFlag = 0
            lOldCode = lNewCode
            aCharacter = lNewCode
            
            m_aBits(x, y) = aCharacter
            
            x = x + 1
            If (x = m_lBytesWidth) Then
                x = 0
                y = y + m_aInterlacedStep(m_lInterlacedGroup)
                If (y > m_lImageHeight) Then
                    m_lInterlacedGroup = m_lInterlacedGroup + 1
                    y = m_aInterlacedInit(m_lInterlacedGroup)
                End If
            End If
            
        ElseIf (lNewCode = m_lClearTable) Then
            
            bClearFlag = -1
            m_lCurrentBits = m_lInitBits
            m_lMaxCode = m_Pow2(m_lCurrentBits) - 1
            lNextCode = m_lClearTable + 2 ' First code = m_lClearTable + 2
            
        Else
        
            '-- Decode string
            If (lNewCode < lNextCode) Then
                lStackIdx = pvLZWDecodeString(aStack, 0, lNewCode)
            ElseIf (lNewCode = lNextCode) Then
                aStack(0) = aCharacter
                lStackIdx = pvLZWDecodeString(aStack, 1, lOldCode)
            End If
            
            '-- Save first char
            aCharacter = aStack(lStackIdx)
            
            '-- Reverse copy stack
            Do: m_aBits(x, y) = aStack(lStackIdx)
                lStackIdx = lStackIdx - 1
                x = x + 1
                If (x = m_lBytesWidth) Then
                    x = 0
                    y = y + m_aInterlacedStep(m_lInterlacedGroup)
                    If (y > m_lImageHeight) Then
                        m_lInterlacedGroup = m_lInterlacedGroup + 1
                        y = m_aInterlacedInit(m_lInterlacedGroup)
                    End If
                End If
            Loop Until lStackIdx < 0
            
            '-- Keep char table up-to-date
            m_lPrefixCode(lNextCode) = lOldCode
            m_aAppendChar(lNextCode) = aCharacter
            lNextCode = lNextCode + 1
            
            '-- Expand code bitsize if max reached
            If (lNextCode > m_lMaxCode) Then
                If (m_lCurrentBits < MAX_BITS) Then
                    m_lCurrentBits = m_lCurrentBits + 1
                    m_lMaxCode = m_Pow2(m_lCurrentBits) - 1
                End If
            End If
            lOldCode = lNewCode
        End If
    Loop Until m_lByte = m_lImageSize
    
On Error GoTo 0
ErrH:
End Sub

Private Sub pvBuildSA(tSA As SAFEARRAY2D, oDIB08 As cDIB)

    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB08.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB08.BytesPerScanline
        .pvData = oDIB08.lpBits
    End With
End Sub

