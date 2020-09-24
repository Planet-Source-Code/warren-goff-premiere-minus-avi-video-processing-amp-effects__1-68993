Attribute VB_Name = "mGIFExt"
'================================================
' Module:        mGIFExt.bas
' Author:        Carles P.V.
' Dependencies:  cGIF.cls
' Last revision: 2003.07.12
'================================================

Option Explicit

'-- API:

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

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

Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT2) As Long
Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (lpDst As Any, ByVal Length As Long, ByVal Fill As Byte)

'//

Public Function OptimizeGlobalPalette(oGIF As cGIF, oProgress As ucProgress) As Byte

  Dim nBefore     As Integer
  Dim aNPal(1023) As Byte
  Dim aGPal(1023) As Byte
  Dim bUsed()     As Boolean
  Dim aTrnE()     As Byte
  Dim aInvE()     As Byte
  Dim lIdx        As Long
  Dim lMax        As Long
  
  Dim nFrm        As Integer
  Dim tSA08       As SAFEARRAY2D
  Dim aBits08()   As Byte
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
    
    With oGIF
        
        oProgress.Max = 2 * .FramesCount
        
        If (.GlobalPaletteExists) Then
        
            '-- Store current number of entries and initialize arrays
            nBefore = .GlobalPaletteEntries
            ReDim bUsed(nBefore - 1)
            ReDim aTrnE(nBefore - 1)
            ReDim aInvE(nBefore - 1)
            ReDim aTransIdx(1 To .FramesCount)
            
            '-- Store Global palette as RGBQUAD byte array
            CopyMemory aGPal(0), ByVal .lpGlobalPalette, 1024
            
            '-- Check all used entries
            For nFrm = 1 To .FramesCount
                
                If (Not .LocalPaletteUsed(nFrm)) Then
                
                    oProgress = nFrm
                
                    '-- Get dimensions
                    W = .FrameDIBXOR(nFrm).Width - 1
                    H = .FrameDIBXOR(nFrm).Height - 1
                
                    '-- Map current 8-bpp DIB bits
                    pvBuild_08bppSA tSA08, .FrameDIBXOR(nFrm)
                    CopyMemory ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4
                    
                    '-- Check used entries...
                    For y = 0 To H
                        For x = 0 To W
                            bUsed(aBits08(x, y)) = -1
                        Next x
                    Next y
                    
                    '-- Unmap DIB bits
                    CopyMemory ByVal VarPtrArray(aBits08()), 0&, 4
                End If
            Next nFrm
        
            '-- 'Strecth' palette...
            For lIdx = 0 To .GlobalPaletteEntries - 1
                If (bUsed(lIdx)) Then
                    aTrnE(lMax) = lIdx ' New index
                    aInvE(lIdx) = lMax ' Inverse index
                    lMax = lMax + 1    ' Current count
                End If
            Next lIdx
            
            '-- Any entry removed [?]
            If (lMax < .GlobalPaletteEntries) Then
        
                If (lMax > 1) Then lMax = lMax - 1
                '-- Build temp. palette with only used entries,
                For lIdx = 0 To lMax
                    aNPal(4 * lIdx + 0) = aGPal(4 * aTrnE(lIdx) + 0)
                    aNPal(4 * lIdx + 1) = aGPal(4 * aTrnE(lIdx) + 1)
                    aNPal(4 * lIdx + 2) = aGPal(4 * aTrnE(lIdx) + 2)
                Next lIdx
                
                '-- set as new Global palette
                CopyMemory ByVal .lpGlobalPalette, aNPal(0), 1024
                .GlobalPaletteEntries = lMax + 1
                
                '-- and set as XOR DIBs palette (and update indexes)
                For nFrm = 1 To .FramesCount
                    
                    If (Not .LocalPaletteUsed(nFrm)) Then
                    
                        oProgress = .FramesCount + nFrm
                        
                        '-- Set new frame DIB palette
                        .FrameDIBXOR(nFrm).SetPalette aNPal()
                        
                        '-- Store temp. copy
                        CopyMemory ByVal .lpLocalPalette(nFrm), aNPal(0), 1024
                        .LocalPaletteEntries(nFrm) = lMax + 1
                        
                        '-- Get dimensions
                        W = .FrameDIBXOR(nFrm).Width - 1
                        H = .FrameDIBXOR(nFrm).Height - 1
            
                        '-- Map current 8-bpp DIB bits
                        pvBuild_08bppSA tSA08, .FrameDIBXOR(nFrm)
                        CopyMemory ByVal VarPtrArray(aBits08()), VarPtr(tSA08), 4
            
                        '-- Set new indexes...
                        For y = 0 To H
                            For x = 0 To W
                                aBits08(x, y) = aInvE(aBits08(x, y))
                            Next x
                        Next y
                        
                        '-- Unmap DIB bits
                        CopyMemory ByVal VarPtrArray(aBits08()), 0&, 4
                         
                        '-- Update transparent index and re-mask frame
                        .FrameTransparentColorIndex(nFrm) = aInvE(.FrameTransparentColorIndex(nFrm))
                        .FrameMask nFrm, .FrameTransparentColorIndex(nFrm)
                    End If
                Next nFrm
            End If
            oProgress = 0
            
            '-- Return removed entries
            OptimizeGlobalPalette = (nBefore - .GlobalPaletteEntries)
        End If
    End With
End Function

Public Function OptimizeFrames(oGIF As cGIF, oProgress As ucProgress) As Boolean
' Method: Minimum Bounding Rectangle (acceptable reduction).
' More info: http://www.webreference.com/dev/gifanim/frame.html
'
' Important:
' This routine only works for GIFs with full-screen frames and
' unique palette (Global): those GIFs got from AVI import rou-
' tine used here.

  Dim bInvalidGIF   As Boolean
                                                    
  Dim tSAPrev       As SAFEARRAY2D
  Dim tSANext       As SAFEARRAY2D
  Dim aBitsPrev()   As Byte
  Dim aBitsNext()   As Byte
  
  Dim bIsTrns       As Boolean
  Dim aTrnsIdx      As Byte
  Dim nFrm          As Integer
  Dim nRct          As Integer
  Dim rCrop()       As RECT2
  Dim oDIBBuff      As New cDIB
  Dim aPalXOR(1023) As Byte
  Dim aPalAND(7)    As Byte
  
  Dim x As Long, y As Long
  Dim W As Long, H As Long
    
    '==
    '== Validate GIF
    '==
    
    For nFrm = 1 To oGIF.FramesCount
        
        '-- Validate frames
        If (oGIF.FrameLeft(nFrm) <> 0 Or _
            oGIF.FrameTop(nFrm) <> 0 Or _
            oGIF.FrameDIBXOR(nFrm).Width <> oGIF.ScreenWidth Or _
            oGIF.FrameDIBXOR(nFrm).Height <> oGIF.ScreenHeight) Then
            
            bInvalidGIF = -1
            Exit For
        End If
    Next nFrm
    
    If (Not bInvalidGIF) Then
    
        '==
        '== Initialize
        '==
        
        '-- Redim. Crop rectangles
        ReDim rCrop(1 To oGIF.FramesCount)
        '-- Prepare palettes
        CopyMemory aPalXOR(0), ByVal oGIF.lpGlobalPalette, 1024
        FillMemory aPalAND(4), 3, &HFF
        
        '-- Set max. prog.
        oProgress.Max = oGIF.FramesCount
        
        '==
        '== Calc. minimum bounding rectangles
        '==
        
        '-- Transparent frames [?] (check first)
        bIsTrns = oGIF.FrameUseTransparentColor(1)
        aTrnsIdx = oGIF.FrameTransparentColorIndex(1)
        
        For nFrm = 1 + -(Not bIsTrns) To oGIF.FramesCount
            
            oProgress = nFrm
               
            '-- Map frame bits
            If (Not bIsTrns) Then pvBuild_08bppSA tSAPrev, oGIF.FrameDIBXOR(nFrm - 1)
            pvBuild_08bppSA tSANext, oGIF.FrameDIBXOR(nFrm)
            If (Not bIsTrns) Then CopyMemory ByVal VarPtrArray(aBitsPrev()), VarPtr(tSAPrev), 4
            CopyMemory ByVal VarPtrArray(aBitsNext()), VarPtr(tSANext), 4
            
            '-- Bounds:
            W = oGIF.FrameDIBXOR(nFrm).Width - 1
            H = oGIF.FrameDIBXOR(nFrm).Height - 1
            
            If (bIsTrns) Then
                
                '-- Top:
                For y = 0 To H
                    For x = 0 To W
                        If aBitsNext(x, y) <> aTrnsIdx Then rCrop(nFrm).y1 = y: GoTo Check_y2T
                    Next x
                Next y
Check_y2T:      '-- Bottom:
                For y = H To 0 Step -1
                    For x = 0 To W
                        If aBitsNext(x, y) <> aTrnsIdx Then rCrop(nFrm).y2 = y + 1: GoTo Check_x1T
                    Next x
                Next y
Check_x1T:      '-- Left:
                For x = 0 To W
                    For y = 0 To H
                        If aBitsNext(x, y) <> aTrnsIdx Then rCrop(nFrm).x1 = x: GoTo Check_x2T
                    Next y
                Next x
Check_x2T:      '-- Right:
                For x = W To 0 Step -1
                    For y = 0 To H
                        If aBitsNext(x, y) <> aTrnsIdx Then rCrop(nFrm).x2 = x + 1: GoTo NextFrameT
                    Next y
                Next x
            
NextFrameT: '-- End checking transparent frames
            Else
                
                '-- Top:
                For y = 0 To H
                    For x = 0 To W
                        If aBitsNext(x, y) <> aBitsPrev(x, y) Then rCrop(nFrm).y1 = y: GoTo Check_y2
                    Next x
                Next y
Check_y2:       '-- Bottom:
                For y = H To 0 Step -1
                    For x = 0 To W
                        If aBitsNext(x, y) <> aBitsPrev(x, y) Then rCrop(nFrm).y2 = y + 1: GoTo Check_x1
                    Next x
                Next y
Check_x1:       '-- Left:
                For x = 0 To W
                    For y = 0 To H
                        If aBitsNext(x, y) <> aBitsPrev(x, y) Then rCrop(nFrm).x1 = x: GoTo Check_x2
                    Next y
                Next x
Check_x2:       '-- Right:
                For x = W To 0 Step -1
                    For y = 0 To H
                        If aBitsNext(x, y) <> aBitsPrev(x, y) Then rCrop(nFrm).x2 = x + 1: GoTo NextFrame
                    Next y
                Next x
                
NextFrame:  '-- End checking not transparent frames
            End If
            
            '-- Unmap frame DIB bits
            If (Not bIsTrns) Then CopyMemory ByVal VarPtrArray(aBitsPrev()), 0&, 4
            CopyMemory ByVal VarPtrArray(aBitsNext()), 0&, 4
        Next nFrm
        
        '==
        '== Remove redundant frames...
        '==
        
        nFrm = 1 + -(Not bIsTrns)
        Do
            '-- Null rectangle [?]:
            If (IsRectEmpty(rCrop(nFrm))) Then
            
                '-- Add delay time of removed frame to previous frame
                oGIF.FrameDelay(nFrm - 1) = oGIF.FrameDelay(nFrm - 1) + oGIF.FrameDelay(nFrm)
                '-- Remove redundant frame
                oGIF.FrameRemove nFrm
                
                '-- Update temp. bounding rectangles array
                For nRct = nFrm To oGIF.FramesCount
                    rCrop(nRct) = rCrop(nRct + 1)
                Next nRct
                
              Else
                '-- Next frame
                nFrm = nFrm + 1
            End If
        Loop Until nFrm > oGIF.FramesCount
        
        '==
        '== Crop frames...
        '==
        
        '-- First frame (not transparent)
        If (Not oGIF.FrameUseTransparentColor(1)) Then
            oGIF.FrameDisposalMethod(1) = [dmDoNotDispose]
        End If
        
        '-- Next frames
        For nFrm = 1 + -(Not bIsTrns) To oGIF.FramesCount
    
            '-- XOR DIB
            With oGIF.FrameDIBXOR(nFrm)
                .SetPalette aPalXOR()
                .CloneTo oDIBBuff
                .Create rCrop(nFrm).x2 - rCrop(nFrm).x1, rCrop(nFrm).y2 - rCrop(nFrm).y1, [08_bpp]
                .SetPalette aPalXOR()
                .LoadBlt oDIBBuff.hDC, rCrop(nFrm).x1, rCrop(nFrm).y1
            End With
            '-- AND DIB
            With oGIF.FrameDIBAND(nFrm)
                .Create rCrop(nFrm).x2 - rCrop(nFrm).x1, rCrop(nFrm).y2 - rCrop(nFrm).y1, [01_bpp]
                .SetPalette aPalAND()
            End With
            '-- Remask frame
            oGIF.FrameMask nFrm, oGIF.FrameTransparentColorIndex(nFrm)
            
            '-- Frame position
            oGIF.FrameLeft(nFrm) = rCrop(nFrm).x1
            oGIF.FrameTop(nFrm) = rCrop(nFrm).y1
            
            '-- Frame disposal method
            If (oGIF.FrameUseTransparentColor(nFrm)) Then
                oGIF.FrameDisposalMethod(nFrm) = [dmRestoreToPrevious]
              Else
                oGIF.FrameDisposalMethod(nFrm) = [dmDoNotDispose]
            End If
        Next nFrm
        
        '-- End/Success
        oProgress = 0
        OptimizeFrames = -1
    End If
End Function

Public Sub RemaskFrames(oGIF As cGIF, ByVal Transparent As Boolean, ByVal TransparentColorIndex As Byte, oProgress As ucProgress)
  
  Dim nFrm As Integer
  
    oProgress.Max = oGIF.FramesCount
    
    '-- Re-mask frames
    For nFrm = 1 To oGIF.FramesCount
        oProgress = nFrm
        oGIF.FrameUseTransparentColor(nFrm) = Transparent
        oGIF.FrameTransparentColorIndex(nFrm) = TransparentColorIndex
        oGIF.FrameMask nFrm, TransparentColorIndex
    Next nFrm
    oProgress = 0
End Sub

'//

Private Sub pvBuild_08bppSA(tSA As SAFEARRAY2D, oDIB As cDIB)

    '-- 8-bpp DIB mapping
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.BytesPerScanline
        .pvData = oDIB.lpBits
    End With
End Sub
