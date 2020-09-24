Attribute VB_Name = "mAVIImp"
'================================================
' Module:        mAVIImp.bas
' Author:        Steve McMahon, vbAccelerator (*)
' Dependencies:  cGIF.cls
' Last revision: 2003.07.12
'================================================
'
' (*)
'
'   From original work:
'   Transparent AVI Control
'   http://www.vbaccelerator.com/codelib/gfx/transavi.htm

Option Explicit

'-- API:

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type TAVISTREAMINFO ' this is the ANSI version
    fccType               As Long
    fccHandler            As Long
    dwFlags               As Long
    dwCaps                As Long
    wPriority             As Integer
    wLanguage             As Integer
    dwScale               As Long
    dwRate                As Long
    dwStart               As Long
    dwLength              As Long
    dwInitialFrames       As Long
    dwSuggestedBufferSize As Long
    dwQuality             As Long
    dwSampleSize          As Long
    rcFrame               As RECT
    dwEditCount           As Long
    dwFormatChangeCount   As Long
    szName(0 To 63)       As Byte
End Type

Private Const OF_READ            As Long = &H0
Private Const OF_SHARE_EXCLUSIVE As Long = &H10
Private Const streamtypeVIDEO    As Long = &H73646976 ' reads "vids"

'-- AVI functions:

Private Declare Sub AVIFileInit Lib "avifil32" ()
Private Declare Sub AVIFileExit Lib "avifil32" ()
Private Declare Function AVIStreamRelease Lib "avifil32" (pavi As Any) As Long
Private Declare Function AVIStreamOpenFromFile Lib "avifil32" Alias "AVIStreamOpenFromFileA" (ppavi As Any, ByVal szFile As String, ByVal fccType As Long, ByVal lParam As Long, ByVal mode As Long, pclsidHandler As Any) As Long
Private Declare Function AVIStreamGetFrameOpen Lib "avifil32" (pavi As Any, lpbiWanted As Any) As Long
Private Declare Function AVIStreamLength Lib "avifil32" (pavi As Any) As Long
Private Declare Function AVIStreamStart Lib "avifil32" (pavi As Any) As Long
Private Declare Function AVIStreamSampleToTime Lib "avifil32" (pavi As Any, ByVal lSample As Long) As Long
Private Declare Function AVIStreamGetFrameClose Lib "avifil32" (pg As Any) As Long
Private Declare Function AVIStreamGetFrame Lib "avifil32" (pg As Any, ByVal lPos As Long) As Long
Private Declare Sub AVIStreamInfo Lib "avifil32" Alias "AVIStreamInfoA" (pavi As Any, psi As TAVISTREAMINFO, ByVal lSize As Long)

'-- DrawDib functions:

Private Declare Function DrawDibOpen Lib "msvfw32" () As Long
Private Declare Function DrawDibClose Lib "msvfw32" (ByVal hDD As Long) As Long
Private Declare Function DrawDibDraw Lib "msvfw32" (ByVal hDD As Long, ByVal hDC As Long, ByVal xDst As Long, ByVal yDst As Long, ByVal dxDst As Long, ByVal dyDst As Long, lpBI As Any, lpBits As Any, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dxSrc As Long, ByVal dySrc As Long, ByVal wFlags As Long) As Long

'-- ...

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)



'========================================================================================
' Methods
'========================================================================================

Public Function ImportAVI(ByVal Filename As String, oGIF As cGIF, oProgress As ucProgress) As Boolean
    
  Dim lRet      As Long
  Dim lpGF      As Long
  Dim lpAS      As Long
  Dim tSI       As TAVISTREAMINFO
  Dim lpBI      As Long
  Dim lhDrawDIB As Long
  Dim oBuffDIB  As New cDIB
  Dim oBuffPal  As New cPal8bpp
  Dim AVIErr    As Boolean
  
  Dim lFrame    As Long
  Dim lFrames   As Long
  Dim lDelay    As Long
   
    '-- Initialization
    AVIFileInit
    lhDrawDIB = DrawDibOpen()
    
    If (lhDrawDIB <> 0) Then
    
        '-- Open AVI
        lRet = AVIStreamOpenFromFile(lpAS, Filename, streamtypeVIDEO, 0, OF_READ Or OF_SHARE_EXCLUSIVE, ByVal 0)
        
        If (Not pvFAILED(lRet)) Then
            
            If (lpAS <> 0) Then
            
                '-- Open frames
                lpGF = AVIStreamGetFrameOpen(ByVal lpAS, ByVal 0&)
                
                If (lpGF <> 0) Then
    
                    '-- Get number of frames
                    lFrames = AVIStreamLength(ByVal lpAS): oProgress.Max = lFrames
                    '-- Calculate timer delay
                    lDelay = pvAVIStreamEndTime(lpAS)
                    lDelay = lDelay / lFrames
                    '-- Get size of AVI animation screen -> tSI.rcFrame RECT
                    AVIStreamInfo ByVal lpAS, tSI, Len(tSI)
                    
                    '-- Build DIB32 buffer to dither from
                    oBuffDIB.Create tSI.rcFrame.Right, tSI.rcFrame.Bottom, [32_bpp]
                    
                    '-- Start frames adquisition/dithering
                    For lFrame = 1 To lFrames
                    
                        '-- Add new frame (Use current 'number of frames' + 1 for adding)
                        oGIF.FrameInsert lFrame, tSI.rcFrame.Right, tSI.rcFrame.Bottom
                        
                        '-- Select current frame
                        lpBI = AVIStreamGetFrame(ByVal lpGF, lFrame - 1)
                        
                        If (lpBI <> 0) Then
                            
                            oProgress = lFrame
                            '-- Dither...
                            lRet = DrawDibDraw(lhDrawDIB, oBuffDIB.hDC, 0, 0, oBuffDIB.Width, oBuffDIB.Height, ByVal lpBI, ByVal 0, 0, 0, -1, -1, 0)
                            mDither8bpp.Dither oBuffDIB, oGIF.FrameDIBXOR(lFrame), oBuffPal, (lFrame > 1)
                            '-- Set some GIF frame properties
                            CopyMemory ByVal oGIF.lpLocalPalette(lFrame), ByVal oBuffPal.lpPalette, 4 * oBuffPal.Entries
                            oGIF.FrameDisposalMethod(lFrame) = [dmRestoreToPrevious]
                            oGIF.FrameDelay(lFrame) = lDelay / 10
                          
                          Else
                            AVIErr = -1: Exit For
                        End If
                    Next lFrame
                    AVIStreamGetFrameClose ByVal lpGF
                    
                    '-- Define Screen Descriptor
                    oGIF.ScreenWidth = tSI.rcFrame.Right
                    oGIF.ScreenHeight = tSI.rcFrame.Bottom
                    oGIF.ScreenPixelAspectRatio = 1
                    
                    '-- Set global palette
                    oGIF.GlobalPaletteExists = -1
                    oGIF.GlobalPaletteEntries = oBuffPal.Entries
                    CopyMemory ByVal oGIF.lpGlobalPalette, ByVal oBuffPal.lpPalette, 4 * oBuffPal.Entries
                    
                    '-- Success [?]
                    oProgress = 0
                    ImportAVI = (AVIErr = 0)
                End If
                AVIStreamRelease ByVal lpAS
            End If
        End If
        DrawDibClose lhDrawDIB
    End If
    AVIFileExit
End Function

Private Function pvFAILED(ByVal lVal As Long) As Boolean
    pvFAILED = Not (pvSUCCEEDED(lVal))
End Function

Private Function pvSUCCEEDED(ByVal lVal As Long) As Boolean
    pvSUCCEEDED = ((lVal And &H80000000) = 0)
End Function

Private Function pvAVIStreamEndTime(ByVal lpAS As Long) As Long
  Dim lSample As Long
    lSample = AVIStreamStart(ByVal lpAS) + AVIStreamLength(ByVal lpAS)
    pvAVIStreamEndTime = AVIStreamSampleToTime(ByVal lpAS, lSample)
End Function
