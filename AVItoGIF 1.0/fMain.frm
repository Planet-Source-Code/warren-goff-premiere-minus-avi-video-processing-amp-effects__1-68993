VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AVItoGIF"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7695
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   513
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox iPalette 
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      ForeColor       =   &H00808080&
      Height          =   1500
      Left            =   6000
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1335
      Width           =   1500
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Left            =   5220
      Top             =   4620
   End
   Begin VB.CommandButton cmdPickColor 
      Caption         =   "Pick &color"
      Enabled         =   0   'False
      Height          =   405
      Left            =   6000
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1170
   End
   Begin VB.CheckBox chkTransparent 
      Appearance      =   0  'Flat
      Caption         =   "&Transparent"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6000
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1335
   End
   Begin AVItoGIF.ucCanvas ucCanvas 
      Height          =   5025
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   165
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   8864
   End
   Begin AVItoGIF.ucProgress ucProgress 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      Top             =   5325
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   397
      BorderStyle     =   1
   End
   Begin VB.Label lblEntriesV 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H80000015&
      Height          =   210
      Left            =   6375
      TabIndex        =   9
      Top             =   2865
      Width           =   1125
   End
   Begin VB.Label lblPaletteV 
      Height          =   195
      Left            =   6675
      TabIndex        =   7
      Top             =   1005
      Width           =   900
   End
   Begin VB.Label lblPalette 
      Caption         =   "Palette:"
      Height          =   195
      Left            =   6015
      TabIndex        =   6
      Top             =   1005
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Note: Enable/disable transparency before GIF optimization."
      ForeColor       =   &H80000015&
      Height          =   660
      Left            =   6000
      TabIndex        =   12
      Top             =   4590
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblScreenV 
      Height          =   195
      Left            =   6675
      TabIndex        =   3
      Top             =   450
      Width           =   900
   End
   Begin VB.Label lblFramesV 
      Height          =   195
      Left            =   6675
      TabIndex        =   5
      Top             =   720
      Width           =   900
   End
   Begin VB.Label lblScreen 
      Caption         =   "Screen:"
      Height          =   195
      Left            =   6015
      TabIndex        =   2
      Top             =   450
      Width           =   735
   End
   Begin VB.Label lblFrames 
      Caption         =   "Frames:"
      Height          =   195
      Left            =   6015
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblInfo 
      Caption         =   "GIF info:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6015
      TabIndex        =   1
      Top             =   150
      Width           =   1260
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Import AVI..."
         Index           =   0
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save test GIF"
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Optimize GIF"
         Index           =   3
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   5
      End
   End
   Begin VB.Menu mnuOptionsTop 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "Ordered dither"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Use optimal (first frame)"
         Checked         =   -1  'True
         Index           =   1
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Project:       AVItoGIF
' Author:        Carles P.V. (*)
' Last revision: 2003.09.06
'================================================
' Commercial use not permitted!
' Email author/s please.
'================================================
'
' (*) All thanks to Vlad Vissoultchev & Ron van Tilburg
'     for GIF Decode/Encode original routines.
'
' Notes:
'
' This 'AVI to GIF' converter is quite 'simple':
' - Optimal palette is got from the first frame.
' - Simple GIF size optimization (M.B.R.).
'
' I'm working now on better GIF optimization. I'll post
' it as soon as possible (I hope).
'
' Please, let me know for any bug/s, speed improvements,
' etc. Thanks.



Option Explicit

Private m_oGIF            As New cGIF  ' Our GIF object
Private m_oBackground     As New cTile ' Frame rendering
Private m_oDIBRestore     As New cDIB  ' Frame rendering
Private m_bTransparent    As Boolean   ' Transparency
Private m_nTransparentIdx As Integer   ' Transparent palette entry
Private m_nFrame          As Integer   ' Current frame
Private m_nFrames         As Integer   ' Number of frames

Private m_sFilename       As String    ' Last file path (AVI import)
Private m_bPicking        As Boolean   ' Picking transparent color
Dim AVIFile As String
Dim sTmpFilename As String

Private Sub Form_Activate()
Dim Strng As String
mnuOptions_Click (0)
mnuOptions_Click (0)
    If Command$ <> "" And Left(Command$, 1) <> "1" Then
        AVIFile = App.Path & "\MyAVI.avi"
        ImportAVI
        Optimum
        SaveGif
    End If
    If Command$ <> "" And Left(Command$, 1) = "1" Then
        AVIFile = App.Path & "\MyAVIalt.avi"
        ImportAVI
        Optimum
        SaveGif
        Unload Me
    End If

End Sub

'//

Private Sub Form_Load()
SetTopMostWindow Me.hWnd, True
    '-- Initalize mDither8bpp module
    mDither8bpp.InitializeLUTs
    mDither8bpp.Palette = [ipBrowser]
    mDither8bpp.DitherMethod = [idmNone]
    
    '-- Initalize pattern brush (empty palette entry)
    mMisc.InitializePatternBrush
            
    '-- Load canvas custom cursor
    Set ucCanvas.UserIcon = LoadResPicture("CURSOR_PICKCOLOR", vbResCursor)
    
    '-- Hook mouse wheel for zooming support
    mWheel.HookWheel
End Sub

Private Sub Form_Paint()

    '-- Some decorative lines
    Me.Line (0, 0)-(ScaleWidth, 0), vb3DShadow
    Me.Line (0, 1)-(ScaleWidth, 1), vb3DHighlight
End Sub

Private Sub Form_Unload(Cancel As Integer)

    '-- Destroy GIF and rendering buffers
    pvCleanUp
    '-- Destroy pattern brush (empty palette entry)
    mMisc.DestroyPatternBrush
    
    '-- Is next line necessary [?]
    Set fMain = Nothing
    End
End Sub

'//

Private Sub mnuFile_Click(Index As Integer)
    
  
    Select Case Index
    
        Case 0 '-- Import AVI...
        
            '-- Show open file dialog
            sTmpFilename = mDialogFile.GetFileName(m_sFilename, "AVI files (*.avi)|*.AVI", , "Load AVI file", -1)
            'sTmpFilename = AVIFile
            If (Len(sTmpFilename)) Then
                m_sFilename = sTmpFilename
                
                '-- Stop animation
                tmrDelay.Enabled = 0
            
                '-- Destroy current GIF
                m_oGIF.Destroy
                
                '-- Disable transparency
                chkTransparent.Enabled = -1
                chkTransparent = 0
                cmdPickColor.Enabled = 0
                m_bTransparent = 0
                m_nTransparentIdx = 0
                
                '-- Import AVI frames...
                DoEvents
                If (mAVIImp.ImportAVI(m_sFilename, m_oGIF, ucProgress)) Then
                    pvInitialize
                    pvShowInfo
                  Else
                    MsgBox "Unexpected error loading AVI file.", vbExclamation
                    pvCleanUp
                    pvShowInfo
                End If
            End If
             
        Case 1 '-- Save GIF...
            
            If (m_oGIF.FramesCount = 0) Then
                
                '-- No GIF
                MsgBox "Nothing to save", vbExclamation
              
              Else
                '-- Save as test file (Test.gif)
                DoEvents
                Screen.MousePointer = vbArrowHourglass
                If (Not m_oGIF.Save(App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "Test.gif")) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Unexpected error saving GIF file.", vbExclamation
                  Else
                    Screen.MousePointer = vbDefault
                End If
            End If
            
        Case 3 ' -- Optimize GIF
            
            If (m_oGIF.FramesCount = 0) Then
            
                '-- Nothing to optimize
                MsgBox "Nothing to optimize", vbExclamation
                
              Else
                '-- Remove unused entries and get minimum bounding rectangles
                tmrDelay.Enabled = 0
                DoEvents
                mGIFExt.OptimizeGlobalPalette m_oGIF, ucProgress: m_nTransparentIdx = m_oGIF.FrameTransparentColorIndex(1)
                mGIFExt.OptimizeFrames m_oGIF, ucProgress
                '-- Initialize
                pvInitialize
                pvShowInfo
                '-- Disable transparency controls
                chkTransparent.Enabled = 0
                cmdPickColor.Enabled = 0
                '-- Disable Picking mode [?]
                If (m_bPicking) Then
                    m_bPicking = 0
                    ucCanvas.WorkMode = [cnvScrollMode]
                End If
            End If
            
        Case 5 ' -- Exit
            Unload Me
    End Select
End Sub
Private Sub Optimum()
                tmrDelay.Enabled = 0
                DoEvents
                mGIFExt.OptimizeGlobalPalette m_oGIF, ucProgress: m_nTransparentIdx = m_oGIF.FrameTransparentColorIndex(1)
                mGIFExt.OptimizeFrames m_oGIF, ucProgress
                '-- Initialize
                pvInitialize
                pvShowInfo
                '-- Disable transparency controls
                chkTransparent.Enabled = 0
                cmdPickColor.Enabled = 0
                '-- Disable Picking mode [?]
                If (m_bPicking) Then
                    m_bPicking = 0
                    ucCanvas.WorkMode = [cnvScrollMode]
                End If
End Sub
Private Sub SaveGif()
            If (m_oGIF.FramesCount = 0) Then
                
                '-- No GIF
                MsgBox "Nothing to save", vbExclamation
              
              Else
                '-- Save as test file (Test.gif)
                DoEvents
                Screen.MousePointer = vbArrowHourglass
                If (Not m_oGIF.Save(App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "Test.gif")) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "Unexpected error saving GIF file.", vbExclamation
                  Else
                    Screen.MousePointer = vbDefault
                End If
            End If

End Sub
Private Sub ImportAVI()
            '-- Show open file dialog
            'sTmpFilename = mDialogFile.GetFileName(m_sFilename, "AVI files (*.avi)|*.AVI", , "Load AVI file", -1)
            sTmpFilename = AVIFile
            If (Len(sTmpFilename)) Then
                m_sFilename = sTmpFilename
                
                '-- Stop animation
                tmrDelay.Enabled = 0
            
                '-- Destroy current GIF
                m_oGIF.Destroy
                
                '-- Disable transparency
                chkTransparent.Enabled = -1
                chkTransparent = 0
                cmdPickColor.Enabled = 0
                m_bTransparent = 0
                m_nTransparentIdx = 0
                
                '-- Import AVI frames...
                DoEvents
                If (mAVIImp.ImportAVI(m_sFilename, m_oGIF, ucProgress)) Then
                    pvInitialize
                    pvShowInfo
                  Else
                    MsgBox "Unexpected error loading AVI file.", vbExclamation
                    pvCleanUp
                    pvShowInfo
                End If
            End If

End Sub
Private Sub mnuOptions_Click(Index As Integer)

    Select Case Index
    
        Case 0 '-- Ordered dither
            With mnuOptions(0)
                .Checked = Not .Checked
                mDither8bpp.DitherMethod = -.Checked * [idmOrdered]
            End With
            
        Case 1 '-- Use optimal
            With mnuOptions(1)
                .Checked = Not .Checked
                mDither8bpp.Palette = -.Checked * [ipOptimal]
            End With
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
                
    '-- Simple About box...
    MsgBox "AVItoGIF v" & App.Major & "." & App.Minor & vbCrLf & vbCrLf & _
           "Simple AVI to GIF converter + Basic GIF size optimization" & vbCrLf & vbCrLf & _
           "All thanks to Vlad Vissoultchev & Ron van Tilburg" & vbCrLf & _
           "for GIF Decode/Encode original routines."
End Sub

'//

Private Sub chkTransparent_Click()
  
  Dim nFrm As Integer
    
    If (m_oGIF.FramesCount) Then
    
        '-- Update controls
        m_bTransparent = -chkTransparent
        cmdPickColor.Enabled = -chkTransparent
    
        '-- Re-mask frames
        Screen.MousePointer = vbHourglass
        mGIFExt.RemaskFrames m_oGIF, m_bTransparent, m_nTransparentIdx, ucProgress
        Screen.MousePointer = vbDefault
        
        '-- Re-start animation
        pvInitialize
    End If
End Sub

Private Sub cmdPickColor_Click()

    '-- Enable color picking
    If (m_bTransparent) Then
        m_bPicking = -1
        ucCanvas.WorkMode = [cnvUserMode]
        
        '-- Show first (full) frame
        tmrDelay.Enabled = 0
        m_nFrame = 1
        pvFrame_Change
    End If
End Sub

'//

Private Sub ucCanvas_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    '-- Force ucCanvas_MouseMove sub.
    If (m_bPicking And ucCanvas.DIB.hDIB <> 0) Then
        ucCanvas_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub ucCanvas_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
    If (Button = vbLeftButton) Then
    
        If (m_bPicking And m_oGIF.FramesCount) Then
            
            With m_oGIF
            
                If (X >= .FrameLeft(1) And _
                    Y >= .FrameTop(1) And _
                    X < .FrameDIBXOR(1).Width - .FrameLeft(1) And _
                    Y < .FrameDIBXOR(1).Height - .FrameTop(1)) Then
                    
                    '-- Get palette index (NOT color)
                    m_nTransparentIdx = mDither8bpp.PaletteIndex(.FrameDIBXOR(m_nFrame), X + .FrameLeft(1), Y + .FrameTop(1))
                End If
            End With
        End If
    End If
End Sub

Private Sub ucCanvas_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    
  Dim nFrm As Integer
  Dim lClr As Long
    
    Select Case Button
    
        Case vbLeftButton
        
            If (m_bPicking) Then
                m_bPicking = 0
                ucCanvas.WorkMode = [cnvScrollMode]
                
                '-- Re-mask frames
                Screen.MousePointer = vbHourglass
                mGIFExt.RemaskFrames m_oGIF, m_bTransparent, m_nTransparentIdx, ucProgress
                Screen.MousePointer = vbDefault
                
                '-- Re-start animation
                pvInitialize
            End If
        
        Case vbRightButton
            
            '-- Change background color
            lClr = mDialogColor.SelectColor(Me.hWnd, ucCanvas.BackColor)
            If (lClr <> -1) Then
                ucCanvas.BackColor = lClr
                If (m_oGIF.FramesCount) Then pvInitialize
            End If
    End Select
End Sub

'//

Private Sub pvCleanUp()
    
    Set m_oGIF = Nothing
    Set ucCanvas.DIB = Nothing
    Set m_oBackground = Nothing
    Set m_oDIBRestore = Nothing
End Sub

Private Sub pvInitialize()

    '-- Stop timer
    tmrDelay.Enabled = 0
    
    '-- Paint palette
    pvPaintPalette

    '-- Initialize buffers
    With m_oGIF
    
        '-- Get number of frames
        m_nFrames = .FramesCount
        
        '-- Create canvas DIB
        ucCanvas.DIB.Create .ScreenWidth, .ScreenHeight, [32_bpp]
        ucCanvas.Resize
        '-- Create restoring DIB
        m_oDIBRestore.Create .ScreenWidth, .ScreenHeight, [24_bpp]
        '-- Create background pattern (solid color) and initialize DIBs
        m_oBackground.SetPatternFromSolidColor ucCanvas.BackColor
        m_oBackground.Tile ucCanvas.DIB.hDC, 0, 0, .ScreenWidth, .ScreenHeight
        m_oBackground.Tile m_oDIBRestore.hDC, 0, 0, .ScreenWidth, .ScreenHeight
       
        '-- Start animation [?]
        If (m_nFrames > 1) Then
            '-- Enable timer
            m_nFrame = m_nFrames
            tmrDelay.Interval = 1
            tmrDelay.Enabled = -1
          Else
            '-- Render only first frame
            .FrameDraw ucCanvas.DIB.hDC, 1: ucCanvas.Repaint
        End If
    End With
End Sub

Private Sub tmrDelay_Timer()
    
    '-- Next frame / First
    If (m_nFrame < m_nFrames) Then
        m_nFrame = m_nFrame + 1
      Else
        m_nFrame = 1
    End If
    pvFrame_Change
End Sub

Private Sub pvFrame_Change()
    
    With m_oGIF
        
        '-- Set current frame delay
        Select Case .FrameDelay(m_nFrame)
            Case Is < 0
                tmrDelay.Interval = 60000 ' Max.: 1 min.
            Case Is = 0
                tmrDelay.Interval = 100   ' Def.: 0.1 sec.
            Case Is < 5
                tmrDelay.Interval = 50    ' Min.: 0.05 sec.
            Case Else
                tmrDelay.Interval = .FrameDelay(m_nFrame) * 10
        End Select
        
        '-- Restore:
        If (m_nFrame = 1) Then
            m_oBackground.Tile ucCanvas.DIB.hDC, 0, 0, .ScreenWidth, .ScreenHeight
          Else
            ucCanvas.DIB.LoadBlt m_oDIBRestore.hDC
        End If
        
        '-- Draw current frame:
        .FrameDraw ucCanvas.DIB.hDC, m_nFrame
        
        '-- Update restoring buffer:
        Select Case .FrameDisposalMethod(m_nFrame)
            Case [dmNotSpecified], [dmDoNotDispose]
                '-- Update from current
                m_oDIBRestore.LoadBlt ucCanvas.DIB.hDC
            Case [dmRestoreToBackground]
                '-- Update from background
                m_oBackground.Tile m_oDIBRestore.hDC, .FrameLeft(m_nFrame), .FrameTop(m_nFrame), .FrameDIBXOR(m_nFrame).Width, .FrameDIBXOR(m_nFrame).Height, 0
            Case [dmRestoreToPrevious]
                '-- Preserve buffer
        End Select
    End With
    
    '-- Paint frame
    ucCanvas.Repaint
End Sub

Private Sub pvShowInfo()
    
  Dim aBPP As Byte
  
    Select Case True
    
        Case ucCanvas.DIB.hDIB <> 0
            '-- Calc. GIF palette color depth
            Do: aBPP = aBPP + 1
            Loop Until 2 ^ aBPP >= m_oGIF.GlobalPaletteEntries
            '-- Show AVI props.
            lblScreenV.Caption = m_oGIF.ScreenWidth & "x" & m_oGIF.ScreenHeight
            lblFramesV.Caption = m_oGIF.FramesCount
            lblPaletteV.Caption = aBPP & " bpp"
            lblEntriesV.Caption = m_oGIF.GlobalPaletteEntries & " entries"
            
        Case ucCanvas.DIB.hDIB = 0
            '-- Reset AVI props.
            lblScreenV.Caption = ""
            lblFramesV.Caption = ""
            lblPaletteV.Caption = ""
            lblEntriesV.Caption = ""
    End Select
End Sub

'//

Private Sub iPalette_Paint()
    pvPaintPalette
End Sub

Private Sub pvPaintPalette()
  
  Dim i As Long, j As Long
  Dim lIdx As Long
  Dim lClr As Long
    
    '-- Show the 256 entries
    For i = 0 To 90 Step 6
        For j = 0 To 90 Step 6
            With m_oGIF
                If (lIdx < .GlobalPaletteEntries) Then
                    lClr = .GlobalPaletteRGBEntry(lIdx)
                  Else
                    lClr = -1
                End If
                mMisc.DrawRectangle iPalette.hDC, j, i, j + 6, i + 6, lClr
                lIdx = lIdx + 1
            End With
        Next j
    Next i
End Sub

