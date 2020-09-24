VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AVI2BMP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AVI video joiner"
   ClientHeight    =   2865
   ClientLeft      =   1785
   ClientTop       =   435
   ClientWidth     =   4020
   Icon            =   "AVI2BMP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   191
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   -75
      Top             =   2460
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3915
      Pattern         =   "*.avi"
      TabIndex        =   13
      Top             =   2535
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2445
      Left            =   -15
      TabIndex        =   0
      Top             =   -75
      Width           =   4095
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   0
         MousePointer    =   99  'Custom
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   40
         Width           =   4020
      End
      Begin VB.CommandButton Command9 
         Caption         =   "redraw"
         Height          =   330
         Left            =   2475
         TabIndex        =   6
         Top             =   2745
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdOpenAVIFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Extract BMPs"
         Height          =   255
         Left            =   1365
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2100
         Width           =   1470
      End
      Begin VB.TextBox txtStatus 
         Enabled         =   0   'False
         Height          =   330
         Left            =   135
         TabIndex        =   1
         Text            =   "No AVI File Selected"
         Top             =   2670
         Visible         =   0   'False
         Width           =   3840
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dir"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   3225
         TabIndex        =   12
         Top             =   2265
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AVIs"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   45
         TabIndex        =   11
         Top             =   2265
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   10
         Top             =   1755
         Width           =   660
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   225
         Left            =   60
         Top             =   2025
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3285
         TabIndex        =   9
         Top             =   1755
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         Height          =   225
         Left            =   3225
         Top             =   2025
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1245
      Left            =   -30
      ScaleHeight     =   1185
      ScaleWidth      =   1380
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   3
      Top             =   33000
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drag and Drop AVIs to re-order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   2670
      Width           =   4080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drag and Drop AVI-files to above box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   270
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   2430
      Width           =   4080
   End
   Begin VB.Image tmpImg 
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   1440
   End
End
Attribute VB_Name = "AVI2BMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If Win16 Then
    Private Declare Function SendMessage& Lib "User" (ByVal hWnd%, ByVal _
            wMsg%, ByVal wParam%, lParam As Any)
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
            (ByVal hWnd As Long, ByVal wMsg As Long, _
            ByVal wParam As Long, lParam As Long) As Long
#End If
Private DragIndex As Integer
Private DragItem As String
Private Dragging As Boolean
Dim ExtractBMP As Boolean

Private Sub Command9_Click()
On Error Resume Next
Dim NewPict As String
tmpImg.Refresh
Picture1.Picture = LoadPicture("")

Dim xImg, yImg As Single
Dim xPic, yPic As Single
xImg = tmpImg.Width
yImg = tmpImg.Height
xPic = Picture1.Width
yPic = Picture1.Height

Dim xRatio, yRatio As Single
xRatio = xImg / xPic
yRatio = yImg / yPic

If xRatio >= yRatio Then
Picture1.PaintPicture tmpImg.Picture, 0, 0, (tmpImg.Width * 15.5 / xRatio), (tmpImg.Height * 15.5 / xRatio)
Else
Picture1.PaintPicture tmpImg.Picture, 0, 0, (tmpImg.Width * 15.5 / yRatio), (tmpImg.Height * 15.5 / yRatio)
End If
Framed = Framed + 1
NewPict = App.Path & "\Images1\" & Format(Now, "ddmmyyhhmmss") & Framed & ".bmp"
SavePicture Picture1.Image, NewPict
'BMP2AVI.lstDIBList.AddItem NewPict
'BMP2AVI.Combo1.AddItem NewPict
End Sub


Private Sub Form_Activate()
On Error Resume Next
List1.MouseIcon = LoadResPicture(101, vbResIcon)
frmMain.Caption = "Extract BMPs from AVIs"
End Sub

Private Sub Form_Initialize()
Hidder = True
Initt = False
End Sub

Private Sub Form_Load()
On Error Resume Next
ExtractBMP = False
    SetTopMostWindow Me.hWnd, True
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
Kill App.Path & "\Images\*.bmp"
Kill App.Path & "\Images1\*.bmp"
'Load BMP2AVI
    Call AVIFileInit   '// opens AVIFile library
    'cmdOpenAVIFile_Click
List1.MouseIcon = LoadResPicture(101, vbResIcon)
'List1.AddItem App.Path & "\missused_by_lesbian_part02.avi"
'List1.AddItem App.Path & "\missused_by_lesbian_part03.avi"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo here
    Call AVIFileExit   '// releases AVIFile library
    Set AVI2BMP = Nothing
    If ExtractBMP = False Then Exit Sub
    Hidder = False
    Load BMP2AVI
    BMP2AVI.Show
    'If Split = False Then
        'BMP2AVI.cmdWriteAVI_Click
    'End If
'Set AVI2BMP = Nothing
Exit Sub
here:
Set AVI2BMP = Nothing
Unload frmMain
Set frmMain = Nothing
Unload BMP2AVI
Set BMP2AVI = Nothing
End Sub

Public Sub cmdOpenAVIFile_Click()
    Dim res As Long         'result code
    Dim ofd As cFileDlg     'OpenFileDialog class
    Dim szFile As String    'filename
    Dim szzFile As String    'filename
    Dim sFile As String    'filename
    Dim pAVIFile As Long    'pointer to AVI file interface (PAVIFILE handle)
    Dim pAVIStream As Long  'pointer to AVI stream interface (PAVISTREAM handle)
    Dim numFrames As Long   'number of frames in video stream
    Dim firstFrame As Long  'position of the first video frame
    Dim fileInfo As AVI_FILE_INFO       'file info struct
    Dim streamInfo As AVI_STREAM_INFO   'stream info struct
    Dim dib As cDIB
    Dim pGetFrameObj As Long    'pointer to GetFrame interface
    Dim pDIB As Long            'pointer to packed DIB in memory
    Dim bih As BITMAPINFOHEADER 'infoheader to pass to GetFrame functions
    Dim i As Long, ii As Long, jj As Long, kk As Integer
    Dim intSave As Integer
    SetTopMostWindow Me.hWnd, False
On Error Resume Next
If AVIInput = False Then
    intSave = MsgBox("You will not be able to Cancel the operation once it starts!" & vbCrLf _
    & "  Do you want to proceed?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intSave
      Case vbYes
      Case vbNo
        Exit Sub
      Case vbCancel
        Exit Sub
    End Select
End If
    ExtractBMP = True
jj = 0: kk = 1
Frame1.Visible = False
For ii = 0 To List1.ListCount - 1
    szFile = List1.List(ii)      'App.Path & "\MyAVI.avi"
    'res = ofd.VBGetOpenFileNamePreview(szFile)
    'If res = False Then GoTo ErrorOut
    'Open the AVI File and get a file interface pointer (PAVIFILE)
    res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo ErrorOut
 
    'Get the first available video stream (PAVISTREAM)
    res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If res <> AVIERR_OK Then GoTo ErrorOut
    
    'get the starting position of the stream (some streams may not start simultaneously)
    firstFrame = AVIStreamStart(pAVIStream)
    If firstFrame = -1 Then GoTo ErrorOut 'this function returns -1 on error
    
    'get the length of video stream in frames
    numFrames = AVIStreamLength(pAVIStream)
    If numFrames = -1 Then GoTo ErrorOut ' this function returns -1 on error
    
'    MsgBox "PAVISTREAM handle is " & pAVIStream & vbCrLf & _
'            "Video stream length - " & numFrames & vbCrLf & _
'            "Stream starts on frame #" & firstFrame & vbCrLf & _
'            "File and Stream info will be written to Immediate Window (from IDE - Ctrl+G to view)", vbInformation, App.title
'
    'get file info struct (UDT)
    res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut
    
'    'print file info to Debug Window
'    Call DebugPrintAVIFileInfo(fileInfo)
    
    'get stream info struct (UDT)
    res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut
    
'    'print stream info to Debug Window
    'Call DebugPrintAVIStreamInfo(streamInfo)
    'set bih attributes which we want GetFrame functions to return
    If Picture1.Height < streamInfo.rcFrame.bottom Then
        Picture1.Height = streamInfo.rcFrame.bottom
        Picture2.Height = streamInfo.rcFrame.bottom
    End If
    If Picture1.Width < streamInfo.rcFrame.Right Then
        Picture1.Width = streamInfo.rcFrame.Right
        Picture2.Width = streamInfo.rcFrame.Right
    End If
    Me.Height = Picture1.Height * 15.5
    Me.Width = Picture1.Width * 15.5
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Picture1.Visible = True
    AVI2BMP.Enabled = False

    Set dib = Nothing
    Call AVIFileExit
    Call AVIFileInit   '// opens AVIFile library

Next
'Exit Sub
For ii = 0 To List1.ListCount - 1
    szFile = List1.List(ii)      'App.Path & "\MyAVI.avi"
    res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If res <> AVIERR_OK Then GoTo ErrorOut
    res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If res <> AVIERR_OK Then GoTo ErrorOut
    firstFrame = AVIStreamStart(pAVIStream)
    If firstFrame = -1 Then GoTo ErrorOut 'this function returns -1 on error
    numFrames = AVIStreamLength(pAVIStream)
    If numFrames = -1 Then GoTo ErrorOut ' this function returns -1 on error
    res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut
    res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
    If res <> AVIERR_OK Then GoTo ErrorOut

    With bih
        .biBitCount = 24
        .biClrImportant = 0
        .biClrUsed = 0
        .biCompression = BI_RGB
        .biHeight = streamInfo.rcFrame.bottom - streamInfo.rcFrame.Top
        .biPlanes = 1
        .biSize = 40
        .biWidth = streamInfo.rcFrame.Right - streamInfo.rcFrame.Left
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight 'calculate total size of RGBQUAD scanlines (DWORD aligned)
    End With
    
    pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, bih) 'force function to return 24bit DIBS
    If pGetFrameObj = 0 Then
        MsgBox "No suitable decompressor found for this video stream!", vbInformation, App.Title
        'Exit Sub
        ExtractBMP = False
        GoTo ErrorOut
    End If
    
    'create a DIB class to load the frames into
    Set dib = New cDIB

    For i = firstFrame To (numFrames - 1) + firstFrame
        jj = jj + i
        sFile = "sFile.bmp"
        pDIB = AVIStreamGetFrame(pGetFrameObj, i)  'returns "packed DIB"
        If dib.CreateFromPackedDIBPointer(pDIB) And kk = Val(frmMain.Text1.Text) Then
            Call dib.WriteToFile(App.Path & "\Images\" & sFile)
            frmMain.Caption = "Bitmap " & i + 1 & " of " & numFrames & " " & List1.List(ii)
            'txtStatus.Refresh
            kk = 1
            tmpImg.Picture = LoadPicture(App.Path & "\Images\" & sFile)
            Command9_Click
        Else
            kk = kk + 1
        End If
    Next
    Set dib = Nothing
    Call AVIFileExit
    Call AVIFileInit   '// opens AVIFile library
Next
Unload Me
Exit Sub

ErrorOut:
    If pGetFrameObj <> 0 Then
        Call AVIStreamGetFrameClose(pGetFrameObj) '//deallocates the GetFrame resources and interface
    End If
    If pAVIStream <> 0 Then
        Call AVIStreamRelease(pAVIStream) '//closes video stream
    End If
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile) '// closes the file
    End If
    
    If (res <> AVIERR_OK) Then 'if there was an error then show feedback to user
        'MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.title
    End If
    Unload Me
    'Resume Next
End Sub

Private Sub tmpPic_Click()

End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo here
Dim ReturnValuePath As String, i As Long
ReturnValuePath = BrowseForFolder(Me.hWnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)
File1.Path = ReturnValuePath
'MsgBox File1.ListCount - 1
If File1.ListCount - 1 >= 0 Then
    For i = 0 To File1.ListCount - 1
        List1.AddItem File1.List(i)
    Next
End If
Exit Sub
here:
File1.Path = App.Path

End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    cdlOpen.Filter = "AVI|*.avi"
    cdlOpen.ShowOpen
    If cdlOpen.filename <> "" Then
      List1.AddItem cdlOpen.filename
    End If


End Sub

Private Sub List1_DblClick()
On Error Resume Next
Dim intSave As String
    SetTopMostWindow Me.hWnd, False
    intSave = MsgBox("Do you want to delete " & vbCrLf & _
                List1.List(List1.ListIndex) & " ?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intSave
      Case vbYes
        List1.RemoveItem List1.ListIndex
        List1.Refresh
        SetTopMostWindow Me.hWnd, True
      Case vbNo
        SetTopMostWindow Me.hWnd, True
        Exit Sub
      Case vbCancel
        SetTopMostWindow Me.hWnd, True
        Exit Sub
    End Select

End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim XY As Long, c As String
        
        For XY = 1 To Data.Files.Count
            c = Data.Files(XY)
            If LCase(Right(c, 3)) = "avi" Then
                List1.AddItem Data.Files(XY), XY - 1
            End If
        Next XY
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo err:
    
    If Button = vbLeftButton Then
        If List1.ListCount > 1 Then
            Dragging = True
            DragIndex = ListRowCalc(List1, Y)
            DragItem = List1.List(DragIndex)
            List1.ListIndex = DragIndex
            List1.MouseIcon = LoadResPicture(101, vbResCursor)
        End If
    End If
    
    Exit Sub

err:

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo err:
    Dim curIndex As Integer
    Dim i As Integer
    If Dragging Then
        curIndex = ListRowCalc(List1, Y)
        If curIndex <> DragIndex And curIndex >= 0 Then
            If curIndex > DragIndex Then
                For i = DragIndex To curIndex - 1
                    List1.List(i) = List1.List(i + 1)
                Next i
            Else
                For i = DragIndex To curIndex + 1 Step -1
                    List1.List(i) = List1.List(i - 1)
                Next i
            End If
            List1.List(curIndex) = DragItem
            List1.ListIndex = curIndex
            DragIndex = curIndex
        End If
    End If
    Exit Sub
err:

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo err:
    If Button = vbLeftButton Then
        Dragging = False
    List1.MouseIcon = LoadResPicture(101, vbResIcon)
    End If
    Exit Sub
err:

End Sub

Private Function ListRowCalc(lstTemp As Control, ByVal Y As Single) As Integer
On Error Resume Next
    
    #If Win16 Then
        Const WM_USER = &H400
        Const LB_GETITEMHEIGHT = (WM_USER + 34)
    #Else
        Const LB_GETITEMHEIGHT = &H1A1
    ' Determines the height of each item in ListBox control in pixels '
    #End If
    
    Dim ItemHeight As Integer
    
    ItemHeight = SendMessage(lstTemp.hWnd, LB_GETITEMHEIGHT, 0, 0)
    
    ListRowCalc = Min(((Y / Screen.TwipsPerPixelY) \ ItemHeight) + _
            lstTemp.TopIndex, lstTemp.ListCount - 1)
    Exit Function
    
End Function
Private Function Min(X As Integer, Y As Integer) As Integer
On Error Resume Next
    If X > Y Then Min = Y Else Min = X
    Exit Function
End Function




