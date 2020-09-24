VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form GraphicalDll 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Special Effects - (none)"
   ClientHeight    =   4320
   ClientLeft      =   465
   ClientTop       =   825
   ClientWidth     =   9960
   Icon            =   "GraphicalDll.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Process Frames"
      Enabled         =   0   'False
      Height          =   270
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   315
      Width           =   1410
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   195
      Left            =   60
      Max             =   255
      Min             =   -255
      TabIndex        =   2
      Top             =   30
      Width           =   9855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12645
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   30
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   645
      Width           =   4860
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3645
      Left            =   4875
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   332
      TabIndex        =   1
      Top             =   660
      Width           =   5040
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   3615
      Left            =   4845
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   9
      Top             =   645
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   2385
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apply Effect"
      Height          =   270
      Left            =   8625
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   300
      Width           =   1320
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Effect"
      Height          =   270
      Left            =   8610
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   300
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   2670
      TabIndex        =   10
      Top             =   3675
      Visible         =   0   'False
      Width           =   1395
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8400
         TabIndex        =   21
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7050
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "255"
         Top             =   75
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7050
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "255"
         Top             =   315
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7050
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "255"
         Top             =   555
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
         Height          =   255
         Left            =   8400
         TabIndex        =   17
         Top             =   60
         Width           =   855
      End
      Begin VB.CommandButton cmdSelectTool 
         DownPicture     =   "GraphicalDll.frx":08CA
         Height          =   495
         Left            =   7815
         Picture         =   "GraphicalDll.frx":0C9B
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   90
         Width           =   495
      End
      Begin VB.ListBox List1 
         Height          =   645
         Left            =   2250
         TabIndex        =   14
         Top             =   195
         Width           =   3750
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add==>"
         Height          =   225
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   780
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Replay"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1455
         TabIndex        =   12
         Top             =   660
         Width           =   675
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   195
         Left            =   1455
         TabIndex        =   11
         Top             =   465
         Width           =   750
      End
      Begin VB.CommandButton cmdCopyImage 
         Caption         =   "Copy Image"
         Height          =   315
         Left            =   75
         TabIndex        =   15
         Top             =   45
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Color"
         Height          =   195
         Left            =   9285
         TabIndex        =   25
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R:"
         Height          =   195
         Left            =   6855
         TabIndex        =   24
         Top             =   75
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   195
         Left            =   6855
         TabIndex        =   23
         Top             =   315
         Width           =   165
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "B:"
         Height          =   195
         Left            =   6855
         TabIndex        =   22
         Top             =   555
         Width           =   150
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You may drag the pictures around"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   3450
      TabIndex        =   28
      Top             =   390
      Width           =   3015
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9060
      TabIndex        =   8
      Top             =   15
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Image with effect applied"
      Height          =   195
      Left            =   6810
      TabIndex        =   7
      Top             =   1035
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Original"
      Height          =   195
      Left            =   1410
      TabIndex        =   6
      Top             =   1215
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   9960
      TabIndex        =   3
      Top             =   4545
      Width           =   1215
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save As..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnutraco 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSair 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "Effects"
      Begin VB.Menu mnu1 
         Caption         =   "Color Adjustment"
         Begin VB.Menu mnuBrilho 
            Caption         =   "Brightness"
         End
         Begin VB.Menu mnuContraste 
            Caption         =   "Contrast"
         End
         Begin VB.Menu mnuNitidez 
            Caption         =   "Sharpening"
         End
         Begin VB.Menu mnuGamma 
            Caption         =   "Gamma Adjust"
         End
         Begin VB.Menu mnuReduceColors 
            Caption         =   "Reduce Colors"
         End
         Begin VB.Menu mnuEightColors 
            Caption         =   "Reduce to 8 colors"
         End
         Begin VB.Menu mnuShift 
            Caption         =   "Shift Effect"
         End
         Begin VB.Menu mnuSaturation 
            Caption         =   "Saturation"
         End
         Begin VB.Menu mnuHue 
            Caption         =   "Hue Adjust"
         End
         Begin VB.Menu mnuColBalance 
            Caption         =   "Color Balance"
         End
         Begin VB.Menu mnuWebColors 
            Caption         =   "WebColors Mode"
         End
         Begin VB.Menu mnuMediumTones 
            Caption         =   "Medium Tones"
         End
         Begin VB.Menu mnuStretchHisto 
            Caption         =   "Stretch Histogram"
         End
      End
      Begin VB.Menu mnu2 
         Caption         =   "Blur"
         Begin VB.Menu mnuAlias 
            Caption         =   "AntiAlias"
         End
         Begin VB.Menu mnuBlur 
            Caption         =   "Blur"
         End
         Begin VB.Menu mnuSmartBlur 
            Caption         =   "SmartBlur"
         End
         Begin VB.Menu mnuMoreBlur 
            Caption         =   "More Blur"
         End
         Begin VB.Menu mnuSoftnerBlur 
            Caption         =   "Softner Blur"
         End
         Begin VB.Menu mnuMotionBlur 
            Caption         =   "Motion Blur"
         End
         Begin VB.Menu mnuFarBlur 
            Caption         =   "Far Blur"
         End
         Begin VB.Menu mnuRadialBlur 
            Caption         =   "Radial Blur"
         End
         Begin VB.Menu mnuZoomBlur 
            Caption         =   "Zoom Blur"
         End
         Begin VB.Menu mnuUnsharpMask 
            Caption         =   "Unsharp Mask"
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "Tones"
         Begin VB.Menu mnuGrayScale 
            Caption         =   "Gray Tones"
         End
         Begin VB.Menu mnuSepia 
            Caption         =   "Sepia Effect"
         End
         Begin VB.Menu mnuAmbient 
            Caption         =   "Ambient Light"
         End
         Begin VB.Menu mnuTone 
            Caption         =   "Tone Adjust"
         End
      End
      Begin VB.Menu mnu4 
         Caption         =   "Distortion"
         Begin VB.Menu mnuMosaico 
            Caption         =   "Mosaic"
         End
         Begin VB.Menu mnuDiffuse 
            Caption         =   "Diffuse"
         End
         Begin VB.Menu mnuRock 
            Caption         =   "Rock Effect"
         End
         Begin VB.Menu mnuNoise 
            Caption         =   "Noise"
         End
         Begin VB.Menu mnuMelt 
            Caption         =   "Melt"
         End
         Begin VB.Menu mnuFishEye 
            Caption         =   "Fish Eye"
         End
         Begin VB.Menu mnuFishEyeEx 
            Caption         =   "Fish Eye Ex"
         End
         Begin VB.Menu mnuTwirl 
            Caption         =   "Twirl"
         End
         Begin VB.Menu mnuTwirlEx 
            Caption         =   "TwirlEx"
         End
         Begin VB.Menu mnuSwirl 
            Caption         =   "Swirl"
         End
         Begin VB.Menu mnu3D 
            Caption         =   "Make 3D"
         End
         Begin VB.Menu mnu4Corners 
            Caption         =   "Four Corners"
         End
         Begin VB.Menu mnuCaricature 
            Caption         =   "Caricature"
         End
         Begin VB.Menu mnuRoll 
            Caption         =   "Enroll"
         End
         Begin VB.Menu mnuPolar 
            Caption         =   "Polar Coordinates"
         End
         Begin VB.Menu mnuCilindrical 
            Caption         =   "Cilindrical"
         End
      End
      Begin VB.Menu mnuW 
         Caption         =   "Waves"
         Begin VB.Menu mnuWave 
            Caption         =   "Waves"
         End
         Begin VB.Menu mnuBlockWaves 
            Caption         =   "Block Waves"
         End
         Begin VB.Menu mnuCircularWaves 
            Caption         =   "Circular Waves"
         End
         Begin VB.Menu mnuCircularWavesEx 
            Caption         =   "Circular Waves Enhanced"
         End
      End
      Begin VB.Menu mnu5 
         Caption         =   "Borders"
         Begin VB.Menu mnuBack 
            Caption         =   "Backdrop Removal"
         End
         Begin VB.Menu mnuEmbEng 
            Caption         =   "Emboss / Engrave"
         End
         Begin VB.Menu mnuNeon 
            Caption         =   "Neon"
         End
         Begin VB.Menu mnuBorders 
            Caption         =   "Detect Borders"
         End
         Begin VB.Menu mnuFindEdges 
            Caption         =   "Find Edges"
         End
         Begin VB.Menu mnuNotePaper 
            Caption         =   "Note Paper"
         End
      End
      Begin VB.Menu mnuBlend 
         Caption         =   "Blend Modes"
         Begin VB.Menu mnuAlphaBlend 
            Caption         =   "AlphaBlend"
         End
         Begin VB.Menu mnuAlpha3D 
            Caption         =   "AlphaBlend 3D Text"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBlendModes 
            Caption         =   "Blend Modes"
         End
         Begin VB.Menu mnuGlassBlendMode 
            Caption         =   "Glass Blend Mode"
         End
      End
      Begin VB.Menu mnuMetal 
         Caption         =   "Metallic Effects"
         Begin VB.Menu mnuMetallic 
            Caption         =   "Metallic"
         End
         Begin VB.Menu mnuGold 
            Caption         =   "Gold"
         End
         Begin VB.Menu mnuIce 
            Caption         =   "Ice"
         End
      End
      Begin VB.Menu mnu6 
         Caption         =   "Other Effects"
         Begin VB.Menu mnuInvertion 
            Caption         =   "Invertion Adjust"
         End
         Begin VB.Menu mnuMono 
            Caption         =   "Monochrome"
         End
         Begin VB.Menu mnuBackN 
            Caption         =   "Replace Color"
         End
         Begin VB.Menu mnuAscii 
            Caption         =   "Ascii Effect"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRandomPoints 
            Caption         =   "Random Points"
         End
         Begin VB.Menu mnuSol 
            Caption         =   "Solarize"
         End
         Begin VB.Menu mnuCanvas 
            Caption         =   "Canvas Adjust"
         End
         Begin VB.Menu mnuRelief 
            Caption         =   "Relief"
         End
         Begin VB.Menu mnuTile 
            Caption         =   "Tile Effect"
         End
         Begin VB.Menu mnuFragment 
            Caption         =   "Fragment"
         End
         Begin VB.Menu mnuFog 
            Caption         =   "Fog Effect"
         End
         Begin VB.Menu mnuOilPaint 
            Caption         =   "Oil Paint"
         End
         Begin VB.Menu mnuFrostGlass 
            Caption         =   "Frost Glass"
         End
         Begin VB.Menu mnuRainDrop 
            Caption         =   "Rain Drop"
         End
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "Image"
      Begin VB.Menu mnuFlipH 
         Caption         =   "Flip Horizontal"
      End
      Begin VB.Menu mnuFlipV 
         Caption         =   "Flip Vertical"
      End
      Begin VB.Menu mnuFlipB 
         Caption         =   "Flip Both"
      End
   End
End
Attribute VB_Name = "GraphicalDll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
    ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 Const HTCAPTION = 2
 Const WM_NCLBUTTONDOWN = &HA1

Private ColorPickTool As Boolean, FirstOne As Boolean
Private LastPath As String
Private Resp As Long
Dim i As Long, BMPIndex As Long
Dim a As Integer, B As String, c As String, d As Integer, _
            e As Integer, f As Integer, G As Boolean, TempStuff As String, TempStuff1 As String
    

Private Sub cmdCopyImage_Click()
    GPX_BitBlt Picture1.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, vbSrcCopy, Resp
    Picture1.Refresh
    CopyImage = True
End Sub

Private Sub cmdSelectTool_Click()
    ColorPickTool = Not ColorPickTool
    If (ColorPickTool) Then
        cmdSelectTool.Picture = LoadPicture(App.Path & "\Cursores\Arrow.gif")
    Else
        cmdSelectTool.Picture = LoadPicture(App.Path & "\Cursores\cursor.jpg")
    End If
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Text2.BackColor = RGB(Text3.Text, Text4.Text, Text5.Text)
End Sub

Private Sub Command2_Click()
    Dim Ticks As Long
    Dim Color As Long
    Command2.Visible = False
    Command7.Visible = True
    'On Error Resume Next
    Color = Text2.BackColor
    If (Effect <> 14) Then
        Picture2.Cls
    End If
    Ticks = GetTickCount
    Select Case Effect
        Case 4
            GPX_AntiAlias Picture2.hdc, Picture1.hdc, 50, Resp
        Case 9
            GPX_ReduceTo8Colors Picture2.hdc, Picture1.hdc, Resp
        Case 11
            GPX_Sepia Picture2.hdc, Picture1.hdc, Resp
        Case 14
            Dim sBuffer As String
            GPX_AllocBufferSize Picture2.hdc, Resp
            sBuffer = Space(Resp)
            GPX_AsciiMorph Picture2.hdc, sBuffer, Resp
            Text1.Text = sBuffer
        Case 18
            GPX_Diffuse Picture2.hdc, Picture1.hdc, Resp
        Case 23
            GPX_Solarize Picture2.hdc, Picture1.hdc, False, Resp
        Case 25
            GPX_Melt Picture2.hdc, Picture1.hdc, Resp
        Case 26
            GPX_FishEye Picture2.hdc, Picture1.hdc, Resp
        Case 33
            GPX_Blur Picture2.hdc, Picture1.hdc, Resp
        Case 34
            GPX_Relief Picture2.hdc, Picture1.hdc, Resp
        Case 39
            GPX_Make3DEffect Picture2.hdc, Picture1.hdc, 6, Resp
        Case 40
            GPX_FourCorners Picture2.hdc, Picture1.hdc, Resp
        Case 41
            GPX_Caricature Picture2.hdc, Picture1.hdc, Resp
        Case 43
            GPX_Roll Picture2.hdc, Picture1.hdc, Resp
        Case 44
            GPX_SmartBlur Picture2.hdc, Picture1.hdc, 20, Resp
        Case 46
            GPX_SoftnerBlur Picture2.hdc, Picture1.hdc, Resp
        Case 53
            GPX_WebColors Picture2.hdc, Picture1.hdc, Resp
        Case 58
            GPX_PolarCoordinates Picture2.hdc, Picture1.hdc, 0, Resp
        Case 60
            GPX_FrostGlass Picture2.hdc, Picture1.hdc, 3, Resp
        Case 63
            GPX_RainDrops Picture2.hdc, Picture1.hdc, 40, 50, 40, Resp
        Case 67
            GPX_StretchHistogram Picture2.hdc, Picture1.hdc, HST_COLOR, 1, Resp
    End Select
    Ticks = GetTickCount - Ticks
    lblTime.Caption = Ticks & " ms"
    Picture2.Refresh
    a = Processor
    B = "Command2"
    c = "Hscroll1"
    d = ScrollMax
    e = ScrollMin
    f = ScrollValue
    G = ApplyEffect
    Command6.Enabled = True
Exit Sub

Open App.Path & "\TempFile" For Output As #1
    Print #1, Processor
    Print #1, "Command2"
    Print #1, "Hscroll1"
    Print #1, ScrollMax
    Print #1, ScrollMin
    Print #1, ScrollValue
    Print #1, ApplyEffect
Close #1
        Open App.Path & "\Tempfile" For Input As #1
                Line Input #1, TempStuff
                a = Val(TempStuff)
                Line Input #1, TempStuff
                B = TempStuff
                Line Input #1, TempStuff
                c = TempStuff
                Line Input #1, TempStuff
                d = Val(TempStuff)
                Line Input #1, TempStuff
                e = Val(TempStuff)
                Line Input #1, TempStuff
                f = Val(TempStuff)
                Line Input #1, TempStuff
                TempStuff1 = TempStuff
        Close #1
        If TempStuff1 = "True" Then
            G = True
        Else
            G = False
        End If
    Command6.Enabled = True

End Sub

Private Sub Command3_Click()
List1.AddItem Processor & "--" & "Command2" & "--" & "Hscroll1" _
              & "--" & ScrollMax & "--" & ScrollMin & "--" & ScrollValue & "--" & _
              ApplyEffect
Open App.Path & "\TempFile" For Output As #1
    Print #1, Processor
    Print #1, "Command2"
    Print #1, "Hscroll1"
    Print #1, ScrollMax
    Print #1, ScrollMin
    Print #1, ScrollValue
    Print #1, ApplyEffect
Close #1
End Sub

Private Sub Command4_Click()
Dim a As Integer, B As String, c As String, d As Integer, _
        e As Integer, f As Integer, G As Boolean, TempStuff As String, TempStuff1 As String
For i = 0 To List1.ListCount - 1
    Open App.Path & "\Tempfile" For Output As #1
        Print #1, Replace(List1.List(i), "--", vbCrLf)
    Close #1
    Open App.Path & "\Tempfile" For Input As #1
            Line Input #1, TempStuff
            a = Val(TempStuff)
            Line Input #1, TempStuff
            B = TempStuff
            Line Input #1, TempStuff
            c = TempStuff
            Line Input #1, TempStuff
            d = Val(TempStuff)
            Line Input #1, TempStuff
            e = Val(TempStuff)
            Line Input #1, TempStuff
            f = Val(TempStuff)
            Line Input #1, TempStuff
            TempStuff1 = TempStuff
    Close #1
    If TempStuff1 = "True" Then
        G = True
    Else
        G = False
    End If
    
    ChangeControls a, Command2, HScroll1, d, e, f
    If G = True Then
        Command2_Click
    Else
        HScroll1_Change
    End If
    cmdCopyImage_Click
Next

End Sub

Private Sub Command6_Click()
On Error Resume Next
Command6.Enabled = False
Command2.Visible = True
Command2.Enabled = False
BMPIndex = 0
'DoEvents
Do While BMPIndex <= BMP2AVI.lstDIBList.ListCount - 1
    If ApplyFlag = True Then
        If BMP2AVI.lstDIBList.Selected(BMPIndex) = True Then
            Picture1.Picture = LoadPicture("")
            Picture1.Refresh
            Picture1.Picture = LoadPicture(BMP2AVI.lstDIBList.List(BMPIndex))
        End If
    Else
        Picture1.Picture = LoadPicture("")
        Picture1.Refresh
        Picture1.Picture = LoadPicture(BMP2AVI.lstDIBList.List(BMPIndex))
    End If
Select Case a
    Case 11, 18, 9, 53, 67, 4, 33, 44, 46, 25, 26, 39, 40, 41, 43, 58, 23, 34, 60, 63
        ChangeControls a, Command2, HScroll1
    Case 30 'canvas
        ChangeControls 30, Command2, HScroll1, 0, Picture1.ScaleWidth, 0
    Case 38
        GPX_AlphaBlend Picture2.hdc, Picture3.hdc, Picture1.hdc, ScrollValue, Resp
        'ChangeControls 38, Command2, HScroll1, 0, 255, 0
        'HScroll1_Change
        cmdCopyImage_Click
        GoTo Skipper
    Case 100
        GPX_Flip Picture1.hdc, Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, 1, 0, Resp
        Picture1.Refresh
        GoTo Skipper
    Case 101
        GPX_Flip Picture1.hdc, Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 1, Resp
        Picture1.Refresh
        GoTo Skipper
    Case 102
        GPX_Flip Picture1.hdc, Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, 1, 1, Resp
        Picture1.Refresh
        GoTo Skipper
    
    Case Else   '6 item stuff
        ChangeControls a, Command2, HScroll1, d, e, f
        
End Select
    If G = True Then
        Command2_Click
    Else
        HScroll1_Change
    End If
    cmdCopyImage_Click
Skipper:
    If ApplyFlag = True Then
        If BMP2AVI.lstDIBList.Selected(BMPIndex) = True Then
            SavePicture Picture1.Image, BMP2AVI.lstDIBList.List(BMPIndex)
            BMPIndex = BMPIndex + 1
        Else
            BMPIndex = BMPIndex + 1
        End If
    Else
        SavePicture Picture1.Image, BMP2AVI.lstDIBList.List(BMPIndex)
        BMPIndex = BMPIndex + 1
    End If
Loop
'Unload Me
End Sub

Private Sub Command7_Click()
    Command7.Visible = False
    Command6.Enabled = False
    Command2.Visible = True
    Command2.Enabled = False
    Picture2.Picture = LoadPicture("")
End Sub


Private Sub Form_Activate()
SetTopMostWindow Me.hWnd, True
On Error Resume Next
BMPIndex = 0    'BMP2AVI.lstDIBList.ListIndex
Picture1.Picture = LoadPicture(BMP2AVI.lstDIBList.List(0))
Picture2.Height = Picture1.Height
Picture2.Width = Picture1.Width
Picture2.Top = Picture1.Top
Picture2.Left = Picture1.Left + Picture1.Width
Picture3.Height = Picture1.Height
Picture3.Width = Picture1.Width
Exit Sub



If ApplyFlag = True Then
        Open App.Path & "\Tempfile" For Input As #1
                Line Input #1, TempStuff
                a = Val(TempStuff)
                Line Input #1, TempStuff
                B = TempStuff
                Line Input #1, TempStuff
                c = TempStuff
                Line Input #1, TempStuff
                d = Val(TempStuff)
                Line Input #1, TempStuff
                e = Val(TempStuff)
                Line Input #1, TempStuff
                f = Val(TempStuff)
                Line Input #1, TempStuff
                TempStuff1 = TempStuff
        Close #1
        If TempStuff1 = "True" Then
            G = True
        Else
            G = False
        End If
        
        ChangeControls a, Command2, HScroll1, d, e, f
        If G = True Then
            Command2_Click
        Else
            HScroll1_Change
        End If
        cmdCopyImage_Click
End If

End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Effect = 0
    LastPath = App.Path
    ColorPickTool = False
    HScroll1.Enabled = False
    Command2.Enabled = False
    Text2.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Form_Resize()
    'PleaseDontResize Me, 11805, 8835
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set GraphicalDLL = Nothing
    BMP2AVI.FXImage(0).Picture = LoadPicture(BMP2AVI.lstDIBList.List(0))
    BMP2AVI.FXImage(0).Refresh
    BMP2AVI.FXImage(1).Picture = LoadPicture(BMP2AVI.lstDIBList.List(1))
    BMP2AVI.FXImage(1).Refresh
    BMP2AVI.FXImage(2).Picture = LoadPicture(BMP2AVI.lstDIBList.List(2))
    BMP2AVI.FXImage(2).Refresh
    BMP2AVI.FXImage(3).Picture = LoadPicture(BMP2AVI.lstDIBList.List(3))
    BMP2AVI.FXImage(3).Refresh
    BMP2AVI.FXImage(4).Picture = LoadPicture(BMP2AVI.lstDIBList.List(4))
    BMP2AVI.FXImage(4).Refresh
    BMP2AVI.FXImage(5).Picture = LoadPicture(BMP2AVI.lstDIBList.List(5))
    BMP2AVI.FXImage(5).Refresh
SetTopMostWindow BMP2AVI.hWnd, True

End Sub

Private Sub HScroll1_Change()
    Dim Color As Long
    Dim Ticks As Long
    
    On Error Resume Next
    Color = Text2.BackColor
    Label1.Caption = HScroll1.Value
    Ticks = GetTickCount
    Select Case Effect
        Case 1
            GPX_Brightness Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 2
            GPX_Contrast Picture2.hdc, Picture1.hdc, HScroll1.Value / 100, HScroll1.Value / 100, HScroll1.Value / 100, Resp
        Case 3
            GPX_Sharpening Picture2.hdc, Picture1.hdc, HScroll1.Value / 100, Resp
        Case 5
            GPX_Gamma Picture2.hdc, Picture1.hdc, HScroll1.Value / 25, Resp
        Case 6
            GPX_GrayScale Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 7
            GPX_Invert Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 8
            GPX_ReduceColors Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 10
            GPX_Stamp Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 12
            GPX_Mosaic Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 13
            GPX_BackDropRemoval Picture2.hdc, Picture1.hdc, Color, Not Color, HScroll1.Value, Resp
        Case 15
            GPX_AmbientLight Picture2.hdc, Picture1.hdc, Color, HScroll1.Value, Resp
        Case 16
            GPX_Tone Picture2.hdc, Picture1.hdc, Color, HScroll1.Value, Resp
        Case 17
            GPX_BackDropRemovalEx Picture2.hdc, Picture1.hdc, Color, Not Color, HScroll1.Value, True, True, True, False, Resp
        Case 19
            GPX_Rock Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 20
            GPX_Emboss Picture2.hdc, Picture1.hdc, HScroll1.Value / 100, Resp
        Case 21
            GPX_ColorRandomize Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 22
            Label1.Caption = HScroll1.Value - 2
            GPX_RandomicalPoints Picture2.hdc, Picture1.hdc, HScroll1.Value, Color, Resp
        Case 24
            GPX_Shift Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 27
            GPX_Twirl Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 28
            GPX_Swirl Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 29
            GPX_Neon Picture2.hdc, Picture1.hdc, HScroll1.Value, 2, Resp
        Case 30
            GPX_Canvas Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 31
            GPX_Waves Picture2.hdc, Picture1.hdc, HScroll1.Value, HScroll1.Value, HScroll1.Value, True, Resp
        Case 32
            GPX_DetectBorders Picture2.hdc, Picture1.hdc, HScroll1.Value, Color, Not Color, Resp
        Case 35
            GPX_Saturation Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 36
            GPX_FindEdges Picture2.hdc, Picture1.hdc, HScroll1.Value, 2, Resp
        Case 37
            GPX_Hue Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 38
            GPX_AlphaBlend Picture2.hdc, Picture3.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 42
            Picture2.Picture = LoadPicture("")
            Picture2.BackColor = vbWhite
            GPX_Tile Picture2.hdc, Picture1.hdc, HScroll1.Value, HScroll1.Value, 6, Resp
        Case 45
            GPX_AdvancedBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, 25, True, Resp
        Case 47
            GPX_MotionBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, 15, Resp
        Case 48
            GPX_ColorBalance Picture2.hdc, Picture1.hdc, 0, 0, HScroll1.Value, Resp
        Case 49
            GPX_Fragment Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 50
            GPX_FarBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 51
            GPX_RadialBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 52
            GPX_ZoomBlur Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 54
            GPX_Fog Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 55
            GPX_MediumTones Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 56
            GPX_CircularWaves Picture2.hdc, Picture1.hdc, HScroll1.Value, HScroll1.Value, Resp
        Case 57
            GPX_CircularWavesEx Picture2.hdc, Picture1.hdc, HScroll1.Value * 4, HScroll1.Value, Resp
        Case 59
            GPX_OilPaint Picture2.hdc, Picture1.hdc, HScroll1.Value / 50, HScroll1.Value, Resp
        Case 61
            GPX_NotePaper Picture2.hdc, Picture1.hdc, HScroll1.Value, 2, 20, 1, Color, Not Color, Resp
        Case 62
            GPX_FishEyeEx Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 64
            GPX_Cilindrical Picture2.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 65
            GPX_UnsharpMask Picture2.hdc, Picture1.hdc, 2, HScroll1.Value / 2, Resp
        Case 66
            GPX_BlockWaves Picture2.hdc, Picture1.hdc, HScroll1.Value, HScroll1.Value / 2, 1, Resp
        Case 68
            GPX_BlendMode Picture2.hdc, Picture3.hdc, Picture1.hdc, HScroll1.Value, Resp
        Case 69
            GPX_TwirlEx Picture2.hdc, Picture1.hdc, -HScroll1.Value, HScroll1.Value, Resp
        Case 70
            GPX_GlassBlendMode Picture2.hdc, Picture3.hdc, Picture1.hdc, HScroll1.Value / 200, 3, Resp
        Case 71
            GPX_Metallic Picture2.hdc, Picture1.hdc, 4, HScroll1.Value, 1, Resp
        Case 72
            GPX_Metallic Picture2.hdc, Picture1.hdc, 4, HScroll1.Value, 2, Resp
        Case 73
            GPX_Metallic Picture2.hdc, Picture1.hdc, 4, HScroll1.Value, 3, Resp
    End Select
    
    ScrollValue = HScroll1.Value
    Ticks = GetTickCount - Ticks
    Picture2.Refresh
    lblTime.Caption = Ticks & " ms"
    a = Processor
    B = "Command2"
    c = "Hscroll1"
    d = ScrollMax
    e = ScrollMin
    f = ScrollValue
    G = ApplyEffect
    Command6.Enabled = True
Exit Sub

Open App.Path & "\TempFile" For Output As #1
    Print #1, Processor
    Print #1, "Command2"
    Print #1, "Hscroll1"
    Print #1, ScrollMax
    Print #1, ScrollMin
    Print #1, ScrollValue
    Print #1, ApplyEffect
Close #1
        Open App.Path & "\Tempfile" For Input As #1
                Line Input #1, TempStuff
                a = Val(TempStuff)
                Line Input #1, TempStuff
                B = TempStuff
                Line Input #1, TempStuff
                c = TempStuff
                Line Input #1, TempStuff
                d = Val(TempStuff)
                Line Input #1, TempStuff
                e = Val(TempStuff)
                Line Input #1, TempStuff
                f = Val(TempStuff)
                Line Input #1, TempStuff
                TempStuff1 = TempStuff
        Close #1
        If TempStuff1 = "True" Then
            G = True
        Else
            G = False
        End If
    Command6.Enabled = True
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
End Sub

Private Sub mnu3D_Click()
    ChangeControls 39, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (3D)"
    Label1.Caption = ""
End Sub

Private Sub mnu4Corners_Click()
    ChangeControls 40, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (4 Corners)"
    Label1.Caption = ""
End Sub

Private Sub mnuAlias_Click()
    ChangeControls 4, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (AntiAlias)"
    Label1.Caption = ""
End Sub


Private Sub mnuAlphaBlend_Click()
    With CommonDialog1
        .CancelError = False
        .DialogTitle = "Open..."
        .InitDir = LastPath
        .Filter = "Compatible Image Files|*.bmp;*.jpg;*.emf;*.gif;*.rle;*.wmf"
        .flags = cdlOFNHideReadOnly
        .ShowOpen
        If Len(.filename) > 0 Then
            Picture3.Picture = LoadPicture()
            Picture3.PaintPicture LoadPicture(.filename), 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight
        End If
        LastPath = Left(.filename, Len(.filename) - Len(.FileTitle))
    End With
    
    ChangeControls 38, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Alpha Blend)"
    HScroll1_Change
End Sub

Private Sub mnuAmbient_Click()
    ChangeControls 15, Command2, HScroll1, 0, 255, 255
    GraphicalDLL.Caption = "Special Effects - (Ambient Light)"
    HScroll1_Change
End Sub

Private Sub mnuAscii_Click()
    ChangeControls 14, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Ascii Morph)"
    Label1.Caption = ""
End Sub

Private Sub mnuBack_Click()
    ChangeControls 17, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Replace Colors)"
    HScroll1_Change
End Sub

Private Sub mnuBackN_Click()
    ChangeControls 13, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Backdrop Removal)"
    HScroll1_Change
End Sub

Private Sub mnuBlendModes_Click()
    With CommonDialog1
        .CancelError = False
        .DialogTitle = "Open..."
        .InitDir = LastPath
        .Filter = "Compatible Image Files|*.bmp;*.jpg;*.emf;*.gif;*.rle;*.wmf"
        .flags = cdlOFNHideReadOnly
        .ShowOpen
        If Len(.filename) > 0 Then
            Picture3.Picture = LoadPicture()
            Picture3.PaintPicture LoadPicture(.filename), 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight
        End If
        LastPath = Left(.filename, Len(.filename) - Len(.FileTitle))
    End With

    ChangeControls 68, Command2, HScroll1, 0, 24, 0
    GraphicalDLL.Caption = "Special Effects - (Blend Modes)"
    HScroll1_Change
End Sub

Private Sub mnuBlockWaves_Click()
    ChangeControls 66, Command2, HScroll1, 0, 20, 0
    GraphicalDLL.Caption = "Special Effects - (Block Waves)"
    HScroll1_Change
End Sub

Private Sub mnuBlur_Click()
    ChangeControls 33, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Blur)"
    Label1.Caption = ""
End Sub

Private Sub mnuBorders_Click()
    ChangeControls 32, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Detect Borders)"
    HScroll1_Change
End Sub

Private Sub mnuBrilho_Click()
    ChangeControls 1, Command2, HScroll1, -255, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Brightness)"
    HScroll1_Change
End Sub

Private Sub mnuCanvas_Click()
    ChangeControls 30, Command2, HScroll1, 0, Picture1.ScaleWidth, 0
    GraphicalDLL.Caption = "Special Effects - (Canvas)"
    HScroll1_Change
End Sub

Private Sub mnuCaricature_Click()
    ChangeControls 41, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Caricature)"
    Label1.Caption = ""
End Sub

Private Sub mnuCilindrical_Click()
    ChangeControls 64, Command2, HScroll1, -30, 30, 0
    GraphicalDLL.Caption = "Special Effects - (Cilindrical)"
    HScroll1_Change
End Sub

Private Sub mnuCircularWaves_Click()
    ChangeControls 56, Command2, HScroll1, 0, 20, 0
    GraphicalDLL.Caption = "Special Effects - (Circular Waves)"
    HScroll1_Change
End Sub

Private Sub mnuCircularWavesEx_Click()
    ChangeControls 57, Command2, HScroll1, 0, 20, 0
    GraphicalDLL.Caption = "Special Effects - (Circular Waves Ex)"
    HScroll1_Change
End Sub

Private Sub mnuColBalance_Click()
    ChangeControls 48, Command2, HScroll1, -255, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Color Balance)"
    HScroll1_Change
End Sub

Private Sub mnuContraste_Click()
    ChangeControls 2, Command2, HScroll1, 0, 255, 100
    GraphicalDLL.Caption = "Special Effects - (Contrast)"
    HScroll1_Change
End Sub

Private Sub mnuDiffuse_Click()
    ChangeControls 18, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Diffuse)"
    Label1.Caption = ""
End Sub

Private Sub mnuEightColors_Click()
    ChangeControls 9, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (8 colors)"
    Label1.Caption = ""
End Sub

Private Sub mnuEmbEng_Click()
    ChangeControls 20, Command2, HScroll1, -255, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Emboss / Engrave)"
    HScroll1_Change
End Sub

Private Sub mnuFarBlur_Click()
    ChangeControls 50, Command2, HScroll1, 0, 50, 0
    GraphicalDLL.Caption = "Special Effects - (Far Blur)"
    HScroll1_Change
End Sub

Private Sub mnuFindEdges_Click()
    ChangeControls 36, Command2, HScroll1, 0, 5, 0
    GraphicalDLL.Caption = "Special Effects - (Find Edges)"
    HScroll1_Change
End Sub

Private Sub mnuFishEye_Click()
    ChangeControls 26, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (FishEye)"
    Label1.Caption = ""
End Sub

Private Sub mnuFishEyeEx_Click()
    ChangeControls 62, Command2, HScroll1, -255, 255, 0
    GraphicalDLL.Caption = "Special Effects - (FishEye Ex)"
    HScroll1_Change
End Sub

Private Sub mnuFlipB_Click()
    Processor = 102
    GPX_Flip Picture1.hdc, Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, 1, 1, Resp
    Picture1.Refresh
    a = Processor
    Command6.Enabled = True
End Sub

Private Sub mnuFlipH_Click()
    Processor = 100
    GPX_Flip Picture1.hdc, Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, 1, 0, Resp
    Picture1.Refresh
    a = Processor
    Command6.Enabled = True
End Sub

Private Sub mnuFlipV_Click()
    Processor = 101
    GPX_Flip Picture1.hdc, Picture1.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 1, Resp
    Picture1.Refresh
    a = Processor
    Command6.Enabled = True
End Sub

Private Sub mnuFog_Click()
    ChangeControls 54, Command2, HScroll1, 0, 127, 0
    HScroll1_Change
End Sub

Private Sub mnuFragment_Click()
    ChangeControls 49, Command2, HScroll1, 0, 50, 0
    GraphicalDLL.Caption = "Special Effects - (Fragment)"
    HScroll1_Change
End Sub

Private Sub mnuFrostGlass_Click()
    ChangeControls 60, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Frost Glass)"
    Label1.Caption = ""
End Sub

Private Sub mnuGamma_Click()
    ChangeControls 5, Command2, HScroll1, 0, 255, 25
    GraphicalDLL.Caption = "Special Effects - (Gamma)"
    HScroll1_Change
End Sub

Private Sub mnuGlassBlendMode_Click()
    With CommonDialog1
        .CancelError = False
        .DialogTitle = "Open..."
        .InitDir = LastPath
        .Filter = "Compatible Image Files|*.bmp;*.jpg;*.emf;*.gif;*.rle;*.wmf"
        .flags = cdlOFNHideReadOnly
        .ShowOpen
        If Len(.filename) > 0 Then
            Picture3.Picture = LoadPicture()
            Picture3.PaintPicture LoadPicture(.filename), 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight
        End If
        LastPath = Left(.filename, Len(.filename) - Len(.FileTitle))
    End With

    ChangeControls 70, Command2, HScroll1, -255, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Glass Blend Mode)"
    HScroll1_Change
End Sub

Private Sub mnuGold_Click()
    ChangeControls 72, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Gold)"
    HScroll1_Change
End Sub

Private Sub mnuGrayScale_Click()
    ChangeControls 6, Command2, HScroll1, -255, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Gray Scale)"
    HScroll1_Change
End Sub

Private Sub mnuHue_Click()
    ChangeControls 37, Command2, HScroll1, 0, 350, 0
    GraphicalDLL.Caption = "Special Effects - (Hue)"
    HScroll1_Change
End Sub

Private Sub mnuIce_Click()
    ChangeControls 73, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Ice)"
    HScroll1_Change
End Sub

Private Sub mnuInvertion_Click()
    ChangeControls 7, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Invertion)"
    HScroll1_Change
End Sub

Private Sub mnuMediumTones_Click()
    ChangeControls 55, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Medium Tones)"
    HScroll1_Change
End Sub

Private Sub mnuMelt_Click()
    ChangeControls 25, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Melt)"
    Label1.Caption = ""
End Sub

Private Sub mnuMetallic_Click()
    ChangeControls 71, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Metallic)"
    HScroll1_Change
End Sub

Private Sub mnuMono_Click()
    ChangeControls 10, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Monochrome)"
    HScroll1_Change
End Sub

Private Sub mnuMoreBlur_Click()
    ChangeControls 45, Command2, HScroll1, 0, 10, 0
    GraphicalDLL.Caption = "Special Effects - (More Blur)"
    HScroll1_Change
End Sub

Private Sub mnuMosaico_Click()
    ChangeControls 12, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Mosaic)"
    HScroll1_Change
End Sub

Private Sub mnuMotionBlur_Click()
    ChangeControls 47, Command2, HScroll1, 0, 360, 0
    GraphicalDLL.Caption = "Special Effects - (Motion Blur)"
    HScroll1_Change
End Sub

Private Sub mnuNeon_Click()
    ChangeControls 29, Command2, HScroll1, 0, 5, 0
    GraphicalDLL.Caption = "Special Effects - (Neon)"
    HScroll1_Change
End Sub

Private Sub mnuNitidez_Click()
    ChangeControls 3, Command2, HScroll1, 0, 60, 0
    GraphicalDLL.Caption = "Special Effects - (Sharpen)"
    HScroll1_Change
End Sub

Private Sub mnuNoise_Click()
    ChangeControls 21, Command2, HScroll1, -255, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Noise)"
    HScroll1_Change
End Sub

Private Sub mnuNotePaper_Click()
    ChangeControls 61, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Note Paper)"
    HScroll1_Change
End Sub

Private Sub mnuOilPaint_Click()
    ChangeControls 59, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (OilPaint)"
    HScroll1_Change
End Sub

Private Sub mnuOpen_Click()
    With CommonDialog1
        .CancelError = False
        .DialogTitle = "Open..."
        .InitDir = LastPath
        .Filter = "Compatible Image Files|*.bmp;*.jpg;*.emf;*.gif;*.rle;*.wmf"
        .flags = cdlOFNHideReadOnly
        .ShowOpen
        If Len(.filename) > 0 Then
            Picture1.Picture = LoadPicture()
            Picture2.Picture = LoadPicture()
            Picture1.PaintPicture LoadPicture(.filename), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
            Picture2.PaintPicture LoadPicture(.filename), 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight
            ImagePath = .filename
        End If
        LastPath = Left(.filename, Len(.filename) - Len(.FileTitle))
    End With
End Sub

Private Sub mnuPolar_Click()
    ChangeControls 58, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Polar Coordinates)"
    Label1.Caption = ""
End Sub

Private Sub mnuRadialBlur_Click()
    ChangeControls 51, Command2, HScroll1, 0, 30, 0
    GraphicalDLL.Caption = "Special Effects - (Radial Blur)"
    HScroll1_Change
End Sub

Private Sub mnuRainDrop_Click()
    ChangeControls 63, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Rain Drop)"
    Label1.Caption = ""
End Sub

Private Sub mnuRandomPoints_Click()
    ChangeControls 22, Command2, HScroll1, 2, 102, 2
    GraphicalDLL.Caption = "Special Effects - (Random Points)"
    HScroll1_Change
End Sub

Private Sub mnuReduceColors_Click()
    ChangeControls 8, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Reduce Colors)"
    HScroll1_Change
End Sub

Private Sub mnuRelief_Click()
    ChangeControls 34, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Relief)"
    Label1.Caption = ""
End Sub

Private Sub mnuRock_Click()
    ChangeControls 19, Command2, HScroll1, 0, 6, 0
    GraphicalDLL.Caption = "Special Effects - (Rock)"
    HScroll1_Change
End Sub

Private Sub mnuRoll_Click()
    ChangeControls 43, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Roll)"
    Label1.Caption = ""
End Sub

Private Sub mnuSair_Click()
    End
End Sub

Private Sub mnuSaturation_Click()
    ChangeControls 35, Command2, HScroll1, -255, 512, 0
    HScroll1_Change
End Sub

Private Sub mnuSave_Click()
    With CommonDialog1
        .CancelError = False
        .DialogTitle = "Save as..."
        .DefaultExt = ".bmp"
        .InitDir = LastPath
        .Filter = "Bitmaps|*.bmp"
        .flags = cdlOFNHideReadOnly
        .ShowSave
        SavePicture Picture2.Image, .filename
    End With
End Sub

Private Sub mnuSepia_Click()
    ChangeControls 11, Command2, HScroll1
    Label1.Caption = ""
End Sub

Private Sub mnuShift_Click()
    ChangeControls 24, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Shift)"
    HScroll1_Change
End Sub

Private Sub mnuSmartBlur_Click()
    ChangeControls 44, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Smart Blur)"
    Label1.Caption = ""
End Sub

Private Sub mnuSoftnerBlur_Click()
    ChangeControls 46, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Softner Blur)"
    Label1.Caption = ""
End Sub

Private Sub mnuSol_Click()
    ChangeControls 23, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Solarize)"
    Label1.Caption = ""
End Sub

Private Sub mnuStretchHisto_Click()
    ChangeControls 67, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Stretch Histogram)"
    Label1.Caption = ""
End Sub

Private Sub mnuSwirl_Click()
    ChangeControls 28, Command2, HScroll1, -255, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Swirl)"
    HScroll1_Change
End Sub

Private Sub mnuTile_Click()
    ChangeControls 42, Command2, HScroll1, 0, 100, 0
    GraphicalDLL.Caption = "Special Effects - (Tile)"
    HScroll1_Change
End Sub

Private Sub mnuTone_Click()
    ChangeControls 16, Command2, HScroll1, 0, 255, 0
    GraphicalDLL.Caption = "Special Effects - (Tone)"
    HScroll1_Change
End Sub


Private Sub mnuTwirl_Click()
    ChangeControls 27, Command2, HScroll1, -100, 100, 0
    GraphicalDLL.Caption = "Special Effects - (Twirl)"
    HScroll1_Change
End Sub

Private Sub mnuTwirlEx_Click()
    ChangeControls 69, Command2, HScroll1, -100, 100, 0
    GraphicalDLL.Caption = "Special Effects - (Twirl Ex)"
    HScroll1_Change
End Sub

Private Sub mnuUnsharpMask_Click()
    ChangeControls 65, Command2, HScroll1, 0, 10, 0
    GraphicalDLL.Caption = "Special Effects - (Unsharp Mask)"
    HScroll1_Change
End Sub

Private Sub mnuWave_Click()
    ChangeControls 31, Command2, HScroll1, 0, 20, 0
    GraphicalDLL.Caption = "Special Effects - (Waves)"
    HScroll1_Change
End Sub

Private Sub mnuWebColors_Click()
    ChangeControls 53, Command2, HScroll1
    GraphicalDLL.Caption = "Special Effects - (Web Colors)"
    Label1.Caption = ""
End Sub

Private Sub mnuZoomBlur_Click()
    ChangeControls 52, Command2, HScroll1, 0, 200, 0
    GraphicalDLL.Caption = "Special Effects - (Zoom Blur)"
    HScroll1_Change
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Color As Long
On Error Resume Next

If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture1.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
    If (ColorPickTool And (Button = vbLeftButton)) Then
        Color = GetPixel(Picture1.hdc, X, Y)
        Text2.BackColor = Color
        Text3.Text = Color And &HFF
        Text4.Text = (Color And &HFF00&) / &H100
        Text5.Text = (Color And &HFF0000) / &H10000
    End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Color As Long
'On Error Resume Next

If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture2.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
    
    If (ColorPickTool And (Button = vbLeftButton)) Then
        Color = GetPixel(Picture2.hdc, X, Y)
        Text2.BackColor = Color
        Text3.Text = Color And &HFF
        Text4.Text = (Color And &HFF00&) / &H100
        Text5.Text = (Color And &HFF0000) / &H10000
    End If
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture3.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If ((Not IsNumeric(Chr(KeyAscii))) And (Chr(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If ((Not IsNumeric(Chr(KeyAscii))) And (Chr(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If ((Not IsNumeric(Chr(KeyAscii))) And (Chr(KeyAscii) <> vbBack)) Then
        KeyAscii = 0
    End If
End Sub

