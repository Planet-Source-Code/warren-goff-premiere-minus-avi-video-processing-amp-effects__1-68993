VERSION 5.00
Begin VB.Form frmAVICreator 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Avi Creation"
   ClientHeight    =   3525
   ClientLeft      =   3105
   ClientTop       =   3225
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAVICreator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTab 
      Height          =   4155
      Index           =   1
      Left            =   12165
      ScaleHeight     =   4095
      ScaleWidth      =   5475
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   5535
      Begin VB.ComboBox cboHandler 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   60
         Width           =   4335
      End
      Begin VB.CommandButton cmdLoadPalette 
         Caption         =   "&Load..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Top             =   2580
         Width           =   975
      End
      Begin VB.PictureBox picPalette 
         AutoRedraw      =   -1  'True
         Enabled         =   0   'False
         Height          =   1335
         Left            =   1080
         ScaleHeight     =   1275
         ScaleWidth      =   4275
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1200
         Width           =   4335
      End
      Begin VB.ComboBox cboPalette 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   840
         Width           =   4335
      End
      Begin VB.ComboBox cboBitsPerPixel 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   420
         Width           =   4335
      End
      Begin VB.Label lblInfo 
         Caption         =   "&Handler"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   9
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lblPalette 
         Caption         =   "Palette:"
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   900
         Width           =   915
      End
      Begin VB.Label lblInfo 
         Caption         =   "Bits/pixel:"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   11
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   4155
      Index           =   0
      Left            =   6480
      ScaleHeight     =   4095
      ScaleWidth      =   5475
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6210
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtFrameDuration 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Text            =   "50"
         Top             =   2100
         Width           =   3975
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1080
         TabIndex        =   6
         Top             =   1560
         Width           =   3975
      End
      Begin VB.CommandButton cmdPick 
         Caption         =   "..."
         Height          =   315
         Left            =   5100
         TabIndex        =   4
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtAVIFile 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1080
         Width           =   3975
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmAVICreator.frx":08CA
         Height          =   855
         Index           =   10
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   5355
      End
      Begin VB.Label lblInfo 
         Caption         =   " Frame (ms)"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   7
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label lblInfo 
         Caption         =   "AVI Name:"
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label lblInfo 
         Caption         =   "File name:"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   1140
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdPickDirectory 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create AVI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1275
      Width           =   1755
   End
   Begin VB.CommandButton cmdPickBitmapStrip 
      Caption         =   "..."
      Height          =   345
      Left            =   10650
      TabIndex        =   43
      Top             =   2820
      Width           =   795
   End
   Begin VB.PictureBox picWizard 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   45
      Picture         =   "frmAVICreator.frx":09A1
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   101
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   5805
      TabIndex        =   32
      Top             =   5205
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   435
      Left            =   4485
      TabIndex        =   31
      Top             =   5205
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   435
      Left            =   3225
      TabIndex        =   30
      Top             =   5205
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picTab 
      Height          =   4155
      Index           =   2
      Left            =   5715
      ScaleHeight     =   4095
      ScaleWidth      =   5475
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3930
      Visible         =   0   'False
      Width           =   5535
      Begin VB.ComboBox cboImageSource 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   0
         Width           =   4395
      End
      Begin VB.PictureBox pnlImageSource 
         BorderStyle     =   0  'None
         Height          =   3615
         Index           =   0
         Left            =   1335
         ScaleHeight     =   3615
         ScaleWidth      =   5475
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1650
         Width           =   5475
         Begin VB.TextBox txtDirectory 
            Height          =   315
            Left            =   1020
            TabIndex        =   25
            Top             =   0
            Width           =   3975
         End
         Begin VB.PictureBox picDirectoryPreview 
            AutoRedraw      =   -1  'True
            Height          =   3195
            Left            =   1020
            ScaleHeight     =   3135
            ScaleWidth      =   4335
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   360
            Width           =   4395
         End
         Begin VB.Label lblDirectory 
            Caption         =   "Directory:"
            Height          =   255
            Left            =   60
            TabIndex        =   24
            Top             =   60
            Width           =   915
         End
      End
      Begin VB.PictureBox pnlImageSource 
         BorderStyle     =   0  'None
         Height          =   3615
         Index           =   1
         Left            =   15
         ScaleHeight     =   3615
         ScaleWidth      =   5475
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   480
         Width           =   5475
         Begin VB.TextBox txtYCells 
            Height          =   315
            Left            =   3540
            TabIndex        =   23
            Text            =   "1"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtXCells 
            Height          =   315
            Left            =   1020
            TabIndex        =   21
            Text            =   "1"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtBitmapStrip 
            Height          =   315
            Left            =   1020
            TabIndex        =   19
            Top             =   0
            Width           =   3975
         End
         Begin VB.PictureBox picStripPreview 
            AutoRedraw      =   -1  'True
            Height          =   2835
            Left            =   0
            ScaleHeight     =   185
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   289
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   0
            Width           =   4395
         End
         Begin VB.Label lblInfo 
            Caption         =   "Y Cells"
            Height          =   255
            Index           =   8
            Left            =   2580
            TabIndex        =   22
            Top             =   420
            Width           =   915
         End
         Begin VB.Label lblInfo 
            Caption         =   "X Cells"
            Height          =   255
            Index           =   7
            Left            =   60
            TabIndex        =   20
            Top             =   420
            Width           =   915
         End
         Begin VB.Label lblInfo 
            Caption         =   "Bitmap File:"
            Height          =   255
            Index           =   6
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   915
         End
      End
      Begin VB.Label lblInfo 
         Caption         =   "Images:"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   4155
      Index           =   3
      Left            =   7065
      ScaleHeight     =   4095
      ScaleWidth      =   5475
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1155
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdCreate 
         Caption         =   "&Create"
         Height          =   435
         Left            =   1860
         TabIndex        =   29
         Top             =   2880
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         Caption         =   "Ready to create your AVI."
         Height          =   255
         Index           =   9
         Left            =   60
         TabIndex        =   26
         Top             =   60
         Width           =   5355
      End
      Begin VB.Label lblInfo 
         Caption         =   "Click Create to build the AVI and write it to the file:"
         Height          =   255
         Index           =   11
         Left            =   60
         TabIndex        =   28
         Top             =   2520
         Width           =   5355
      End
      Begin VB.Label lblSummary 
         Caption         =   "Summary:"
         Height          =   1815
         Left            =   525
         TabIndex        =   27
         Top             =   465
         Width           =   4695
      End
   End
   Begin VB.Line linSep 
      BorderColor     =   &H80000010&
      X1              =   8550
      X2              =   15690
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Label lblStage 
      BackColor       =   &H80000010&
      Caption         =   " Main AVI Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   10110
      TabIndex        =   0
      Top             =   960
      Width           =   5535
   End
End
Attribute VB_Name = "frmAVICreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' AVI Creator object
Private m_cAVI As New cAVICreator
Private m_cVH As New cVideoHandlers

' DC for rendering bitmaps
Private m_cDC As New cMemDC
' Bitmap Strip image
Private m_cBmp As New cBmp

' Wizard control
Private m_iWizardPanel As Long
Private m_iPanelCount As Long
Private m_sPanelCaption() As String

Private Sub wizardNavigate(ByVal iDir As Long)
Dim iNewPanel As Long
Dim i As Long
Dim sMsg As String
Dim offendingCtl As Control

   If (iDir = -1) Then
      iNewPanel = m_iWizardPanel - 1
   Else
      If wizardValidate(m_iWizardPanel, sMsg, offendingCtl) Then
         iNewPanel = m_iWizardPanel + 1
      Else
         If Not (offendingCtl Is Nothing) Then
            offendingCtl.SetFocus
         End If
         MsgBox sMsg, vbInformation
         Exit Sub
      End If
   End If
   
   If Not (iNewPanel = m_iWizardPanel) Then
      picTab(iNewPanel).Move picTab(0).Left, picTab(0).TOp, picTab(0).Width, picTab(0).Height
      picTab(iNewPanel).BorderStyle = 0
      picTab(iNewPanel).Visible = True
      If (m_iWizardPanel >= 0) Then
         picTab(m_iWizardPanel).Visible = False
      End If
      m_iWizardPanel = iNewPanel
      
      lblStage.Caption = m_sPanelCaption(m_iWizardPanel + 1)
      
      cmdBack.Enabled = (m_iWizardPanel > 0)
      cmdNext.Enabled = (m_iWizardPanel < m_iPanelCount - 1)
   End If
   
End Sub

Private Function wizardValidate(ByVal iPanel As Long, ByRef sMsg As String, ByRef offendingCtl As Control) As Boolean
   wizardValidate = True
   Select Case iPanel
   Case 0
      wizardValidate = validateNamePanel(sMsg, offendingCtl)
   Case 1
      wizardValidate = validateTypePanel(sMsg, offendingCtl)
   Case 2
      wizardValidate = validateSourcePanel(sMsg, offendingCtl)
   Case 3
   End Select
End Function

Private Function validateNamePanel(ByRef sMsg As String, ByRef offendingCtl As Control) As Boolean
   If Len(txtAVIFile.Text) > 0 Then
      Dim lFrame As Long
      On Error Resume Next
      lFrame = CLng(txtFrameDuration.Text)
      If (lFrame > 0) And (Err.Number = 0) Then
         
         m_cAVI.Filename = txtAVIFile.Text
         m_cAVI.Name = txtName.Text
         m_cAVI.FrameDuration = lFrame
         
         validateNamePanel = True
         
      Else
         sMsg = "Frame duration must be entered."
         Set offendingCtl = txtFrameDuration
      End If
   Else
      sMsg = "Must choose a file name to write to."
      Set offendingCtl = txtAVIFile
   End If
End Function

Private Function validateTypePanel(ByRef sMsg As String, ByRef offendingCtl As Control) As Boolean
   If (cboHandler.ListIndex > -1) Then
      If (cboBitsPerPixel.ListIndex > -1) Then
         
         m_cAVI.bitsPerPixel = cboBitsPerPixel.ItemData(cboBitsPerPixel.ListIndex)
         m_cAVI.VideoHandlerFourCC = m_cVH.Handler(cboHandler.ListIndex + 1).FourCC
         
         validateTypePanel = True
      Else
         sMsg = "Must choose a colour depth for the AVI."
         Set offendingCtl = cboBitsPerPixel
      End If
   Else
      sMsg = "Must choose a video handler to write with."
      Set offendingCtl = cboHandler
   End If
End Function

Private Function validateSourcePanel(ByRef sMsg As String, ByRef offendingCtl As Control) As Boolean
   If (cboImageSource.ListIndex > -1) Then
      Select Case cboImageSource.ListIndex
      Case 0
         ' Directory
         If directoryExists(txtDirectory.Text) And (m_cBmp.Height > 0) Then
                     
            validateSourcePanel = True
         Else
            sMsg = "Choose a valid directory containing images."
            Set offendingCtl = txtDirectory
         End If
      Case 1
         ' Bitmap strip
         If fileExists(txtBitmapStrip.Text) Then
            On Error Resume Next
            Dim lXCell As Long
            lXCell = CLng(txtXCells.Text)
            If (lXCell > 0) Then
               Dim lYCell As Long
               lYCell = CLng(txtYCells.Text)
               If (lYCell > 0) Then
                  pSetSummary
                  validateSourcePanel = True
               Else
                  sMsg = "Set the number of cells in the X direction to a positive integer."
                  Set offendingCtl = txtYCells
               End If
            Else
               sMsg = "Set the number of cells in the X direction to a positive integer."
               Set offendingCtl = txtXCells
            End If
         Else
            sMsg = "Choose a bitmap strip file."
            Set offendingCtl = txtBitmapStrip
         End If
      End Select
   End If
End Function

Private Sub pSetSummary()
Dim sSummary As String
   sSummary = "Filename: " & m_cAVI.Filename & vbCrLf
   sSummary = sSummary & "Name: " & m_cAVI.Name & vbCrLf
   sSummary = sSummary & "Frame Length: " & m_cAVI.FrameDuration & vbCrLf
   sSummary = sSummary & "Bits/pixel: " & m_cAVI.bitsPerPixel & vbCrLf
   sSummary = sSummary & "Handler: " & m_cAVI.FourCCToString(m_cAVI.VideoHandlerFourCC) & vbCrLf
   sSummary = sSummary & vbCrLf
   sSummary = sSummary & "Image source: "
   If (cboImageSource.ListIndex = 0) Then
      sSummary = sSummary & "Directory " & vbCrLf & vbTab & txtDirectory.Text
   Else
      sSummary = sSummary & "Bitmap Strip" & vbCrLf & vbTab & txtBitmapStrip.Text
   End If
   lblSummary.Caption = sSummary
End Sub

Private Sub pRenderPalette()
   
   picPalette.Cls

   Dim cP As cPalette
   Set cP = m_cAVI.Palette
   If Not cP Is Nothing Then ' else > 8bpp
      
      Dim Index As Long
      Dim x As Long
      Dim y As Long
      Dim palItemWidth As Long
      Dim palItemHeight As Long
      
      palItemWidth = picPalette.ScaleWidth \ 16
      If (cP.Count > 16) Then
         palItemHeight = picPalette.ScaleHeight \ (cP.Count \ 16)
      Else
         palItemHeight = 16
      End If
      
      For Index = 0 To cP.Count - 1
         picPalette.Line (x, y)-(x + palItemWidth, y + palItemHeight), RGB(cP.Red(Index), cP.Green(Index), cP.Blue(Index)), BF
         x = x + palItemWidth
         If (x > picPalette.ScaleWidth) Then
            x = 0
            y = y + palItemHeight
         End If
      Next Index
      
   End If
   picPalette.Refresh

End Sub

Private Function directoryExists(ByVal sDir As String) As Boolean
Dim sTestDir As String
On Error Resume Next
   sTestDir = Dir(sDir, vbDirectory)
   If (Err.Number = 0) And (Len(sTestDir) > 0) Then
      directoryExists = ((GetAttr(sDir) And vbDirectory) = vbDirectory)
   End If
End Function

Private Function fileExists(ByVal sFile As String) As Boolean
Dim sTestFile As String
On Error Resume Next
   sTestFile = Dir(sFile)
   If (Err.Number = 0) And (Len(sTestFile) > 0) Then
      fileExists = ((GetAttr(sFile) And vbDirectory) = 0)
   End If
End Function

Private Function fileNameToName(ByVal sFile As String) As String
Dim iExtPos As Long
Dim iSlashPos As Long
Dim i As Long
Dim sChar As String
   For i = Len(sFile) To 1 Step -1
      sChar = Mid(sFile, i, 1)
      If (sChar = ".") Then
         iExtPos = i
      ElseIf (sChar = "\") Then
         iSlashPos = i
         Exit For
      End If
   Next i
   If (iSlashPos > 0) Then
      If (iExtPos > iSlashPos) Then
         fileNameToName = Mid(sFile, iSlashPos + 1, iExtPos - iSlashPos - 1)
      Else
         fileNameToName = Mid(sFile, iSlashPos + 1)
      End If
   End If
End Function

Private Sub cboBitsPerPixel_Click()
Dim bPalette As Boolean
Dim iHandler As Long
Dim lFourCC As Long
   
   ' Choose default
   lFourCC = m_cVH.SuggestedVideoHandlerFourCC(cboBitsPerPixel.ItemData(cboBitsPerPixel.ListIndex))
   iHandler = m_cVH.IndexForFourCC(lFourCC)
   cboHandler.ListIndex = iHandler - 1
   
   ' Enable palette controls as appropriate
   bPalette = (cboBitsPerPixel.ListIndex = 0)
   lblPalette.ForeColor = IIf(bPalette, vbWindowText, vbButtonShadow)
   cboPalette.Enabled = bPalette
   cmdLoadPalette.Enabled = ((cboPalette.ListIndex = 3) And bPalette)
   If (bPalette) Then
      pRenderPalette
   Else
      picPalette.Cls
   End If
   '
End Sub

Private Sub cboImageSource_Click()
   pnlImageSource(cboImageSource.ListIndex).Visible = True
   pnlImageSource((1 - cboImageSource.ListIndex)).Visible = False
End Sub

Private Sub cboPalette_Click()
Dim cP As New cPalette
   
   Select Case cboPalette.ListIndex
   Case 0
      cP.Create16Colour
      Set m_cAVI.Palette = cP
   Case 1
      cP.CreateHalfTone
      Set m_cAVI.Palette = cP
   Case 2
      cP.CreateWebSafe
      Set m_cAVI.Palette = cP
   Case 3
   End Select
   cmdLoadPalette.Enabled = (cboPalette.ListIndex = 3)
   pRenderPalette
   
End Sub

Private Sub cmdBack_Click()
   wizardNavigate -1
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdCreate_Click()
Dim cb As cBmp, i As Long
Dim iImageIndex As Long
Dim bStreamOpen As Boolean
   
On Error GoTo ErrorHandler
   cmdPickDirectory.Visible = False
  DoEvents
   iImageIndex = 1
   Set cb = pbGetImage(iImageIndex)
   If Not (cb Is Nothing) Then
      
      m_cDC.SelectObject cb
      m_cDC.PaintPicture Me.hdc, 0  '640
      m_cDC.UnselectObject
      
      m_cAVI.StreamCreate cb
      
      bStreamOpen = True
      
      Do
         DoEvents
         iImageIndex = iImageIndex + 1
         Set cb = pbGetImage(iImageIndex)
         If Not (cb Is Nothing) Then
         
            m_cDC.SelectObject cb
            m_cDC.PaintPicture Me.hdc, 0    '640
            m_cDC.UnselectObject

            m_cAVI.StreamAdd cb
         End If
          i = i + 1
            If i / 2 = Int(i / 2) Then
                'cmdPickDirectory.Caption = "<_>"
                Me.Caption = "<_>"
            Else
                'cmdPickDirectory.Caption = "|"
                Me.Caption = " "
            End If
         
      Loop While Not (cb Is Nothing)
      
      m_cAVI.StreamClose
      bStreamOpen = False
      
   End If
   If Dir(App.Path & "\MyAvi.avi") = "" Then
        FileCopy txtAVIFile.Text, App.Path & "\MyAvi.avi"
    Else
        FileCopy txtAVIFile.Text, App.Path & "\MyAviAlt.avi"
    End If
   cmdPickDirectory.Visible = True
   cmdPickDirectory.Caption = "AVI created!"
   Me.Caption = "AVI created!"
   Delay 1
   'MsgBox "AVI created successfully both to Desktop and application directory.", vbInformation
   Unload Me
   Exit Sub
   
ErrorHandler:
   If bStreamOpen Then
      m_cAVI.StreamClose
   End If
   MsgBox "An error occurred whilst creating the AVI: " & Err.Description, vbExclamation
   Exit Sub
End Sub

Private Function pbGetImage(ByVal iImageIndex As Long) As cBmp
Static sDir As String
Static sBaseDir As String
Static cellWidth As Long
Static cellHeight As Long
Static xCell As Long
Static yCell As Long

   If (iImageIndex = 1) Then
      
      ' Initialise
      If (cboImageSource.ListIndex = 0) Then
         sBaseDir = txtDirectory.Text
         If (Right(sBaseDir, 1) <> "\") Then sBaseDir = sBaseDir & "\"
         sDir = Dir(sBaseDir & "*.bmp")
      Else
         cellWidth = m_cBmp.Width \ CLng(txtXCells.Text)
         cellHeight = m_cBmp.Height \ CLng(txtYCells.Text)
         xCell = 0
         yCell = 0
      End If
   
   Else
   
      ' Move to next item
      If (cboImageSource.ListIndex = 0) Then
         sDir = Dir
      Else
         xCell = xCell + cellWidth
         If (xCell >= m_cBmp.Width) Then
            xCell = 0
            yCell = yCell + cellHeight
         End If
      End If
   
   End If
   
   ' Get the image:
   Dim cb As New cBmp
   If (cboImageSource.ListIndex = 0) Then
      If Len(sDir) > 0 Then
         cb.Load sBaseDir & sDir
         Set pbGetImage = cb
      End If
   Else
      If (yCell < m_cBmp.Height) Then
         
         m_cDC.SelectObject m_cBmp
   
         cb.Create cellWidth, cellHeight
         Dim cDC As New cMemDC
         cDC.Create
         cDC.SelectObject cb
         m_cDC.PaintPicture cDC.hdc, 0, 0, cellWidth, cellHeight, xCell, yCell, cellWidth, cellHeight
         cDC.UnselectObject
         
         Set pbGetImage = cb
         
         m_cDC.UnselectObject
      End If
   End If
      
End Function

Private Sub cmdLoadPalette_Click()
Dim cD As New cCommonDialog
Dim sFile As String
   If (cD.VBGetOpenFileName(sFile, _
      Filter:="Palette Files (*.PAL)|*.PAL|All Files (*.*)|*.*", _
      DefaultExt:="PAL", _
      Owner:=Me.hwnd)) Then
      m_cAVI.Palette.LoadFromJASCFile sFile
      pRenderPalette
   End If
End Sub

Private Sub cmdNext_Click()
   wizardNavigate 1
End Sub

Private Sub cmdPick_Click()
   Dim cD As New cCommonDialog
   Dim sFile As String
   Dim sOrig As String
   sOrig = txtName.Text
   If (cD.VBGetSaveFileName(sFile, _
      Filter:="AVI Files (*.AVI)|*.AVI|All Files (*.*)|*.*", _
      DefaultExt:="AVI", _
      Owner:=Me.hwnd)) Then
      txtAVIFile.Text = sFile
      If Len(txtName.Text) = 0 Or (txtName.Text = fileNameToName(sOrig)) Then
         txtName.Text = fileNameToName(sFile)
      End If
   End If
End Sub

Private Sub cmdPickDirectory_Click()
'Dim cBF As New cBrowseForFolder
'Me.Hide
Dim sFolder As String
   sFolder = App.Path & "\Images1\"      'cBF.BrowseForFolder()
   If Len(sFolder) > 0 Then
      txtDirectory.Text = sFolder
      pLoadDirectoryImages
      pRenderDirectoryImages
   End If
   DoEvents
   cmdNext_Click
   cmdCreate_Click
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
cmdNext_Click
cmdNext_Click
End Sub

Private Sub Form_Load()
Dim DesktopPath As String, FormatNow As String
DesktopPath = GetShellFolderPath(&H0)
FormatNow = DesktopPath & "\" & Format(Now, "ddmmyyhhmmss") & "MyAvi.avi"
Me.Left = 0
Me.TOp = 0
txtAVIFile.Text = FormatNow 'App.Path & "\MyAvi.avi"
txtName.Text = FormatNow        'App.Path & "\MyAvi.avi"
    SetTopMostWindow Me.hwnd, True
' Video Handler options:
   Dim i As Long
   For i = 1 To m_cVH.HandlerCount
      cboHandler.AddItem m_cVH.Handler(i).Name & " (" & m_cVH.Handler(i).Description & ")"
   Next i
   
   ' Bits/pixel options:
   cboBitsPerPixel.AddItem "8 Bits/Pixel"
   cboBitsPerPixel.ItemData(cboBitsPerPixel.NewIndex) = 8
   cboBitsPerPixel.AddItem "24 Bits/Pixel"
   cboBitsPerPixel.ItemData(cboBitsPerPixel.NewIndex) = 24
   cboBitsPerPixel.ListIndex = 1
   
   ' Palette options
   cboPalette.AddItem "16 Colour"
   cboPalette.AddItem "Halftone"
   cboPalette.AddItem "Websafe"
   cboPalette.AddItem "Custom"
   cboPalette.ListIndex = 3
   
   ' Image source options:
   cboImageSource.AddItem "Bitmaps in a Directory"
   cboImageSource.AddItem "Picture Strip"
   cboImageSource.ListIndex = 0

   ' Mem DC for drawing:
   m_cDC.Create
   
   ' Initialise
   m_iWizardPanel = -1
   m_iPanelCount = 4
   ReDim m_sPanelCaption(1 To m_iPanelCount) As String
   m_sPanelCaption(1) = "Set Main AVI Options"
   m_sPanelCaption(2) = "Set Video and Colour Options"
   m_sPanelCaption(3) = "Set Source Images"
   m_sPanelCaption(4) = "Create Your AVI"
   wizardNavigate 1
   
End Sub

Private Sub pValidateRenderDirectoryImages()
End Sub

Private Sub pLoadDirectoryImages()
   
   txtDirectory.Tag = txtDirectory.Text
      
   ' Get the bitmaps in the directory:
   Dim sBaseDir As String
   Dim sFile As String
   Dim cb As cBmp
   Dim colBmp As New Collection
   
   sBaseDir = txtDirectory.Text
   If (Right(sBaseDir, 1) <> "\") Then sBaseDir = sBaseDir & "\"
   sFile = Dir(sBaseDir & "*.bmp")
   Do While Len(sFile) > 0
      If (GetAttr(sBaseDir & sFile) And vbDirectory) = 0 Then
         Set cb = New cBmp
         cb.Load sBaseDir & sFile
         colBmp.Add cb
      End If
      sFile = Dir
   Loop
   
   Dim lWidth As Long
   Dim lHeight As Long
   Dim lOverallWidth As Long
   Dim lOverallHeight As Long
   Dim bFirstTime As Boolean
   Dim bWarn As Boolean
   
   For Each cb In colBmp
      lOverallHeight = lOverallHeight + cb.Height
      If (bFirstTime) Then
         lWidth = cb.Width
         lHeight = cb.Height
         lOverallWidth = lWidth
      ElseIf (cb.Width <> lWidth) Or (cb.Height <> lHeight) Then
         bWarn = True
         If (cb.Width > lOverallWidth) Then
            lOverallWidth = cb.Width
         End If
      End If
   Next
   
   If bWarn Then
      'MsgBox "Warning: the images in this directory have different sizes.", vbInformation
   End If
   
   m_cBmp.Create lOverallWidth, lOverallHeight
   m_cDC.SelectObject m_cBmp
   
   Dim cDC As New cMemDC
   Dim y As Long
   cDC.Create
   For Each cb In colBmp
      cDC.SelectObject cb
      cDC.PaintPicture m_cDC.hdc, , y
      cDC.UnselectObject
      y = y + cb.Height
   Next
   
   m_cDC.UnselectObject
   
   pRenderDirectoryImages
   
End Sub

Private Sub pRenderDirectoryImages()
   
   picDirectoryPreview.Cls
      
   Dim lWidth As Long
   Dim lHeight As Long
   Dim fScale As Double
   
   lWidth = m_cBmp.Width
   lHeight = m_cBmp.Height
   If (lWidth > picStripPreview.ScaleWidth) Then
      fScale = picStripPreview.ScaleWidth / (lWidth * 1#)
      lWidth = picStripPreview.ScaleWidth
      lHeight = lHeight * fScale
   End If
      
   If (lHeight > picStripPreview.ScaleHeight) Then
      fScale = picStripPreview.ScaleHeight / (lHeight * 1#)
      lHeight = picStripPreview.ScaleHeight
      lWidth = lWidth * fScale
   End If

   m_cDC.SelectObject m_cBmp
   m_cDC.PaintPicture picDirectoryPreview.hdc, lWidth:=lWidth, lHeight:=lHeight
   m_cDC.UnselectObject
   
   picDirectoryPreview.Refresh

End Sub

Private Sub pValidateLoadBitmapStrip()
   'If Not (StrComp(txtBitmapStrip.Text, txtBitmapStrip.Tag) = 0) Then
      'If (fileExists(txtBitmapStrip.Text)) Then
         'pLoadBitmapStrip
      'End If
   'End If
End Sub
Private Sub pLoadBitmapStrip()
    txtBitmapStrip.Text = App.Path & "\Images\"
   m_cBmp.Load txtBitmapStrip.Text
   txtBitmapStrip.Tag = txtBitmapStrip.Text
   pRenderBitmapStrip
End Sub
Private Sub pRenderBitmapStrip()
   
   picStripPreview.Cls
      
   Dim lWidth As Long
   Dim lHeight As Long
   Dim fScale As Double
   
   lWidth = m_cBmp.Width
   lHeight = m_cBmp.Height
   If (lWidth > picStripPreview.ScaleWidth) Then
      fScale = picStripPreview.ScaleWidth / (lWidth * 1#)
      lWidth = picStripPreview.ScaleWidth
      lHeight = lHeight * fScale
   End If
      
   If (lHeight > picStripPreview.ScaleHeight) Then
      fScale = picStripPreview.ScaleHeight / (lHeight * 1#)
      lHeight = picStripPreview.ScaleHeight
      lWidth = lWidth * fScale
   End If
   
   m_cDC.SelectObject m_cBmp
   m_cDC.PaintPicture picStripPreview.hdc, lWidth:=lWidth, lHeight:=lHeight
   m_cDC.UnselectObject
   
   Dim lXCells As Long
   Dim lYCells As Long
   Dim lXCell As Long
   Dim lYCell As Long
   Dim lGridWidth As Long
   Dim lGridHeight As Long
   Dim x As Long
   Dim y As Long
   
   On Error Resume Next
   lXCells = CLng(txtXCells.Text)
   lYCells = CLng(txtYCells.Text)
   
   If (lXCells > 0) And (lYCells > 0) Then
      ' Grid cells:
      lGridWidth = lWidth / lXCells
      lGridHeight = lHeight / lYCells
      
      picStripPreview.ForeColor = &H0&
      x = 0
      y = 0
      For lXCell = 0 To lXCells
         picStripPreview.Line (x, y)-(x, y + lHeight)
         x = x + lGridWidth
      Next lXCell
      x = 0
      y = 0
      For lYCell = 0 To lHeight
         picStripPreview.Line (x, y)-(x + lWidth, y)
         y = y + lGridHeight
      Next lYCell
   End If
   
   
   picStripPreview.Refresh
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmAVICreator
Set frmAVICreator = Nothing
End Sub

Private Sub txtDirectory_KeyPress(KeyAscii As Integer)
   If (KeyAscii = 13) Then
      pValidateRenderDirectoryImages
   End If
End Sub

Private Sub txtDirectory_LostFocus()
   pValidateRenderDirectoryImages
End Sub

Private Sub txtXCells_Change()
   pRenderBitmapStrip
End Sub

Private Sub txtYCells_Change()
   pRenderBitmapStrip
End Sub
