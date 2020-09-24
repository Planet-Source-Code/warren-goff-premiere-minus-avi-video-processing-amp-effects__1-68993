VERSION 5.00
Begin VB.Form BMP2AVI 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Write AVI"
   ClientHeight    =   4665
   ClientLeft      =   1110
   ClientTop       =   2280
   ClientWidth     =   9720
   DrawMode        =   6  'Mask Pen Not
   FillColor       =   &H00E0E0E0&
   FillStyle       =   6  'Cross
   Icon            =   "BMP2AVI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save Selected"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1500
      Width           =   1785
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Make Anim GIF from AVI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1500
      Width           =   1605
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Restore"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6315
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1800
      Width           =   1440
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Backup"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6315
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1500
      Width           =   1440
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Write All Frames"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   150
      Width           =   1605
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000008&
      Caption         =   "Invert the Selection"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1875
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4215
      Width           =   2580
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000008&
      Caption         =   "Delete Selected Completely"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5385
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4215
      Width           =   2580
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3900
      Width           =   1245
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Deselect All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8445
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3900
      Width           =   1245
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4185
      TabIndex        =   14
      Top             =   5835
      Width           =   825
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   15
      TabIndex        =   12
      Tag             =   "6495"
      Top             =   2115
      Width           =   9690
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         LargeChange     =   5
         Left            =   45
         TabIndex        =   13
         Top             =   1380
         Width           =   9615
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00E0E0E0&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   6  'Cross
         Height          =   195
         Left            =   2595
         Top             =   1605
         Width           =   4545
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   1230
         Index           =   0
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   1230
         Index           =   5
         Left            =   8100
         Shape           =   4  'Rounded Rectangle
         Tag             =   "6495"
         Top             =   90
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   1230
         Index           =   4
         Left            =   6495
         Shape           =   4  'Rounded Rectangle
         Tag             =   "6495"
         Top             =   90
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   1230
         Index           =   3
         Left            =   4875
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   1230
         Index           =   2
         Left            =   3270
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   1230
         Index           =   1
         Left            =   1650
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Image FXImage 
         Height          =   1230
         Index           =   5
         Left            =   8100
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1560
      End
      Begin VB.Image FXImage 
         Height          =   1230
         Index           =   4
         Left            =   6495
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1560
      End
      Begin VB.Image FXImage 
         Height          =   1230
         Index           =   3
         Left            =   4890
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1560
      End
      Begin VB.Image FXImage 
         Height          =   1230
         Index           =   2
         Left            =   3270
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1560
      End
      Begin VB.Image FXImage 
         Height          =   1230
         Index           =   1
         Left            =   1650
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1560
      End
      Begin VB.Image FXImage 
         Height          =   1230
         Index           =   0
         Left            =   45
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1560
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   210
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   6000
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fx for all Frames"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   150
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fx for Selected"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7485
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   795
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Write Selected"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   795
      Width           =   1605
   End
   Begin VB.CommandButton cmdWriteAVI 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Write All Frames"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.ListBox lstDIBList 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000007&
      Height          =   645
      Left            =   420
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   5115
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.CommandButton cmdFileOpen 
      Caption         =   "Add BMP file to list..."
      Height          =   480
      Left            =   10080
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Clear file list"
      Height          =   480
      Left            =   10365
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.TextBox txtFPS 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Text            =   "1"
      Top             =   4620
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "to AVI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3150
      TabIndex        =   25
      Top             =   600
      Width           =   885
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   3675
      Picture         =   "BMP2AVI.frx":08CA
      Stretch         =   -1  'True
      ToolTipText     =   "Right Click Select to"
      Top             =   1665
      Width           =   450
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   135
      Picture         =   "BMP2AVI.frx":1097
      Stretch         =   -1  'True
      ToolTipText     =   "Left Click Select (From)"
      Top             =   1665
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      Index           =   4
      X1              =   414
      X2              =   652
      Y1              =   95
      Y2              =   95
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      Index           =   2
      X1              =   0
      X2              =   280
      Y1              =   95
      Y2              =   95
   End
   Begin VB.Image imgPreview 
      BorderStyle     =   1  'Fixed Single
      Height          =   2010
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   105
      Width           =   2010
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      Index           =   1
      X1              =   414
      X2              =   652
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Label lblStatus 
      Height          =   195
      Left            =   9720
      TabIndex        =   2
      Top             =   1035
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      Index           =   0
      X1              =   0
      X2              =   280
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You may Select the frames to process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   0
      Left            =   2895
      TabIndex        =   7
      Top             =   3825
      Width           =   4125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You may Select the frames to process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Index           =   1
      Left            =   2925
      TabIndex        =   15
      Top             =   3840
      Width           =   4125
   End
   Begin VB.Label lblfps 
      BackStyle       =   0  'Transparent
      Caption         =   "Frames per second             (1 - 30)"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   -120
      TabIndex        =   5
      Top             =   4635
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      Index           =   3
      X1              =   81
      X2              =   591
      Y1              =   274
      Y2              =   274
   End
End
Attribute VB_Name = "BMP2AVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ig As Long

'simple UDT containing parameters of first BMP file user chooses
'all the following BMPs should be the same format so there will be no problems in writing the vidstream
Private Type PARAMS
    Init As Boolean
    Width As Long
    Height As Long
    bpp As Long
End Type

Private Declare Function SetRect Lib "user32.dll" _
    (ByRef lprc As AVI_RECT, ByVal xLeft As Long, ByVal yTop As Long, ByVal xRight As Long, ByVal yBottom As Long) As Long 'BOOL

Private m_params As PARAMS

Dim ImageIndex As Long
Dim ListSel1 As Long, ListSel2 As Long

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long, intSave As String
On Error Resume Next
    Check1.Value = False
    intSave = MsgBox("You are about to DELETE the selected files from the disk. " & _
     vbCrLf & "Proceed?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intSave
      Case vbYes
            For i = 0 To lstDIBList.ListCount - 1
                If lstDIBList.Selected(i) = True Then
                    Kill lstDIBList.List(i)
                End If
            Next
            File1.Refresh
            lstDIBList.Clear
            Combo1.Clear
            For i = 0 To File1.ListCount - 1
                lstDIBList.AddItem App.Path & "\Images1\" & File1.List(i)
                Combo1.AddItem App.Path & "\Images1\" & File1.List(i)
            Next
            FXImage(0).Picture = LoadPicture(lstDIBList.List(0))
            FXImage(0).Refresh
            FXImage(1).Picture = LoadPicture(lstDIBList.List(1))
            FXImage(1).Refresh
            FXImage(2).Picture = LoadPicture(lstDIBList.List(2))
            FXImage(2).Refresh
            FXImage(3).Picture = LoadPicture(lstDIBList.List(3))
            FXImage(3).Refresh
            FXImage(4).Picture = LoadPicture(lstDIBList.List(4))
            FXImage(4).Refresh
            FXImage(5).Picture = LoadPicture(lstDIBList.List(5))
            FXImage(5).Refresh
            HScroll1.Min = 0
            HScroll1.Max = Combo1.ListCount - 1
      Case vbNo
        Exit Sub
      Case vbCancel
        Exit Sub
    End Select


End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Check1.Value = False
Dim i As Long
lstDIBList.Refresh
For i = 0 To lstDIBList.ListCount - 1
    If lstDIBList.Selected(i) = True Then
        lstDIBList.Selected(i) = False
    Else
        lstDIBList.Selected(i) = True
    End If
Next
lstDIBList.Refresh
For i = 0 To 5
        If Shape2(i).Visible = False Then
            Shape2(i).Visible = True
        Else
            Shape2(i).Visible = False
        End If
Next

End Sub

Private Sub Combo1_Click()
On Error Resume Next
Dim iii As Integer
    imgPreview.Picture = LoadPicture(Combo1.List(Combo1.ListIndex))
    imgPreview.Refresh
If Trim(Combo1.List(Combo1.ListIndex)) <> "" Then
    FXImage(0).Picture = LoadPicture(Combo1.List(Combo1.ListIndex))
    FXImage(0).Refresh: ig = Combo1.ListIndex
Else
    FXImage(0).Picture = LoadPicture("")
    FXImage(0).Refresh: ig = Combo1.ListIndex
End If
If Trim(Combo1.List(Combo1.ListIndex + 1)) <> "" Then
    FXImage(1).Picture = LoadPicture(Combo1.List(Combo1.ListIndex + 1))
    FXImage(1).Refresh
Else
    FXImage(1).Picture = LoadPicture("")
    FXImage(1).Refresh
End If
If Trim(Combo1.List(Combo1.ListIndex) + 2) <> "" Then
    FXImage(2).Picture = LoadPicture(Combo1.List(Combo1.ListIndex + 2))
    FXImage(2).Refresh
Else
    FXImage(2).Picture = LoadPicture("")
    FXImage(2).Refresh
End If
If Trim(Combo1.List(Combo1.ListIndex + 3)) <> "" Then
    FXImage(3).Picture = LoadPicture(Combo1.List(Combo1.ListIndex + 3))
    FXImage(3).Refresh
Else
    FXImage(3).Picture = LoadPicture("")
    FXImage(3).Refresh
End If
If Trim(Combo1.List(Combo1.ListIndex + 4)) <> "" Then
    FXImage(4).Picture = LoadPicture(Combo1.List(Combo1.ListIndex + 4))
    FXImage(4).Refresh
Else
    FXImage(4).Picture = LoadPicture("")
    FXImage(4).Refresh
End If
If Trim(Combo1.List(Combo1.ListIndex + 5)) <> "" Then
    FXImage(5).Picture = LoadPicture(Combo1.List(Combo1.ListIndex + 5))
    FXImage(5).Refresh
Else
    FXImage(5).Picture = LoadPicture("")
    FXImage(5).Refresh
End If
HScroll1.Value = Combo1.ListIndex
lstDIBList.ListIndex = Combo1.ListIndex

If Trim(Combo1.List(Combo1.ListIndex)) <> "" Then
    If lstDIBList.Selected(lstDIBList.ListIndex) = True Then
        Shape2(0).Visible = True
    Else
        Shape2(0).Visible = False
    End If
Else
    Exit Sub
End If
If Trim(Combo1.List(Combo1.ListIndex + 1)) <> "" Then
    If lstDIBList.Selected(lstDIBList.ListIndex + 1) = True Then
        Shape2(1).Visible = True
    Else
        Shape2(1).Visible = False
    End If
Else
    Exit Sub
End If
If Trim(Combo1.List(Combo1.ListIndex + 2)) <> "" Then
    If lstDIBList.Selected(lstDIBList.ListIndex + 2) = True Then
        Shape2(2).Visible = True
    Else
        Shape2(2).Visible = False
    End If
Else
    Exit Sub
End If
If Trim(Combo1.List(Combo1.ListIndex + 3)) <> "" Then
    If lstDIBList.Selected(lstDIBList.ListIndex + 3) = True Then
        Shape2(3).Visible = True
    Else
        Shape2(3).Visible = False
    End If
Else
    Exit Sub
End If
If Trim(Combo1.List(Combo1.ListIndex + 4)) <> "" Then
    If lstDIBList.Selected(lstDIBList.ListIndex + 4) = True Then
        Shape2(4).Visible = True
    Else
        Shape2(4).Visible = False
    End If
Else
    Exit Sub
End If
If Trim(Combo1.List(Combo1.ListIndex + 5)) <> "" Then
    If lstDIBList.Selected(lstDIBList.ListIndex + 5) = True Then
        Shape2(5).Visible = True
    Else
        Shape2(5).Visible = False
    End If
Else
    Exit Sub
End If

End Sub

Private Sub Combo2_Click()
On Error Resume Next
'HScroll1.SmallChange = Val(Combo2.Text)
HScroll1.SetFocus

End Sub

Private Sub Combo3_Click()
On Error Resume Next
'HScroll1.SmallChange = Val(Combo3.Text)
HScroll1.SetFocus

End Sub

Private Sub Command1_Click()
On Error GoTo error
Set AVI2BMP = Nothing
'Me.Enabled = False
    Dim file As cFileDlg
    Dim InitDir As String
    Dim szOutputAVIFile As String
    Dim res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim bmp As cDIB
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim bi As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim i As Long, j As Long, ii As Long
    Dim intSave As Integer
    If lstDIBList.SelCount = 0 Then
            intSave = MsgBox("You haven't selected any frames. " & vbCrLf _
            & "Do you want to process all?", _
                             vbYesNoCancel + vbExclamation)
            Select Case intSave
            Case vbYes
                cmdWriteAVI_Click
            Case vbNo
                Exit Sub
            Case vbCancel
                Exit Sub
            End Select
    End If
            intSave = MsgBox("The unselected BMPs will be deleted. " & vbCrLf _
            & "Do you want to proceed?", _
                             vbYesNoCancel + vbExclamation)
            Select Case intSave
            Case vbYes
                For i = 0 To lstDIBList.ListCount - 1
                    If lstDIBList.Selected(i) = False Then
                        Kill lstDIBList.List(i)
                    End If
                Next
                File1.Refresh
                lstDIBList.Clear
                Combo1.Clear
            For i = 0 To File1.ListCount - 1
                lstDIBList.AddItem App.Path & "\Images1\" & File1.List(i)
                Combo1.AddItem App.Path & "\Images1\" & File1.List(i)
            Next
                FXImage(0).Picture = LoadPicture(lstDIBList.List(0))
                FXImage(0).Refresh
                FXImage(1).Picture = LoadPicture(lstDIBList.List(1))
                FXImage(1).Refresh
                FXImage(2).Picture = LoadPicture(lstDIBList.List(2))
                FXImage(2).Refresh
                FXImage(3).Picture = LoadPicture(lstDIBList.List(3))
                FXImage(3).Refresh
                FXImage(4).Picture = LoadPicture(lstDIBList.List(4))
                FXImage(4).Refresh
                FXImage(5).Picture = LoadPicture(lstDIBList.List(5))
                FXImage(5).Refresh
                HScroll1.Min = 0
                HScroll1.Max = Combo1.ListCount - 1
            Case vbNo
                Exit Sub
            Case vbCancel
                Exit Sub
            End Select
   'Debug.Print
    'Set file = New cFileDlg
    'get an avi filename from user
    'With file
        '.DefaultExt = "avi"
        '.DlgTitle = "Choose a filename to save AVI to..."
        '.Filter = "AVI Files|*.avi"
        '.OwnerHwnd = Me.hWnd
    'End With
    Shell App.Path & "\AVICreator.exe", vbNormalFocus
Exit Sub
    szOutputAVIFile = App.Path & "\MyAVI.avi"
    'If file.VBGetSaveFileName(szOutputAVIFile) <> True Then Exit Sub
        
'    Open the file for writing
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error

    'Get the first bmp in the list for setting format
    Set bmp = New cDIB
    lstDIBList.ListIndex = ListSel1
    If bmp.CreateFromFile(lstDIBList.Text) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
        GoTo error
    End If

'   Fill in the header for the video stream
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&                             '// default AVI handler
        .dwScale = 1
        .dwRate = Val(txtFPS)                        '// fps
        .dwSuggestedBufferSize = bmp.SizeImage       '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)       '// rectangle for stream
    End With
    
    'validate user input
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

'   And create the stream
    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error

    'get the compression options from the user
    'Careful! this API requires a pointer to a pointer to a UDT
    pOpts = VarPtr(opts)
    res = AVISaveOptions(Me.hWnd, _
                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
                        1, _
                        ps, _
                        pOpts) 'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then 'In C TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo error
    End If
    
    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error
    
    'set format of stream according to the bitmap
    With bi
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

'   Now write out each video frame
    For i = 0 To lstDIBList.ListCount - 1
        If lstDIBList.Selected(i) = True Then
            lstDIBList.ListIndex = i
            bmp.CreateFromFile (lstDIBList.Text) 'load the bitmap (ignore errors)
            res = AVIStreamWrite(psCompressed, _
                                i, _
                                1, _
                                bmp.PointerToBits, _
                                bmp.SizeImage, _
                                AVIIF_KEYFRAME, _
                                ByVal 0&, _
                                ByVal 0&)
            If res <> AVIERR_OK Then GoTo error
            'Show user feedback
            imgPreview.Picture = LoadPicture(lstDIBList.Text)
            imgPreview.Refresh
            lblStatus = i & " saved"
            lblStatus.Refresh
        End If
    Next
    lblStatus = "Finished!"

error:
'   Now close the file
    Set file = Nothing
    Set bmp = Nothing
    
    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit
    If (res <> AVIERR_OK) Then
        MsgBox "There was an error writing the file.", vbInformation, App.Title
    End If
Unload Me

End Sub

Private Sub Command10_Click()
Dim i As Long
    For i = 0 To lstDIBList.ListCount - 1
        If lstDIBList.Selected(i) = True Then
            FileCopy lstDIBList.List(i), App.Path & "\SavedBMPs\" & Format(Now, "ddmmyyhhmmss") & "MyBMP.bmp"
        End If
    Next i
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Dim intSave As Integer
    If lstDIBList.SelCount = 0 Then
            intSave = MsgBox("You haven't selected any frames. " & vbCrLf _
            & "Do you want to process all?", _
                             vbYesNoCancel + vbExclamation)
            Select Case intSave
            Case vbYes
                Command3_Click
            Case vbNo
                Exit Sub
            Case vbCancel
                Exit Sub
            End Select
    End If
    ApplyFlag = True
    Load GraphicalDLL
    GraphicalDLL.Show

End Sub

Private Sub Command3_Click()
On Error Resume Next
    ApplyFlag = False
    Load GraphicalDLL
    GraphicalDLL.Show
End Sub

Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Set AVI2BMP = Nothing
SetTopMostWindow Me.hWnd, False
    Dim file As cFileDlg
    Dim InitDir As String
    Dim szOutputAVIFile As String
    Dim res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim bmp As cDIB
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim bi As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim i As Long, j As Long, ii As Long
    Dim intSave As Integer
    Shell App.Path & "\AVICreator.exe", vbNormalFocus
End Sub

Private Sub Command5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
File1.Path = App.Path & "\Images1\"
File1.Refresh
For i = 0 To File1.ListCount - 1
    FileCopy App.Path & "\Images1\" & File1.List(i), App.Path & "\Backup\" & File1.List(i)
Next
End Sub

Private Sub Command6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim i As Long
For i = 0 To lstDIBList.ListCount - 1
    lstDIBList.Selected(i) = False
Next
For i = 0 To 5
        Shape2(i).Visible = False
        lstDIBList.Selected(lstDIBList.ListIndex + i) = False
Next

End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim i As Long
For i = 0 To lstDIBList.ListCount - 1
    lstDIBList.Selected(i) = True
Next
For i = 0 To 5
        Shape2(i).Visible = True
        lstDIBList.Selected(lstDIBList.ListIndex + i) = True
Next

End Sub

Private Sub Command8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Kill App.Path & "\Images1\*.bmp"
Dim i As Long
File1.Path = App.Path & "\Backup\"
File1.Refresh
lstDIBList.Clear
Combo1.Clear
For i = 0 To File1.ListCount - 1
    FileCopy App.Path & "\Backup\" & File1.List(i), App.Path & "\Images1\" & File1.List(i)
    lstDIBList.AddItem App.Path & "\Images1\" & File1.List(i)
    Combo1.AddItem App.Path & "\Images1\" & File1.List(i)
Next
FXImage(0).Picture = LoadPicture(lstDIBList.List(0))
FXImage(0).Refresh
FXImage(1).Picture = LoadPicture(lstDIBList.List(1))
FXImage(1).Refresh
FXImage(2).Picture = LoadPicture(lstDIBList.List(2))
FXImage(2).Refresh
FXImage(3).Picture = LoadPicture(lstDIBList.List(3))
FXImage(3).Refresh
FXImage(4).Picture = LoadPicture(lstDIBList.List(4))
FXImage(4).Refresh
FXImage(5).Picture = LoadPicture(lstDIBList.List(5))
FXImage(5).Refresh
HScroll1.Min = 0
HScroll1.Max = Combo1.ListCount - 1

File1.Path = App.Path & "\Images1\"

End Sub

Private Sub Command9_Click()
    Shell App.Path & "\AVItoGIF.exe 1", vbNormalFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Long
ApplyFlag = False
Me.Show
Me.WindowState = 0
SetTopMostWindow Me.hWnd, True
File1.Path = App.Path & "\Images1\"
File1.Refresh
For i = 0 To File1.ListCount - 1
    lstDIBList.AddItem App.Path & "\Images1\" & File1.List(i)
    Combo1.AddItem App.Path & "\Images1\" & File1.List(i)
Next
    FXImage(0).Picture = LoadPicture(lstDIBList.List(0))
    FXImage(0).Refresh
    FXImage(1).Picture = LoadPicture(lstDIBList.List(1))
    FXImage(1).Refresh
    FXImage(2).Picture = LoadPicture(lstDIBList.List(2))
    FXImage(2).Refresh
    FXImage(3).Picture = LoadPicture(lstDIBList.List(3))
    FXImage(3).Refresh
    FXImage(4).Picture = LoadPicture(lstDIBList.List(4))
    FXImage(4).Refresh
    FXImage(5).Picture = LoadPicture(lstDIBList.List(5))
    FXImage(5).Refresh
HScroll1.Min = 0
HScroll1.Max = Combo1.ListCount - 1
   
   'Call AVIFileInit   '// opens AVIFile library
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Combo1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SetTopMostWindow Me.hWnd, False
Dim intSave As Integer
    intSave = MsgBox("Do you want to Delete used BMPs?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intSave
      Case vbYes
        On Error Resume Next
        Kill App.Path & "\Images\*.bmp"
        Kill App.Path & "\Images1\*.bmp"
        
    Case vbNo
    Case vbCancel
        Exit Sub
    End Select
    intSave = MsgBox("View the Video File?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intSave
      Case vbYes
        Dim ngReturnNumber As Long
        ngReturnNumber = ShellExecLaunchFile(App.Path & "\MyAVI.avi", "", App.Path)
    End Select
    
    Call AVIFileExit   '// releases AVIFile library
    Set AVI2BMP = Nothing
    'CloseAll
End Sub
Sub CloseAll()
    On Error Resume Next
    Dim intFrmNum As Integer
    intFrmNum = Forms.Count


    Do Until intFrmNum = 0
        Unload Forms(intFrmNum - 1)
        Set Forms(intFrmNum - 1) = Nothing
        intFrmNum = intFrmNum - 1
    Loop
End Sub
Private Sub cmdFileOpen_Click()
'adds a bmp to list of files to create video stream from
Dim szFileName As String
Dim file As cFileDlg
Dim bmp As cDIB
Static InitDir As String
On Error Resume Next

'Set file dialog parameters
Set file = New cFileDlg
With file
    .DlgTitle = "Choose BMP file to add to video stream"
    .Filter = "BMP Files|*.bmp:DIB Files|*.dib"
    If InitDir <> "" Then
        file.InitDirectory = InitDir
    End If
End With

'get filename from user
If file.VBGetOpenFileName(szFileName) = True Then
    Set bmp = New cDIB
    If bmp.CreateFromFile(szFileName) Then 'file is a valid BMP
        If m_params.Init Then 'this is not the first file - it must be the same format
            If (bmp.Width <> m_params.Width) _
                Or (bmp.Height <> m_params.Height) _
                Or (bmp.BitCount <> m_params.bpp) Then
                MsgBox "Chosen bitmap file is a different format!", vbInformation, App.Title 'format is wrong
            Else
                imgPreview.Picture = LoadPicture(szFileName) 'format is OK -add file to list
                lstDIBList.AddItem szFileName
            End If
        Else 'this is the first file in the list so save format info too
            With m_params
                .Init = True
                .Width = bmp.Width
                .Height = bmp.Height
                .bpp = bmp.BitCount
            End With
            imgPreview.Picture = LoadPicture(szFileName)
            lstDIBList.AddItem szFileName
        End If
        cmdClearList.Enabled = True 'make sure clear button is enabled
        cmdWriteAVI.Enabled = True 'allow user to call AVI write functions
    End If
    Set bmp = Nothing
End If
'save last directory for user
InitDir = file.InitDirectory
Set file = Nothing
End Sub

Private Sub cmdClearList_Click()
On Error Resume Next
'reset file list - unload picture - reset format params
lstDIBList.Clear
imgPreview.Picture = LoadPicture()
With m_params
    .bpp = 0
    .Height = 0
    .Width = 0
    .Init = False
End With
cmdClearList.Enabled = False
cmdWriteAVI.Enabled = False
End Sub

Public Sub cmdWriteAVI_Click()
On Error GoTo error
Set AVI2BMP = Nothing
SetTopMostWindow Me.hWnd, False
'Me.Enabled = False
    Dim file As cFileDlg
    Dim InitDir As String
    Dim szOutputAVIFile As String
    Dim res As Long
    Dim pfile As Long 'ptr PAVIFILE
    Dim bmp As cDIB
    Dim ps As Long 'ptr PAVISTREAM
    Dim psCompressed As Long 'ptr PAVISTREAM
    Dim strhdr As AVI_STREAM_INFO
    Dim bi As BITMAPINFOHEADER
    Dim opts As AVI_COMPRESS_OPTIONS
    Dim pOpts As Long
    Dim i As Long, j As Long, ii As Long
    Dim intSave As Integer
    szOutputAVIFile = App.Path & "\MyAVI.avi"
    res = AVIFileOpen(pfile, szOutputAVIFile, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error
    Set bmp = New cDIB
    lstDIBList.ListIndex = ListSel1
    If bmp.CreateFromFile(lstDIBList.Text) <> True Then
        MsgBox "Could not load first bitmap file in list!", vbExclamation, App.Title
        GoTo error
    End If

    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)    '// stream type video
        .fccHandler = 0&                             '// default AVI handler
        .dwScale = 1
        .dwRate = Val(txtFPS)                        '// fps
        .dwSuggestedBufferSize = bmp.SizeImage       '// size of one frame pixels
        Call SetRect(.rcFrame, 0, 0, bmp.Width, bmp.Height)       '// rectangle for stream
    End With
    
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30

    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error
    pOpts = VarPtr(opts)
    res = AVISaveOptions(Me.hWnd, _
                        ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, _
                        1, _
                        ps, _
                        pOpts) 'returns TRUE if User presses OK, FALSE if Cancel, or error code
    If res <> 1 Then 'In C TRUE = 1
        Call AVISaveOptionsFree(1, pOpts)
        GoTo error
    End If
    
    'make compressed stream
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error
    
    'set format of stream according to the bitmap
    With bi
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With
    
    'set the format of the compressed stream
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error

'   Now write out each video frame
    For i = 0 To lstDIBList.ListCount - 1
        lstDIBList.ListIndex = i
        bmp.CreateFromFile (lstDIBList.Text) 'load the bitmap (ignore errors)
        res = AVIStreamWrite(psCompressed, _
                            i, _
                            1, _
                            bmp.PointerToBits, _
                            bmp.SizeImage, _
                            AVIIF_KEYFRAME, _
                            ByVal 0&, _
                            ByVal 0&)
        If res <> AVIERR_OK Then GoTo error
        'Show user feedback
        imgPreview.Picture = LoadPicture(lstDIBList.Text)
        imgPreview.Refresh
        lblStatus.Caption = i & " saved"
        lblStatus.Refresh
    Next
    lblStatus.Caption = "Finished!"

error:
'   Now close the file
    Set file = Nothing
    Set bmp = Nothing
    
    If (ps <> 0) Then Call AVIStreamClose(ps)

    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)

    If (pfile <> 0) Then Call AVIFileClose(pfile)

    Call AVIFileExit
    If (res <> AVIERR_OK) Then
        MsgBox "There was an error writing the file.", vbInformation, App.Title
    End If
    Unload Me
End Sub

Private Sub FXImage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Combo1.SetFocus

End Sub

Private Sub FXImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim i As Long, j As Long
If lstDIBList.List(ig + Index) <> "" Then imgPreview.Picture = FXImage(Index).Picture
If Shape2(Index).Visible = False And lstDIBList.List(ig + Index) <> "" Then
    Shape2(Index).Visible = True
    lstDIBList.Selected(ig + Index) = True
    lstDIBList.Refresh
Else
    Shape2(Index).Visible = False
    lstDIBList.Selected(ig + Index) = False
    lstDIBList.Refresh
End If
'MsgBox ig   'lstDIBList.ListIndex
If lstDIBList.SelCount = 0 Then
    Command1.Enabled = False
    Command2.Enabled = False
Else
    Command1.Enabled = True
    Command2.Enabled = True
End If
If Button = vbLeftButton Then
    ImageIndex = ig + Index     ': MsgBox ImageIndex
End If
If Button = vbRightButton Then
    For i = ImageIndex + 1 To ig + Index
        lstDIBList.Selected(i) = True
        lstDIBList.Refresh
    Next
    HScroll1.Value = HScroll1.Value + 1
End If
End Sub

Private Sub HScroll1_Change()
On Error Resume Next
Combo1.ListIndex = HScroll1.Value
End Sub

Private Sub imgPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Combo1.SetFocus

End Sub

Private Sub lblStatus_Change()
On Error Resume Next
DoEvents
frmMain.Caption = lblStatus.Caption
End Sub

Public Sub lstDIBList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    imgPreview.Picture = LoadPicture(lstDIBList.Text)
    imgPreview.Refresh
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Combo2.SetFocus
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Combo3.SetFocus
End Sub
