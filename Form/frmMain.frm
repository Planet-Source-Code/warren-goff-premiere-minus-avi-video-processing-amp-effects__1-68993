VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8475
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   8475
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8010
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   15
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8040
      Picture         =   "frmMain.frx":09C4
      ScaleHeight     =   480
      ScaleWidth      =   420
      TabIndex        =   4
      ToolTipText     =   "Help"
      Top             =   45
      Width           =   420
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Resume Processing BMPs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2445
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   645
      Width           =   3045
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6195
      TabIndex        =   1
      Text            =   "20"
      Top             =   90
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start Processing AVIs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   45
      Width           =   3045
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8475
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8475
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Save Extracted BMP every              th frame"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3285
      TabIndex        =   2
      Top             =   105
      Width           =   5010
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Split = False
        Me.Caption = "Status: Deleting Old Frames!"
        Load AVI2BMP
        AVI2BMP.Show
        SetTopMostWindow AVI2BMP.hWnd, True
        If AVIInput = True Then AVI2BMP.cmdOpenAVIFile_Click
End Sub

Private Sub Command2_Click()
Load BMP2AVI
BMP2AVI.Show
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = False
Picture4.Visible = True

End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.Visible = True
Picture4.Visible = False
OpenBrowser App.Path & "\Help\Premiere-Minus.htm", Me.hWnd

End Sub
Public Function OpenBrowser(strURL As String, lngHwnd As Long)
    OpenBrowser = ShellExecute(lngHwnd, "", strURL, "", _
    "c:\", 10)
End Function

Private Sub Form_Load()
On Error Resume Next
    Split = False
    SetTopMostWindow Me.hWnd, True
    Initt = True
    If Command$ = "" Then
        AVIInput = False
    Else
        AVIInput = True
    End If
    MkDir App.Path & "\Images"
    MkDir App.Path & "\Images1"
    MkDir App.Path & "\SavedBMPs"
    If Dir(App.Path & "\GraphicalDLL.dll") = "" Then
        MsgBox "You must compile the GraphicalDLL.dll file and " & vbCrLf & "place it in the application directory!"
    End If
    If Dir(App.Path & "\AVICreator.exe") = "" Then
        MsgBox "You must compile the AVICreator.exe file and " & vbCrLf & "place it in the application directory!"
    End If
    If Dir(App.Path & "\AVItoGIF.exe") = "" Then
        MsgBox "You must compile the AVItoGIF.exe file and " & vbCrLf & "place it in the application directory!"
    End If
If AVIInput = True Then
    'Load AVI2BMP
    Text1.Text = 1
    AVI2BMP.List1.AddItem App.Path & "\" & Command$
    Command1_Click
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = 1 Then
   If Hidder = True Then
        AVI2BMP.Hide
   Else
        BMP2AVI.Hide
   End If
Else
    If Initt = False Then
        If Hidder = True Then
            AVI2BMP.Show
        Else
            BMP2AVI.Show
        End If
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim ret As Long
ret = AVISaveCallback(100)
Call AVIFileExit   '// releases AVIFile library
CloseAll

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

Private Sub Text1_Change()
If Val(Text1.Text) <= 0 Then Text1.Text = 1
End Sub
