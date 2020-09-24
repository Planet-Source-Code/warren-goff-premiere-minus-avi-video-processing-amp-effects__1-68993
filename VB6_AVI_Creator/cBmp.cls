VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" ( _
    ByVal lpDriverName As String, lpDeviceName As Any, _
   lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const IMAGE_BITMAP = 0

Private m_hBmp As Long
Private m_lWidth As Long
Private m_lHeight As Long

Public Property Get Width() As Long
   Width = m_lWidth
End Property
Public Property Get Height() As Long
   Height = m_lHeight
End Property

Public Property Get hBmp() As Long
   hBmp = m_hBmp
End Property

Public Sub Load(ByVal sFile As String)
Dim hBmpLoad As Long
   
   hBmpLoad = LoadImage(App.hInstance, sFile, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
   Init hBmpLoad
   
End Sub

Public Sub Create(ByVal lWidth As Long, ByVal lHeight As Long)
Dim lhDC As Long
   Dispose
   lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   m_hBmp = CreateCompatibleBitmap(lhDC, lWidth, lHeight)
   m_lWidth = lWidth
   m_lHeight = lHeight
   DeleteDC lhDC
End Sub

Public Sub Init(ByVal hBmp As Long)
   Dispose
   m_hBmp = hBmp
   Dim tBM As BITMAP
   GetObjectAPI m_hBmp, LenB(tBM), tBM
   m_lWidth = tBM.bmWidth
   m_lHeight = tBM.bmHeight
End Sub

Public Function ExtracthBmp() As Long
   ExtracthBmp = m_hBmp
   m_hBmp = 0
End Function

Public Sub Dispose()
   If Not (m_hBmp = 0) Then
      DeleteObject m_hBmp
   End If
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub
