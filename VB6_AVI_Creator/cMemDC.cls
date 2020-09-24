VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
   Left As Long
   TOp As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObjectAPI Lib "gdi32" Alias "SelectObject" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" ( _
    ByVal lpDriverName As String, lpDeviceName As Any, _
   lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private m_hBmpOld As Long
Private m_hDC As Long
Private m_cBmp As cBmp

Public Sub Create()
Dim lhDC As Long
   Dispose
   lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   m_hDC = CreateCompatibleDC(lhDC)
   DeleteDC lhDC
End Sub

Public Sub Dispose()
   UnselectObject
   If Not (m_hDC = 0) Then
      DeleteDC m_hDC
      m_hDC = 0
   End If
End Sub

Public Sub SelectObject(cb As cBmp)
   UnselectObject
   m_hBmpOld = SelectObjectAPI(m_hDC, cb.hBmp)
   Set m_cBmp = cb
End Sub

Public Sub UnselectObject()
   If Not (m_hBmpOld = 0) Then
      SelectObjectAPI m_hDC, m_hBmpOld
      m_hBmpOld = 0
      Set m_cBmp = Nothing
   End If
End Sub

Public Sub Fill(ByVal oColor As OLE_COLOR)
   If Not m_cBmp Is Nothing Then
      Dim tR As RECT
      tR.Right = m_cBmp.Width
      tR.Bottom = m_cBmp.Height
      Dim hBr As Long
      hBr = CreateSolidBrush(TranslateColor(oColor))
      FillRect m_hDC, tR, hBr
      DeleteObject hBr
   Else
      Err.Raise 7, App.EXEName & ".cMemDC", "Cannot fill DC as no bitmap is selected."
   End If
End Sub

Public Sub PaintPicture( _
      ByVal lHDCTo As Long, _
      Optional ByVal x As Long = 0, _
      Optional ByVal y As Long = 0, _
      Optional ByVal lWidth As Long = -1, _
      Optional ByVal lHeight As Long = -1, _
      Optional ByVal srcX As Long = 0, _
      Optional ByVal srcY As Long = 0, _
      Optional ByVal srcWidth As Long = -1, _
      Optional ByVal srcHeight As Long = -1 _
   )
   If Not m_cBmp Is Nothing Then
      If (lWidth < 0) Then lWidth = m_cBmp.Width
      If (lHeight < 0) Then lHeight = m_cBmp.Height
      If (srcWidth < 0) Then srcWidth = m_cBmp.Width
      If (srcHeight < 0) Then srcHeight = m_cBmp.Height
      If (srcWidth = lWidth) And (srcHeight = lHeight) Then
         BitBlt lHDCTo, x, y, lWidth, lHeight, m_hDC, srcX, srcY, vbSrcCopy
      Else
         StretchBlt lHDCTo, x, y, lWidth, lHeight, m_hDC, srcX, srcY, srcWidth, srcHeight, vbSrcCopy
      End If
   Else
      Err.Raise 7, App.EXEName & ".cMemDC", "Cannot fill DC as no bitmap is selected."
   End If
End Sub

Public Property Get hdc() As Long
   hdc = m_hDC
End Property

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = -1 'CLR_INVALID
    End If
End Function
Private Sub Class_Terminate()
   Dispose
End Sub
