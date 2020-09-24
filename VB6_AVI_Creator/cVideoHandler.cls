VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cVideoHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sName As String
Private m_sDescription As String
Private m_lFourCC As Long

Friend Sub fInit(ByVal sName As String, ByVal sDescription, ByVal lFourCC As Long)
   m_lFourCC = lFourCC
   m_sName = sName
   m_sDescription = sDescription
End Sub

Public Property Get Name() As String
   Name = m_sName
End Property
Public Property Get Description() As String
   Description = m_sDescription
End Property
Public Property Get FourCC() As Long
   FourCC = m_lFourCC
End Property
