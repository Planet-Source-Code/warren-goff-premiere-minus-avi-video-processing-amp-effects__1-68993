Attribute VB_Name = "modFunctions"
Public Effect As Integer

Public Sub PleaseDontResize(Formulario As Form, Largura As Long, Altura As Long)
    
    On Error Resume Next
    If Formulario.Width < Largura Or Formulario.Width > Largura Then
        Formulario.Width = Largura
    End If
    If Formulario.Height < Altura Or Formulario.Height > Altura Then
        Formulario.Height = Altura
    End If
    
End Sub

Public Sub ChangeControls(EffectNumber As Integer, Button As CommandButton, _
    Scroll As HScrollBar, Optional Min, Optional Max, Optional Value)
On Error Resume Next
    Effect = EffectNumber
    Processor = EffectNumber
    ScrollMin = Min
    ScrollMax = Max
    If (IsMissing(Min)) Then
        Button.Enabled = True
        Scroll.Enabled = False
        ApplyEffect = True
    Else
        Button.Enabled = False
        Scroll.Enabled = True
        Scroll.Min = Min
        Scroll.Max = Max
        Scroll.Value = Value
        ApplyEffect = False
    End If
End Sub

