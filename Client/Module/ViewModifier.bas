Attribute VB_Name = "ViewModifier"
Public Sub ObjectViewResizeHorizontal(Form As Form, List As Control, DistanceFromRight As Integer)
    List.Width = Form.ScaleWidth - List.Left - 30 - DistanceFromRight
End Sub

Public Sub ObjectViewResizeVertical(Form As Form, List As Control, DistanceFormBottom As Integer)
    List.Height = Form.ScaleHeight - List.Top - DistanceFormBottom - 30
End Sub

Public Sub ObjectAlignRight(Form As Form, Object As Control, DistanceFormRight As Integer)
    Object.Left = Form.ScaleWidth - Object.Width - 30 - DistanceFormRight
End Sub

Public Sub ObjectAlignBottom(Form As Form, Object As Control, DistanceFormBottom As Integer)
    Object.Top = Form.ScaleHeight - Object.Height - DistanceFormBottom
End Sub

Public Sub ObjectAlignLeft(Object As Control, DistanceFormLeft As Integer)
    Object.Left = DistanceFormLeft + 30
End Sub

Public Sub ObjectAlignObjectHorizontal(MeControl As Control, Target As Control, Distance As Integer)
    MeControl.Left = Target.Left - Distance - MeControl.Width
End Sub

Public Sub ObjectAlignObjectVertical(MeControl As Control, Target As Control, Distance As Integer)
    MeControl.Top = Target.Top - Distance - MeControl.Height
End Sub

Public Sub ObjectAlignCenterHorizontal(Form As Form, Object As Control)
    Object.Left = (Form.ScaleWidth / 2) - (Object.Width / 2) + 30
End Sub


Public Function MaxMinForm(Form As Form, MinHeight As Integer, MinWidth As Integer) As Boolean
    MaxMinForm = False
    If (Form.ScaleHeight > MinHeight) Or (Form.ScaleWidth > MinWidth) Then
        MaxMinForm = True
    End If
End Function

Public Sub HoldFormScale(Form As Form, Height As Integer, Width As Integer)
    Form.Width = Width + 240
    Form.Height = Height + 550
End Sub

Public Sub HoldFormScaleWidth(Form As Form, Width As Integer)
    Form.Width = Width + 240 + 50
End Sub

Public Sub HoldFormScaleHeight(Form As Form, Height As Integer)
    Form.Height = Height + 550 + 50
End Sub

Public Sub FitInScreen(Form As Form, Object As Control)
    Object.Height = Form.ScaleHeight
    Object.Width = Form.ScaleWidth
End Sub

