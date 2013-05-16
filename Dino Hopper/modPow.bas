Attribute VB_Name = "modPow"
Public Sub getPowEffect(ByVal index As Integer, ByVal strType As String)
With frmMain
.tmrPow(index).Tag = strType
If strType = "Scare" Then
    If index = 0 Then
        'make enemies run from player
    End If
ElseIf strType = "Speed" Then
    .tmrChar(index).Interval = 25
    .tmrFrame(index).Interval = 25
    .tmrPow(index).Enabled = True
ElseIf strType = "Freeze" Then
End If
End With
End Sub
