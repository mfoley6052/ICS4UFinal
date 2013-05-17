Attribute VB_Name = "modPow"
Public Sub getPowEffect(ByVal index As Integer, ByVal strType As String)
With frmMain
If .tmrPow(index).Tag <> "" Then
    Call getPowExpire(index, .tmrPow(index).Tag)
End If
.tmrPow(index).Tag = strType
tmrPowCounter = 0
.tmrPow(index).Enabled = True
If strType = "Scare" Then
    If index = 0 Then
        'make enemies run from player
    End If
ElseIf strType = "Speed" Then
    .tmrChar(index).Interval = 20
    .tmrFrame(index).Interval = 20
ElseIf strType = "Freeze" Then
    .tmrChar(index).Tag = strType
End If
End With
End Sub

Public Sub getPowExpire(ByVal index As Integer, ByVal strPow As String)
With frmMain
If .tmrPow(index).Tag = "Scare" Then
ElseIf .tmrPow(index).Tag = "Speed" Then
    .tmrChar(index).Interval = 40
    .tmrFrame(index).Interval = 40
    .tmrPow(index).Enabled = False
ElseIf .tmrPow(index).Tag = "Freeze" Then
    .tmrChar(index).Tag = ""
End If
End With
End Sub
