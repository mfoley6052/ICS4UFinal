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

Public Sub getChangeTerrain(inputTile As terrain, strType As String, blnKill As Boolean)
If Not blnKill Then
    inputTile.hasObj = True
    inputTile.objType(0) = "Terrain"
    inputTile.objType(1) = strType
    inputTile.terType = inputTile.terType & strType
    Set inputTile.tempPic = inputTile.picTile
    With frmMain
    If inputTile.cutoff = "" Then
        Set inputTile.picTile = .picTile(4)
    ElseIf inputTile.cutoff = "L" Then
        Set inputTile.picTile = .picTileL(4)
    ElseIf inputTile.cutoff = "R" Then
        Set inputTile.picTile = .picTileR(4)
    ElseIf inputTile.cutoff = "LR" Then
        Set inputTile.picTile = .picTileLR(4)
    End If
    End With
    objTileCount = objTileCount + 1
    Call clearTile(inputTile, False, -1, "ObjXYF")
Else
    inputTile.terType = Mid(inputTile.terType, 1, 1)
    Set inputTile.picTile = inputTile.tempPic
    Call clearTile(inputTile, False, -1, "ObjXYF")
End If
End Sub
