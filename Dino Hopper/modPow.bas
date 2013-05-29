Attribute VB_Name = "modPow"
Public Sub getPowEffect(ByVal index As Integer, ByVal strType As String)
With frmMain
If .tmrPow(index).Tag <> "" Then
    Call getPowExpire(index, .tmrPow(index).Tag)
End If
.tmrPow(index).Tag = strType
tmrPowCounter = 0
If gameMode <> 1 Then
    .tmrPow(index).Enabled = True
End If
If strType = "Scare" Then
    If index = 0 Then
        'make enemies run from player
    End If
ElseIf strType = "Speed" Then
    .tmrChar(index).Interval = 20
ElseIf strType = "Freeze" Then
    .tmrChar(index).Tag = strType
End If
End With
End Sub

Public Sub getPowExpire(ByVal index As Integer, ByVal strPow As String)
With frmMain
If .tmrPow(index).Tag = "Recover" Then
    blnRecover(index) = False
ElseIf .tmrPow(index).Tag = "Scare" Then
ElseIf .tmrPow(index).Tag = "Speed" Then
    .tmrChar(index).Interval = 40
ElseIf .tmrPow(index).Tag = "Freeze" Then
    .tmrChar(index).Tag = ""
End If
.tmrPow(index).Enabled = False
End With
End Sub

Public Sub getPowTick(ByVal index As Integer)
tmrPowCounter = tmrPowCounter + 1
With frmMain
If tmrPowCounter >= tmrPowLimit And .tmrChar(index).Tag <> "Recover" Then
    Call getPowExpire(index, .tmrPow(index).Tag)
    tmrPowCounter = 0
    .tmrPow(index).Tag = ""
    .tmrPow(index).Enabled = False
ElseIf .tmrChar(index).Tag = "Recover" And tmrPowCounter >= tmrPowLimit * 0.75 Then
    Call getPowExpire(index, .tmrPow(index).Tag)
    tmrPowCounter = 0
    .tmrPow(index).Tag = ""
    .tmrPow(index).Enabled = False
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
    Dim intType As Integer
    If strType = "I" Then
        intType = 4
    End If
    With frmMain
    If inputTile.cutoff = "" Then
        Set inputTile.picTile = .picTile(intType)
    ElseIf inputTile.cutoff = "L" Then
        Set inputTile.picTile = .picTileL(intType)
    ElseIf inputTile.cutoff = "R" Then
        Set inputTile.picTile = .picTileR(intType)
    ElseIf inputTile.cutoff = "LR" Then
        Set inputTile.picTile = .picTileLR(intType)
    End If
    End With
    objTileCount = objTileCount + 1
Else
    inputTile.terType = Mid(inputTile.terType, 1, 1)
    Set inputTile.picTile = inputTile.tempPic
End If
End Sub
