Attribute VB_Name = "modPow"
Public Sub getPowEffect(ByVal index As Integer, ByVal strType As String)
With frmMain
If .tmrPow(index).Tag <> "" Then
    Call getPowExpire(index, .tmrPow(index).Tag)
End If
.tmrPow(index).Tag = strType
tmrPowCounter(index) = 0
If gameMode <> 1 Then
    .tmrPow(index).Enabled = True
End If
If strType = "Scare" Then
    If index = 0 Then
        'make enemies run from player
    End If
ElseIf strType = "Speed" Then
    If gameMode <> 1 Then
        .tmrChar(index).Interval = 20
        If index > 0 Then
            .tmrCPUMove(index).Interval = 150
        End If
    Else
        intMoves(index) = 2
    End If
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
    If gameMode <> 1 Then
        .tmrChar(index).Interval = 40
        If index > 0 Then
            .tmrCPUMove(index).Interval = 300
        End If
    Else
        intMoves(index) = 1
    End If
ElseIf .tmrPow(index).Tag = "Freeze" Then
    .tmrChar(index).Tag = ""
End If
.tmrPow(index).Enabled = False
End With
End Sub

Public Sub getPowTick(ByVal index As Integer)
tmrPowCounter(index) = tmrPowCounter(index) + 1
With frmMain
If tmrPowCounter(index) >= tmrPowLimit And .tmrChar(index).Tag <> "Recover" Then
    Call getPowExpire(index, .tmrPow(index).Tag)
    tmrPowCounter(index) = 0
    .tmrPow(index).Tag = ""
    .tmrPow(index).Enabled = False
ElseIf .tmrChar(index).Tag = "Recover" And tmrPowCounter(index) >= tmrPowLimit * 0.75 Then
    Call getPowExpire(index, .tmrPow(index).Tag)
    tmrPowCounter(index) = 0
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
tile(inputTile.Xc, inputTile.Yc) = inputTile
End Sub
