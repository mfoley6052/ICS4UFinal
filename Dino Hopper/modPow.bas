Attribute VB_Name = "modPow"
'power-up effect
Public Sub getPowEffect(ByVal Index As Integer, ByVal strType As String)
With frmMain
'if character had another power-up, end that one
If .tmrPow(Index).Tag <> "" Then
    Call getPowExpire(Index, .tmrPow(Index).Tag)
End If
.tmrPow(Index).Tag = strType
tmrPowCounter(Index) = 0
If gameMode = 0 Then
    .tmrPow(Index).Enabled = True
End If
'set power-ups to do what they do; play sounds
If strType = "Scare" Then
    If Index = 0 Then
        isScared = True
        If canPlay Then
            frmMain.mmcPow(1).Command = "prev"
            frmMain.mmcPow(1).Command = "open"
            frmMain.mmcPow(1).Command = "play"
        End If
    End If
ElseIf strType = "Speed" Then
    If canPlay Then
        frmMain.mmcPow(0).Command = "prev"
        frmMain.mmcPow(0).Command = "open"
        frmMain.mmcPow(0).Command = "play"
    End If
    If gameMode = 0 Then
        .tmrChar(Index).Interval = 20
        If Index > 0 Then
            .tmrCPUMove(Index).Interval = 150
        End If
    Else
        intMoves(Index) = 2
    End If
ElseIf strType = "Freeze" Then
    .tmrChar(Index).Tag = strType
End If
End With
End Sub

'power-up has expired
Public Sub getPowExpire(ByVal Index As Integer, ByVal strPow As String)
With frmMain
'if character's recover has ended, set those effects to end
If strPow = "Recover" Then
    blnRecover(Index) = False
    .tmrChar(Index).Tag = ""
    If gameMode = 0 Then
        .tmrStun(Index).Enabled = False
    End If
'end power-up effects depending on type of power-up
ElseIf strPow = "Scare" Then
    isScared = False
ElseIf strPow = "Speed" Then
    If gameMode = 0 Then
        .tmrChar(Index).Interval = 40
        If Index > 0 Then
            .tmrCPUMove(Index).Interval = 300
        End If
    Else
        intMoves(Index) = 1
    End If
ElseIf strPow = "Freeze" Then
    .tmrChar(Index).Tag = ""
End If
.tmrPow(Index).Enabled = False
End With
End Sub

'tick for power-up; power-up expires when limit is reached
Public Sub getPowTick(ByVal Index As Integer)
tmrPowCounter(Index) = tmrPowCounter(Index) + 1
With frmMain
If tmrPowCounter(Index) >= tmrPowLimit And .tmrChar(Index).Tag <> "Recover" Then
    Call getPowExpire(Index, .tmrPow(Index).Tag)
    tmrPowCounter(Index) = 0
    .tmrPow(Index).Tag = ""
    .tmrPow(Index).Enabled = False
ElseIf .tmrChar(Index).Tag = "Recover" Then
    If (tmrPowCounter(Index) >= tmrPowLimit * 0.75) Or (gameMode <> 0 And tmrPowCounter(Index) >= 1) Then
        Call getPowExpire(Index, .tmrPow(Index).Tag)
        tmrPowCounter(Index) = 0
        .tmrPow(Index).Tag = ""
        .tmrPow(Index).Enabled = False
    End If
End If
End With
End Sub

'change terrain type and picture (for ice tiles); treat this type as an object
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
