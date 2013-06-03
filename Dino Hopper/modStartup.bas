Attribute VB_Name = "modStartup"
Public Sub gameStart()
With frmMain
.lblScore.Visible = True
.lblLives.Visible = True
.lblMulti.Visible = True
.PaintPicture frmMain.picEggMask(0).Image, .lblLives.Left - 24, 0, 22, 30, 0, 0, 22, 30, vbSrcAnd
.PaintPicture frmMain.picEggG(0).Image, .lblLives.Left - 24, 0, 22, 30, 0, 0, 22, 30, vbSrcPaint
defaultTile(0) = tile(mapWidth - 1, mapHeight - 1)
defaultTile(1) = tile(0, mapHeight - 1)
defaultTile(2) = tile(0, 0)
defaultTile(3) = tile(mapWidth - 1, 0)
Dim dtVal(0 To 3) As Integer
For dt = 0 To 3
    If dt > 0 Then
        Dim openVal As Boolean
        Do Until openVal
            dtVal(dt) = randInt(0, 3)
            openVal = True
            For dtv = 0 To dt - 1
                If dtVal(dt) = dtVal(dtv) Then
                    openVal = False
                End If
            Next dtv
        Loop
    Else
        dtVal(dt) = randInt(0, 3)
    End If
    curX(dt) = defaultTile(dtVal(dt)).Xc
    curY(dt) = defaultTile(dtVal(dt)).Yc
    defaultTile(dtVal(dt)).hasChar = True
    If dtVal(dt) = 0 Then
        strDir(dt) = "L"
    ElseIf dtVal(dt) = 1 Then
        strDir(dt) = "U"
    ElseIf dtVal(dt) = 2 Then
        strDir(dt) = "R"
    ElseIf dtVal(dt) = 3 Then
        strDir(dt) = "D"
    End If
    spriteX(dt) = defaultTile(dtVal(dt)).X + 25
    spriteY(dt) = defaultTile(dtVal(dt)).Y - 15
Next dt
.tmrChar(0).Enabled = True
.tmrChar(1).Enabled = True
.tmrObj.Enabled = True
If gameMode <> 1 Then
    .tmrObjEvent.Enabled = True
    objExpire = 125
    .tmrCPUMove(1).Enabled = True
    counterLimit(1) = 1 'cpu 1 moves every (counterLimit + 1) seconds
    tmrPowLimit = 20
ElseIf gameMode = 1 Then
    objExpire = 8
    tmrPowLimit = 8
    intMoves(0) = 1
    intMoves(1) = 1
End If
.tmrAlternate.Enabled = True
blnPlayerMoveable(0) = True
blnPlayerMoveable(1) = True
intLives(0) = 3
intMulti(0) = 1
strState(0) = "I"
strState(1) = "I"
frameLimit(0) = 10
frameLimit(1) = 10
frameLimit(2) = 10
frameLimit(3) = 10
frameProg(0) = 1
frameProg(1) = 1
frameProg(2) = 1
frameProg(3) = 1
End With
End Sub

