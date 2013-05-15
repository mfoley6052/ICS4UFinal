Attribute VB_Name = "modStartup"
Option Explicit


Public Sub gameStart()

With frmMain
.lblScore.Visible = True
.lblLives.Visible = True
.lblMulti.Visible = True
.PaintPicture frmMain.picEggMask(0).Image, .lblLives.Left - 24, 0, 22, 30, 0, 0, 22, 30, vbSrcAnd
.PaintPicture frmMain.picEggG(0).Image, .lblLives.Left - 24, 0, 22, 30, 0, 0, 22, 30, vbSrcPaint
.tmrChar(0).Enabled = True
.tmrChar(1).Enabled = True
.tmrObj.Enabled = True
If gameMode <> 1 Then
    .tmrObjEvent.Enabled = True
    objExpire = 100
    .tmrCPUMove(1).Enabled = True
    counterLimit(1) = 1 'cpu 1 moves every (counterLimit + 1) seconds
End If
blnPlayerMoveable = True
intLives(0) = 3
intMulti(0) = 1
curX(0) = (mapWidth - 1) \ 2
curY(0) = (mapHeight - 1) \ 2
nextX(0) = (mapWidth - 1) \ 2
nextY(0) = (mapHeight - 1) \ 2
curX(1) = 0
curY(1) = 0
nextX(1) = 0
nextY(1) = 0
tile(curX(0), curY(0)).hasChar = True
tile(curX(1), curY(1)).hasChar = True
spriteX(0) = tile(curX(0), curY(0)).x + 25
spriteY(0) = tile(curX(0), curY(0)).y - 15
spriteX(1) = tile(curX(1), curY(1)).x + 25
spriteY(1) = tile(curX(1), curY(1)).y - 15
strState(0) = "I"
strState(1) = "I"
strDir(0) = "L"
strDir(1) = "R"
frameLimit(0) = 10
frameLimit(1) = 10
frameLimit(2) = 10
frameLimit(3) = 10
End With
End Sub

