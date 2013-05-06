Attribute VB_Name = "modStartup"
Option Explicit


Public Sub gameStart()
With frmMain
.tmrChar(0).Enabled = True
.tmrChar(1).Enabled = True
.tmrObjEvent.Enabled = True
.tmrObj.Enabled = True
.tmrCPUMove(1).Enabled = True
blnPlayerMoveable = True
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
strState(0) = "I"
strState(1) = "I"
strDir(0) = "L"
strDir(1) = "R"
spriteX(0) = tile(curX(0), curY(0)).x + 25
spriteY(0) = tile(curX(0), curY(0)).y - 15
spriteX(1) = tile(curX(1), curY(1)).x + 25
spriteY(1) = tile(curX(1), curY(1)).y - 15
Call PaintCharSprite(curX(0), curY(0), 0)
Call PaintCharSprite(curX(1), curY(1), 1)
frameLimit(0) = 10
frameLimit(1) = 10
frameLimit(2) = 10
frameLimit(3) = 10
'cpu 1 moves every (counterLimit + 1) seconds
counterLimit(1) = 1
End With
End Sub

