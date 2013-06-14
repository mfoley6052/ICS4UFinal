Attribute VB_Name = "modStartup"
'prepare map for game
Public Sub gameStart()
With frmGUI
.Left = frmMain.Left
.Top = frmMain.Top + 200
.Visible = True
'set timer to enabled and GUI values for every input-controlled player
For x = numPlayers - 1 To 0 Step -1
    frmMain.tmrChar(x).Enabled = True
    .lblScore(x).Visible = True
    .lblLives(x).Visible = True
    .lblMulti(x).Visible = True
    .PaintPicture frmMain.picEggMask(0).Image, .lblLives(x).Left - 24, .lblLives(x).Top - (frmMain.picEggG(0).Height - .lblLives(x).Height), 22, 30, 0, 0, 22, 30, vbSrcAnd
    .PaintPicture frmMain.picEggG(0).Image, .lblLives(x).Left - 24, .lblLives(x).Top - (frmMain.picEggG(0).Height - .lblLives(x).Height), 22, 30, 0, 0, 22, 30, vbSrcPaint
    intLives(x) = 3
    intMulti(x) = 1
    intScore(x) = 0
Next x
End With
With frmMain
'set timer enabled and attack mode/speed for computer playrs
If numCPU > 0 Then
    For y = 1 To numCPU
        .tmrChar(y).Enabled = True
        attackMode(y) = randInt(0, 2)
        CPUWait(y) = attackMode(y)
    Next y
End If
defaultTile(0) = tile(mapWidth - 1, mapHeight - 1)
defaultTile(1) = tile(0, mapHeight - 1)
defaultTile(2) = tile(0, 0)
defaultTile(3) = tile(mapWidth - 1, 0)
Dim dtVal(0 To 3) As Integer
'get random default tiles for each character
For dt = 0 To 3
    If dt > 0 Then
        Dim openVal As Boolean
        openVal = False
        Do Until openVal
            openVal = True
            dtVal(dt) = randInt(0, 3)
            For dtv = 0 To dt - 1
                If dtVal(dt) = dtVal(dtv) Then
                    openVal = False
                End If
            Next dtv
        Loop
        If .tmrChar(dt).Enabled Then
            Call assignSprites(dt)
        End If
    Else
        dtVal(dt) = randInt(0, 3)
    End If
    'set tile values for each player
    prevX(dt) = -1
    prevY(dt) = -1
    curX(dt) = defaultTile(dtVal(dt)).Xc
    curY(dt) = defaultTile(dtVal(dt)).Yc
    nextX(dt) = defaultTile(dtVal(dt)).Xc
    nextY(dt) = defaultTile(dtVal(dt)).Yc
    'set tile to have a character on it if the character is enabled
    If .tmrChar(dt).Enabled Then
        tile(defaultTile(dtVal(dt)).Xc, defaultTile(dtVal(dt)).Yc).hasChar = True
    End If
    'choose character directions to face middle based on chosen tile
    If dtVal(dt) = 0 Then
        strDir(dt) = "L"
    ElseIf dtVal(dt) = 1 Then
        strDir(dt) = "U"
    ElseIf dtVal(dt) = 2 Then
        strDir(dt) = "R"
    ElseIf dtVal(dt) = 3 Then
        strDir(dt) = "D"
    End If
    blnBounceJump(dt) = False
    blnEdgeJump(dt) = False
    'set character sprite location
    spriteX(dt) = defaultTile(dtVal(dt)).x + 25
    spriteY(dt) = defaultTile(dtVal(dt)).y - 15
    'set characters to moveable and idle, to land jumps at frame 10 and to jump forward
    blnPlayerMoveable(dt) = True
    strState(dt) = "I"
    frameLimit(dt) = 10
    frameProg(dt) = 1
Next dt
.tmrObj.Enabled = True
'in arcade mode, set objects to appear and expire in real time
If gameMode = 0 Then
    .tmrObjEvent.Enabled = True
    objExpire = 125
    'enable CPU timers in real time if the CPU's are enabled
    If numCPU > 0 Then
        For enableCPU = 1 To numCPU
            .tmrCPUMove(enableCPU).Enabled = True
            counterLimit(enableCPU) = 1 'cpu 1 moves every (counterLimit + 1) seconds
        Next enableCPU
    End If
    'power up time limit
    tmrPowLimit = 20
Else
    'for turn-based modes (turn-based, puzzle), set values for expiry and ticks
    objExpire = 8
    tmrPowLimit = 8
    intMoves(0) = 1
    intMoves(1) = 1
    intMoves(2) = 1
    intMoves(3) = 1
    If gameMode = 2 Then
        blnMoveOnTick(0) = True
    End If
End If
.tmrAlternate.Enabled = True
'game is started
gameStarted = True
'GUI is displayed on top (GUI form takes key input)
frmGUI.SetFocus
End With
End Sub

'set sprite for characters where players 2, 3, and 4 will have different colours if they are controlled by a player or by CPU
Private Sub assignSprites(ByVal Index As Integer)
With frmMain
Dim strPath As String
If isPlayer(Index) Then
    'player 2 is yellow if player-controlled
    If Index = 1 Then
        strPath = "Y\"
    'player 3 is green if player-controlled
    ElseIf Index = 2 Then
        strPath = "G\"
    End If
Else
    'CPU controlled players are different brightnesses of red
    If Index = 1 Then
        strPath = "R1\"
    ElseIf Index = 2 Then
        strPath = "R2\"
    ElseIf Index = 3 Then
        strPath = "R3\"
    End If
End If
'load all images to their pictureboxes on main form
If Index = 1 Then
    .picP2IR.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleR.gif")
    .picP2CR(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1R.gif")
    .picP2CR(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2R.gif")
    .picP2CR(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3R.gif")
    .picP2JR.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpR.gif")
    .picP2IU.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleU.gif")
    .picP2CU(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1U.gif")
    .picP2CU(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2U.gif")
    .picP2CU(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3U.gif")
    .picP2JU.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpU.gif")
    .picP2ID.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleD.gif")
    .picP2CD(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1D.gif")
    .picP2CD(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2D.gif")
    .picP2CD(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3D.gif")
    .picP2JD.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpD.gif")
    .picP2IL.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleL.gif")
    .picP2CL(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1L.gif")
    .picP2CL(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2L.gif")
    .picP2CL(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3L.gif")
    .picP2JL.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpL.gif")
ElseIf Index = 2 Then
    .picP3IR.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleR.gif")
    .picP3CR(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1R.gif")
    .picP3CR(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2R.gif")
    .picP3CR(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3R.gif")
    .picP3JR.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpR.gif")
    .picP3IU.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleU.gif")
    .picP3CU(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1U.gif")
    .picP3CU(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2U.gif")
    .picP3CU(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3U.gif")
    .picP3JU.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpU.gif")
    .picP3ID.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleD.gif")
    .picP3CD(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1D.gif")
    .picP3CD(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2D.gif")
    .picP3CD(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3D.gif")
    .picP3JD.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpD.gif")
    .picP3IL.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleL.gif")
    .picP3CL(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1L.gif")
    .picP3CL(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2L.gif")
    .picP3CL(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3L.gif")
    .picP3JL.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpL.gif")
ElseIf Index = 3 Then
    .picP4IR.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleR.gif")
    .picP4CR(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1R.gif")
    .picP4CR(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2R.gif")
    .picP4CR(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3R.gif")
    .picP4JR.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpR.gif")
    .picP4IU.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleU.gif")
    .picP4CU(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1U.gif")
    .picP4CU(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2U.gif")
    .picP4CU(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3U.gif")
    .picP4JU.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpU.gif")
    .picP4ID.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleD.gif")
    .picP4CD(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1D.gif")
    .picP4CD(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2D.gif")
    .picP4CD(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3D.gif")
    .picP4JD.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpD.gif")
    .picP4IL.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "IdleL.gif")
    .picP4CL(1).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch1L.gif")
    .picP4CL(2).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch2L.gif")
    .picP4CL(3).Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "Crouch3L.gif")
    .picP4JL.Picture = LoadPicture(App.Path & "\Images\Char\" & strPath & "JumpL.gif")
End If
End With
End Sub

'game end
Public Sub getGameEnd()
With frmMain
numPlayers = 1
numCPU = 0
'turn off timers
.tmrTileAnim.Enabled = False
.tmrTileAnimDelay.Enabled = False
.tmrObjEvent.Enabled = False
.tmrObj.Enabled = False
.tmrAlternate.Enabled = False
.tmrRefresh.Enabled = False
intMoveCount = 0
'set character values false
For x = 0 To 3
    blnRecover(x) = False
    tmrPowCounter(x) = False
    .tmrChar(x).Enabled = False
    .tmrChar(x).Tag = ""
    .tmrPow(x).Enabled = False
    .tmrPow(x).Tag = ""
    .tmrStun(x).Enabled = False
    If x > 0 Then
        .tmrCPUMove(x).Enabled = False
    End If
    If x < 3 Then
        frmGUI.lblLives(x).Enabled = False
        frmGUI.lblMulti(x).Enabled = False
        frmGUI.lblScore(x).Enabled = False
    End If
Next x
'set object values false
Dim curTile As terrain
For T = 0 To intTileCount
    tileSwitch(T) = False
    If tile(getTileFromInt(True, T), getTileFromInt(False, T)).hasObj Then
        tile(getTileFromInt(True, T), getTileFromInt(False, T)).objTimer = 0
        tile(getTileFromInt(True, T), getTileFromInt(False, T)).objType(0) = ""
        tile(getTileFromInt(True, T), getTileFromInt(False, T)).objFrame = 0
        tile(getTileFromInt(True, T), getTileFromInt(False, T)).hasObj = False
    End If
Next T
'GUI not visible
frmGUI.Visible = False
'game is not started
gameStarted = False
If canPlay Then
    frmMain.mmcLose.Command = "prev"
    frmMain.mmcLose.Command = "open"
    frmMain.mmcLose.Command = "play"
End If
frmMain.picBackground.Picture = picBG
'paint game over on screen
frmMain.PaintPicture frmMain.picBackground.Image, 0, 0, frmMain.Width, frmMain.Height, 0, 0, frmMain.Width, frmMain.Height, vbSrcCopy
frmMain.PaintPicture frmMain.picGameOverMask, frmMain.Width / 2 - frmMain.picGameOver.Width, frmMain.Height / 2 - frmMain.picGameOver.Height, frmMain.picGameOver.Width, frmMain.picGameOver.Height, 0, 0, frmMain.picGameOver.Width, frmMain.picGameOver.Height, vbSrcAnd
frmMain.PaintPicture frmMain.picGameOver, frmMain.Width / 2 - frmMain.picGameOver.Width, frmMain.Height / 2 - frmMain.picGameOver.Height, frmMain.picGameOver.Width, frmMain.picGameOver.Height, 0, 0, frmMain.picGameOver.Width, frmMain.picGameOver.Height, vbSrcPaint
End With
End Sub
