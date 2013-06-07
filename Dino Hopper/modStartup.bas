Attribute VB_Name = "modStartup"
Public Sub gameStart()
With frmMain
'gameMode = 2
.lblScore.Visible = True
.lblLives.Visible = True
.lblMulti.Visible = True
.PaintPicture frmMain.picEggMask(0).Image, .lblLives.Left - 24, 0, 22, 30, 0, 0, 22, 30, vbSrcAnd
.PaintPicture frmMain.picEggG(0).Image, .lblLives.Left - 24, 0, 22, 30, 0, 0, 22, 30, vbSrcPaint
.tmrChar(0).Enabled = True
If numCPU >= 1 Then
    For X = 1 To numCPU
        .tmrChar(X).Enabled = True
    Next X
End If
numPlayers = 1
defaultTile(0) = tile(mapWidth - 1, mapHeight - 1)
defaultTile(1) = tile(0, mapHeight - 1)
defaultTile(2) = tile(0, 0)
defaultTile(3) = tile(mapWidth - 1, 0)
Dim dtVal(0 To 3) As Integer
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
    curX(dt) = defaultTile(dtVal(dt)).Xc
    curY(dt) = defaultTile(dtVal(dt)).Yc
    nextX(dt) = defaultTile(dtVal(dt)).Xc
    nextY(dt) = defaultTile(dtVal(dt)).Yc
    tile(defaultTile(dtVal(dt)).Xc, defaultTile(dtVal(dt)).Yc).hasChar = True
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
.tmrObj.Enabled = True
If gameMode = 0 Then
    .tmrObjEvent.Enabled = True
    objExpire = 125
    If Not isPlayer(1) Then
        .tmrCPUMove(1).Enabled = True
        counterLimit(1) = 1 'cpu 1 moves every (counterLimit + 1) seconds
    End If
    If Not isPlayer(2) Then
        .tmrCPUMove(2).Enabled = True
         counterLimit(2) = 1
    End If
    If Not isPlayer(3) Then
        .tmrCPUMove(3).Enabled = True
        counterLimit(3) = 1
    End If
    tmrPowLimit = 20
Else
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
blnPlayerMoveable(0) = True
blnPlayerMoveable(1) = True
blnPlayerMoveable(2) = True
blnPlayerMoveable(3) = True
intLives(0) = 3
intMulti(0) = 1
strState(0) = "I"
strState(1) = "I"
strState(2) = "I"
strState(3) = "I"
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

Private Sub assignSprites(ByVal index As Integer)
With frmMain
Dim strPath As String
If isPlayer(index) Then
    If index = 1 Then
        strPath = "Y\"
    ElseIf index = 2 Then
        strPath = "G\"
    End If
Else
    If index = 1 Then
        strPath = "R1\"
    ElseIf index = 2 Then
        strPath = "R2\"
    ElseIf index = 3 Then
        strPath = "R3\"
    End If
End If
If index = 1 Then
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
ElseIf index = 2 Then
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
ElseIf index = 3 Then
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
