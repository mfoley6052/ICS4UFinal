Attribute VB_Name = "modMovement"
Public Sub charAction(ByVal Index As Integer, nextTile As terrain)
If frameCounter(Index) = 1 Then
    strState(Index) = "C"
ElseIf frameCounter(Index) = 5 Then
    strState(Index) = "J"
End If
If gameMode <> 0 Then
    targIndex(Index) = -1
End If
If Not blnEdgeJump(Index) Then
    If gameMode = 0 Then
        Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), nextTile.x, nextTile.y)
    End If
    For charIndex = 0 To 3
        With frmMain
        If charIndex <> Index And .tmrChar(charIndex).Enabled Then
            If nextX(Index) = nextX(charIndex) And nextY(Index) = nextY(charIndex) And gameMode <> 0 Then
                If curY(Index) < curY(charIndex) Then
                    targIndex(Index) = charIndex
                    Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), nextTile.x, spriteY(targIndex(Index)))
                ElseIf gameMode = 1 And curY(Index) = curY(charIndex) Then
                    If frameCounter(Index) = 9 Then
                        frameProg(Index) = -1
                    End If
                    Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), nextTile.x, nextTile.y)
                ElseIf gameMode = 2 And curY(Index) > curY(charIndex) Then
                    If frameCounter(Index) = 9 Then
                        frameProg(Index) = -1
                        Call getHurt(Index, charIndex)
                    End If
                End If
            ElseIf gameMode = 0 And frameCounter(Index) = 9 Then
                If nextX(Index) = curX(charIndex) And nextY(Index) = curY(charIndex) Then
                    frameProg(Index) = -1
                    If frameCounter(charIndex) > 5 Then
                        If (strDir(Index) = "L" And strDir(charIndex) = "R") Or (strDir(Index) = "U" And strDir(charIndex) = "D") Or (strDir(Index) = "R" And strDir(charIndex) = "U") Or (strDir(Index) = "D" And strDir(charIndex) = "U") Then
                            frameProg(charIndex) = -1
                        End If
                        If Not blnRecover(charIndex) Then
                            Call getHurt(charIndex, Index)
                        End If
                        If Not blnRecover(Index) And frameCounter(charIndex) = 9 Then
                            Call getHurt(Index, charIndex)
                        End If
                    Else
                        If Not blnRecover(charIndex) Then
                            Call getJump(charIndex, strDir(Index), evalMove(charIndex, strDir(Index)))
                            If strDir(Index) = "L" Then
                                strDir(charIndex) = "R"
                            ElseIf strDir(Index) = "U" Then
                                strDir(charIndex) = "D"
                            ElseIf strDir(Index) = "R" Then
                                strDir(charIndex) = "L"
                            ElseIf strDir(Index) = "D" Then
                                strDir(charIndex) = "U"
                            End If
                            Call getHurt(charIndex, Index)
                            frameCounter(charIndex) = 6
                        End If
                    End If
                ElseIf nextX(Index) = nextX(charIndex) And nextY(Index) = nextY(charIndex) Then 'going toward same tile
                    If frameCounter(Index) < frameCounter(charIndex) Then 'current char ahead of charIndex char
                        If Not blnRecover(charIndex) Then
                            Call getJump(charIndex, strDir(Index), evalMove(charIndex, strDir(Index)))
                            If strDir(Index) = "L" Then
                                strDir(charIndex) = "R"
                            ElseIf strDir(Index) = "U" Then
                                strDir(charIndex) = "D"
                            ElseIf strDir(Index) = "R" Then
                                strDir(charIndex) = "L"
                            ElseIf strDir(Index) = "D" Then
                                strDir(charIndex) = "U"
                            End If
                            Call getHurt(charIndex, Index)
                            frameCounter(charIndex) = 6
                        Else
                            frameProg(Index) = -1
                            Call getHurt(Index, charIndex)
                        End If
                    ElseIf frameCounter(Index) = frameCounter(charIndex) Then 'both chars at same frame; head-on collision
                        frameProg(Index) = -1
                        frameProg(charIndex) = -1
                        If Not blnRecover(Index) Then
                            Call getHurt(Index, charIndex)
                        End If
                        If Not blnRecover(charIndex) Then
                            Call getHurt(charIndex, Index)
                        End If
                    End If
                End If
            End If
        End If
        End With
    Next charIndex
    If targIndex(Index) < 0 Then
        Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), tile(nextX(Index), nextY(Index)).x, tile(nextX(Index), nextY(Index)).y)
    End If
    Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
    If frameCounter(Index) = 10 Then
        strState(Index) = "I"
    End If
    If (frameCounter(Index) = frameLimit(Index)) Or (frameProg(Index) < 0 And frameCounter(Index) <= 5) Then
        If gameMode = 0 Then
            Call getJumpComplete(Index)
            frameProg(Index) = 1
        ElseIf checkTickComplete And gameMode = 1 Then
            For getNext = 0 To 3
                Call getTickComplete(getNext)
            Next getNext
        ElseIf gameMode = 2 Then
            Call getTickComplete(Index)
            If Index < 3 Then
                If Not isPlayer(Index + 1) Then
                    Call getTick(Index + 1)
                End If
            End If
        End If
    Else
        frameCounter(Index) = frameCounter(Index) + frameProg(Index)
    End If
Else 'jump off edge
    Dim curTile As terrain
    curTile = tile(curX(Index), curY(Index))
    Dim altTile As terrain
    If strDir(Index) = "L" Then
        Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), tile(curX(Index), curY(Index)).x - 50, tile(curX(Index), curY(Index)).y - 75)
        If curY(Index) = 0 Then
            If spriteY(Index) < curTile.y + 50 Or (spriteX(Index) > -35 And spriteX(Index) < tile(mapWidth - 1, 0).x + 135 And spriteY(Index) < curTile.y + 25) Then
                Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
            End If
            Call clearTile(curTile, True, Index, "CharTop+X-Y")
            If curX(Index) > 0 Then
                altTile = tile(curTile.Xc - 1, curTile.Yc)
                If altTile.hasObj Then
                    Call PaintObj(altTile.objType(0), altTile.objType(1), altTile.objFrame, altTile.Xc, altTile.Yc, False)
                End If
            End If
        ElseIf curX(Index) = 0 And curY(Index) > 0 Then
            Call clearTile(curTile, True, Index, "CharSide-X-Y")
            Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
        End If
    ElseIf strDir(Index) = "U" Then
        Call getCharJumpAnim(Index, frameCounter(Index), curTile, curTile.x + 50, curTile.y - 75)
        If curY(Index) = 0 Then
            If spriteY(Index) < curTile.y + 50 Or (spriteX(Index) > -35 And spriteX(Index) < tile(mapWidth - 1, 0).x + 135 And spriteY(Index) < curTile.y + 25) Then
                Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
            End If
            Call clearTile(curTile, True, Index, "CharTop+X-Y")
            If curX(Index) < mapWidth - 1 Then
                altTile = tile(curTile.Xc + 1, curTile.Yc)
                If altTile.hasObj Then
                    Call PaintObj(altTile.objType(0), altTile.objType(1), altTile.objFrame, altTile.Xc, altTile.Yc, False)
                End If
            End If
        ElseIf curX(Index) = mapWidth And curY(Index) > 0 Then
            Call clearTile(curTile, True, Index, "CharSide+X-Y")
            Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
        End If
    ElseIf strDir(Index) = "R" Then
        Call getCharJumpAnim(Index, frameCounter(Index), curTile, curTile.x + 50, curTile.y + 75)
        If curX(Index) = mapWidth And curY(Index) > 0 Then
            Call clearTile(curTile, True, Index, "CharSide+X+Y")
        ElseIf curY(Index) = mapHeight - 1 Then
            Call clearTile(curTile, True, Index, "CharBottom+X+Y")
        End If
        Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
    ElseIf strDir(Index) = "D" Then
        Call getCharJumpAnim(Index, frameCounter(Index), curTile, curTile.x - 50, curTile.y + 75)
        If curX(Index) = 0 And (curY(Index) > 0 And curY(Index) < mapHeight - 1) Then
            Call clearTile(curTile, True, Index, "CharSide-X+Y")
        ElseIf curY(Index) = mapHeight - 1 Then
            Call clearTile(curTile, True, Index, "CharBottom-X+Y")
        End If
        Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
    End If
    If frameCounter(Index) = frameLimit(Index) * 1.6 Then
        spriteX(Index) = tile(nextX(Index), nextY(Index)).x + 25
        spriteY(Index) = tile(nextX(Index), nextY(Index)).y - 15
        If gameMode = 0 Then
            Call getJumpComplete(Index)
            frameProg(Index) = 1
        ElseIf checkTickComplete And gameMode = 1 Then
            For getNext = 0 To 3
                Call getTickComplete(getNext)
            Next getNext
        ElseIf gameMode = 2 Then
            Call getTickComplete(Index)
            If Index < 3 Then
                If Not isPlayer(Index + 1) Then
                    Call getTick(Index + 1)
                End If
            End If
        End If
    Else
        frameCounter(Index) = frameCounter(Index) + frameProg(Index)
    End If
End If
End Sub

Public Function getJump(ByVal Index As Integer, ByVal strDirJ As String, ByVal blnTile As Boolean)
With frmMain
    If frameCounter(Index) = 0 Then
        strDir(Index) = strDirJ
        blnEdgeJump(Index) = Not blnTile
        If blnPlayerMoveable(Index) Or (gameMode <> 0 And blnMoveOnTick(Index)) Then
            blnPlayerMoveable(Index) = False
            If blnTile Then
                If strDir(Index) = "L" Then
                    'if y row is odd
                    If oddRow(curY(Index)) Then
                        nextX(Index) = curX(Index) - 1
                    Else
                        nextX(Index) = curX(Index)
                    End If
                    nextY(Index) = curY(Index) - 1
                ElseIf strDir(Index) = "U" Then
                    'if y row is even
                    If Not oddRow(curY(Index)) Then
                        nextX(Index) = curX(Index) + 1
                    Else
                        nextX(Index) = curX(Index)
                    End If
                    nextY(Index) = curY(Index) - 1
                ElseIf strDir(Index) = "R" Then
                    'if y row is even
                    If Not oddRow(curY(Index)) Then
                        nextX(Index) = curX(Index) + 1
                    Else
                        nextX(Index) = curX(Index)
                    End If
                    nextY(Index) = curY(Index) + 1
                ElseIf strDir(Index) = "D" Then
                    'if y row is odd
                    If oddRow(curY(Index)) Then
                        nextX(Index) = curX(Index) - 1
                    Else
                        nextX(Index) = curX(Index)
                    End If
                    nextY(Index) = curY(Index) + 1
                End If
            ElseIf Not blnRecover(Index) Then
                Call setCharRespawn(Index, tile(curX(Index), curY(Index)))
            End If
            If gameMode = 0 Then
                frameCounter(Index) = 1
            End If
        End If
    End If
End With
End Function
Public Function randDir() As String
Dim temp As Integer
temp = Int(Rnd() * 4) + 1
If temp = 1 Then
    randDir = "L"
ElseIf temp = 2 Then
    randDir = "U"
ElseIf temp = 3 Then
    randDir = "R"
Else
    randDir = "D"
End If
End Function
Public Sub getJumpComplete(ByVal Index As Integer)
Dim pScore As Integer
Dim q As Integer
With frmMain
prevX(Index) = curX(Index)
prevY(Index) = curY(Index)

strState(Index) = "I"
frameCounter(Index) = 0
blnPlayerMoveable(Index) = True

tile(curX(Index), curY(Index)).hasChar = False
If blnBounceJump(Index) And targIndex(Index) >= 0 Then
    blnBounceJump(Index) = False
    tile(curX(targIndex(Index)), curY(targIndex(Index))).hasChar = False
End If
If blnEdgeJump(Index) Then
    tile(curX(Index), curY(Index)).hasChar = False
    Call getHurt(Index, Index)
End If

If gameMode = 0 And frameProg(Index) <= 0 Then
    nextX(Index) = curX(Index)
    nextY(Index) = curY(Index)
End If

curX(Index) = nextX(Index)
curY(Index) = nextY(Index)

Dim inputTile As terrain
inputTile = tile(curX(Index), curY(Index))

tile(curX(Index), curY(Index)).hasChar = True
If isPlayer(Index) Then
    If inputTile.hasObj Then
        Dim blnMulti As Boolean
        Dim blnLives As Boolean
        If inputTile.objType(0) = "Coin" Then
            'play coin sound
            If inputTile.objType(1) = "Y" Then
                pScore = 100
            ElseIf inputTile.objType(1) = "R" Then
                pScore = 250
            ElseIf inputTile.objType(1) = "B" Then
                pScore = 500
            End If
        ElseIf inputTile.objType(0) = "Pow" Then
            'play scare power-up sound
            If inputTile.objType(1) = "Scare" Then
                pScore = 200
            End If
        ElseIf inputTile.objType(0) = "Egg" Then
            If inputTile.objType(1) = "M" Then
                intMulti(Index) = intMulti(Index) + 1
                pScore = 1000
                blnMulti = True
            ElseIf inputTile.objType(1) = "G" Then
                intLives(Index) = intLives(Index) + 1
                blnLives = True
            End If
        End If
    ElseIf blnBounceJump(Index) Then 'jumping on enemy
        pScore = 1000
    End If
    If inputTile.terType = "G" Then
        pScore = pScore + 25
    End If
    If Not blnEdgeJump(Index) Then
        Call addScore(Index, pScore)
    End If
    Call refreshLabels(Index, True, blnLives, blnMulti)
End If
If inputTile.hasObj Then
    If inputTile.objType(0) <> "Terrain" Then
        strState(Index) = "I"
        Call killObj(tile(curX(Index), curY(Index)))
        If inputTile.objType(0) = "Pow" Then
            Call getPowEffect(Index, inputTile.objType(1))
        End If
    Else
        If inputTile.objType(1) = "I" Then
            Call getJump(Index, strDir(Index), evalMove(Index, strDir(Index)))
            frameCounter(Index) = 1
        End If
    End If
End If
If Not blnBounceJump(Index) Then
    blnClearPrevTile(Index) = True
End If
End With
End Sub

Public Function evalMove(ByVal Index As Integer, ByVal strDirMove As String) As Boolean
If strDirMove = "L" Then
    If oddRow(curY(Index)) Then
        If curX(Index) = 0 Then
            evalMove = False
        ElseIf curY(Index) > 0 Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY(Index) > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "U" Then
    If oddRow(curY(Index)) Then
        If curX(Index) < mapWidth Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curX(Index) < (mapWidth) And curY(Index) > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "R" Then
    If oddRow(curY(Index)) Then
        If curX(Index) = mapWidth Then
            evalMove = False
        ElseIf curY(Index) > 0 Then
            evalMove = True
        End If
    ElseIf curY(Index) < (mapWidth - 1) And curX(Index) < mapWidth Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "D" Then
    If oddRow(curY(Index)) Then
        If curX(Index) > 0 And curY(Index) < mapHeight Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY(Index) < (mapHeight - 1) Then
        evalMove = True
    Else
        evalMove = False
    End If
End If
End Function

Public Function checkTickComplete() As Boolean
checkTickComplete = True
With frmMain
For checkCharJump = 0 To 3
    If .tmrChar(checkCharJump).Enabled Then
        If frameCounter(checkCharJump) > 0 And Not blnEdgeJump(checkCharJump) Then
            If frameCounter(checkCharJump) < frameLimit(checkCharJump) And frameProg(checkCharJump) > 0 Then
                checkTickComplete = False
            ElseIf frameProg(checkCharJump) < 0 Then
                If frameCounter(checkCharJump) > 5 Then
                    checkTickComplete = False
                End If
            End If
        ElseIf blnEdgeJump(charCheckJump) Then
            If frameCounter(checkCharJump) < frameLimit(checkCharJump) * 1.6 Then
                checkTickComplete = False
            End If
        End If
    End If
Next checkCharJump
End With
End Function

Public Sub getTickComplete(ByVal Index As Integer)
With frmMain
If .tmrChar(Index).Enabled And blnMoveOnTick(Index) Then
    If frameProg(Index) < 0 Then
        frameProg(Index) = 1
        nextX(Index) = curX(Index)
        nextY(Index) = curY(Index)
    End If
    Call getJumpComplete(Index)
    If Not blnBounceJump(Index) And targIndex(Index) >= 0 Then
        blnBounceJump(Index) = True
        If Not blnRecover(targIndex(Index)) Then
            Call getHurt(targIndex(Index), Index)
        End If
        Call getJump(Index, strDir(Index), evalMove(Index, strDir(Index)))
        frameCounter(Index) = 1
    ElseIf gameMode = 2 And Index < 3 Then
        blnMoveOnTick(Index) = False
        If isPlayer(Index + 1) Then
            blnMoveOnTick(Index + 1) = True
        ElseIf isPlayer(Index) And numPlayers = 1 Then
            Call getTick(Index + 1)
        End If
    End If
End If
End With
End Sub

Public Sub getTick(Optional Index As Integer)
Dim loopMin As Integer
Dim loopMax As Integer
If gameMode = 1 Then
    Index = -1
    loopMin = 0
    loopMax = 3
ElseIf gameMode = 2 Then
    loopMin = Index
    loopMax = Index
    With frmMain
    For disableMove = 0 To 3
        If disableMove <> Index And .tmrChar(disableMove).Enabled Then
            blnMoveOnTick(disableMove) = False
        End If
    Next disableMove
    End With
End If
With frmMain
Dim highestMove As Integer 'highest move count out of all characters
highestMove = 1 'set highest move to default
For moveCheck = 0 To 3
    If .tmrChar(moveCheck).Enabled Then 'check for highest move and set it
        If intMoves(moveCheck) > highestMove Then
            highestMove = intMoves(moveCheck)
        End If
    End If
Next moveCheck
For moveCheck2 = loopMin To loopMax 'disable blnMove for characters that are not set to move according to higher character speeds
    If intMoves(moveCheck2) < highestMove - intMoveCount Then
        blnMoveOnTick(moveCheck2) = False
    ElseIf .tmrChar(moveCheck2).Enabled Then
        blnMoveOnTick(moveCheck2) = True
    End If
Next moveCheck2
For setMove = loopMin To loopMax
    If isPlayer(setMove) Then 'if setMove is a player, call getJump (direction change and next values set)
        Call getJump(setMove, strDir(setMove), evalMove(setMove, strDir(setMove)))
    ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(setMove)) Then
        Call cpuAI(setMove) 'if computer player, call AI (direction change and next values)
    End If
Next setMove
For moveCheck3 = loopMin To loopMax
    If blnMoveOnTick(moveCheck3) Then 'if character can move
        For checkTarg = 0 To 3
            If checkTarg <> moveCheck3 And .tmrChar(checkTarg).Enabled Then 'if index is different and checkTarg char is active
                If gameMode = 1 Then
                    'if checkMove3 char and checkTarg char are moving toward eachother
                    If nextX(moveCheck3) = curX(checkTarg) And nextY(moveCheck3) = curY(checkTarg) And curX(moveCheck3) = nextX(checkTarg) And curY(moveCheck3) = nextY(checkTarg) Then
                        If curY(moveCheck3) < curY(checkTarg) Then 'if checkTarg char is lower on map
                            blnMoveOnTick(checkTarg) = False 'checkTarg char can no longer move
                        End If
                    End If
                ElseIf gameMode = 2 Then
                    If prevX(checkTarg) = curX(moveCheck3) And prevY(checkTarg) = curY(moveCheck3) Then
                        blnMoveOnTick(moveCheck3) = False 'checkTarg char can no longer move
                    End If
                End If
            End If
        Next checkTarg
    End If
Next moveCheck3
For getMove = loopMin To loopMax 'get movement (or not)
    If .tmrPow(getMove).Tag <> "" Then 'if corresponding pow timer has a tag (power-up activated), call a tick
        Call getPowTick(getMove)
    End If
    If blnMoveOnTick(getMove) Then 'if char can move, set it to start animation
        frameCounter(getMove) = 1
    Else 'if not, set next values to current values
        nextX(getMove) = curX(getMove)
        nextY(getMove) = curY(getMove)
        If gameMode = 2 And Index < 3 Then
            If isPlayer(Index + 1) Then
                blnMoveOnTick(Index + 1) = True
            ElseIf numPlayers = 1 Then
                Call getTick(Index + 1)
            End If
        End If
    End If
Next getMove
If gameMode = 1 Or (gameMode = 2 And Index = 3) Then
    intMoveCount = intMoveCount + 1 'move count increases by one
    If highestMove - intMoveCount <= 0 Then 'reset intMoveCount if it, subtracted from the highest move, is below default speed
        intMoveCount = 0
    End If
    For T = 0 To tileCount - 1 'check each tile for object, if it has one, call an object timer tick
        If tile(getTileFromInt(True, T), getTileFromInt(False, T)).hasObj Then
            tile(getTileFromInt(True, T), getTileFromInt(False, T)).objTimer = tile(getTileFromInt(True, T), getTileFromInt(False, T)).objTimer + 1
        End If
    Next T
    Call .tmrObjEvent_Timer 'call an object to appear
    If gameMode = 2 Then
        blnMoveOnTick(0) = True
    End If
End If
End With
End Sub
