Attribute VB_Name = "modMovement"
'execute character jump (or any other action)
Public Sub charAction(ByVal Index As Integer, nextTile As terrain)
If blnDebug Then
    'debug data for character locations if debug is activated
    With frmDbg
    .lblTest(0) = "P1 curTile: (" & curX(0) & ", " & curY(0) & ")"
    .lblTest(1) = "P1 nextTile: (" & nextX(0) & ", " & nextY(0) & ")"
    .lblTest(2) = "P2 curTile: (" & curX(1) & ", " & curY(1) & ")"
    .lblTest(3) = "P2 nextTile: (" & nextX(1) & ", " & nextY(1) & ")"
    .lblTest(4) = "P3 curTile: (" & curX(2) & ", " & curY(2) & ")"
    .lblTest(5) = "P3 nextTile: (" & nextX(2) & ", " & nextY(2) & ")"
    .lblTest(6) = "P4 curTile: (" & curX(3) & ", " & curY(3) & ")"
    .lblTest(7) = "P4 nextTile: (" & nextX(3) & ", " & nextY(3) & ")"
    End With
End If
If frameCounter(Index) = 1 Then
    strState(Index) = "C" 'character is crouching on frame 1
ElseIf frameCounter(Index) = 5 Then
    strState(Index) = "J" 'character is jumping on frame 5 and on
End If
'if game mode isn't 0 then target value is set to -1 so it isn't use with an index
If gameMode <> 0 Then
    targIndex(Index) = -1
End If
'if not jumping off edge
If Not blnEdgeJump(Index) Then
    'get new location for character in jump
    If gameMode = 0 Then
        Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), nextTile.x, nextTile.y)
    End If
    For charIndex = 0 To 3
        With frmMain
        If charIndex <> Index And .tmrChar(charIndex).Enabled Then
            'character and target are jumping to same tile
            If nextX(Index) = nextX(charIndex) And nextY(Index) = nextY(charIndex) And gameMode <> 0 Then
                'if character at higher tile than target, jump on character
                If curY(Index) < curY(charIndex) Then
                    targIndex(Index) = charIndex
                    Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), nextTile.x, spriteY(targIndex(Index)))
                'character collision in puzzle mode
                ElseIf gameMode = 1 And curY(Index) = curY(charIndex) Then
                    'index character jumps back to its tile (set to reverse)
                    If frameCounter(Index) = 9 Then
                        frameProg(Index) = -1
                    End If
                    Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), nextTile.x, nextTile.y)
                'character hit in turn-based mode
                ElseIf gameMode = 2 And curY(Index) > curY(charIndex) Then
                    If frameCounter(Index) = 9 Then
                        frameProg(Index) = -1
                        'target is damaged if not recovering
                        If Not blnRecover(charIndex) Then
                            Call getHurt(Index, charIndex)
                            'this ends the sub to end codes when a game over is reached
                            If Not gameStarted Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            'arcade mode and character is at frame 9
            ElseIf gameMode = 0 And frameCounter(Index) = 9 Then
                'characters jumping to same tile
                If nextX(Index) = curX(charIndex) And nextY(Index) = curY(charIndex) Then
                    frameProg(Index) = -1
                    'target has started jumping too
                    If frameCounter(charIndex) > 5 Then
                        'target moves back if facing attacking character
                        If (strDir(Index) = "L" And strDir(charIndex) = "R") Or (strDir(Index) = "U" And strDir(charIndex) = "D") Or (strDir(Index) = "R" And strDir(charIndex) = "U") Or (strDir(Index) = "D" And strDir(charIndex) = "U") Then
                            frameProg(charIndex) = -1
                        End If
                        'target is damaged if not recovering
                        If Not blnRecover(charIndex) Then
                            Call getHurt(charIndex, Index)
                            If Not gameStarted Then
                                Exit Sub
                            End If
                        End If
                        'if characters collide (target jumped at same time), character is damaged if not recovering
                        If Not blnRecover(Index) And frameCounter(charIndex) = 9 Then
                            Call getHurt(Index, charIndex)
                            If Not gameStarted Then
                                Exit Sub
                            End If
                        End If
                    Else
                        'target is on ground
                        If Not blnRecover(charIndex) Then
                            'character is set to get pushed backward to the tile behind
                            Call getJump(charIndex, strDir(Index), evalMove(charIndex, strDir(Index)))
                            If strDir(Index) = "L" Then
                                altDir(charIndex) = "R"
                            ElseIf strDir(Index) = "U" Then
                                altDir(charIndex) = "D"
                            ElseIf strDir(Index) = "R" Then
                                altDir(charIndex) = "L"
                            ElseIf strDir(Index) = "D" Then
                                altDir(charIndex) = "U"
                            End If
                            If Not blnRecover(charIndex) Then
                                Call getHurt(charIndex, Index)
                                If Not gameStarted Then
                                    Exit Sub
                                End If
                            End If
                            frameCounter(charIndex) = 6
                        End If
                    End If
                ElseIf nextX(Index) = nextX(charIndex) And nextY(Index) = nextY(charIndex) Then 'going toward same tile
                    If frameCounter(Index) < frameCounter(charIndex) Then 'character ahead of target
                        If Not blnRecover(charIndex) Then
                            Call getJump(charIndex, strDir(Index), evalMove(charIndex, strDir(Index)))
                            If strDir(Index) = "L" Then
                                altDir(charIndex) = "R"
                            ElseIf strDir(Index) = "U" Then
                                altDir(charIndex) = "D"
                            ElseIf strDir(Index) = "R" Then
                                altDir(charIndex) = "L"
                            ElseIf strDir(Index) = "D" Then
                                altDir(charIndex) = "U"
                            End If
                            If Not blnRecover(charIndex) Then
                                Call getHurt(charIndex, Index)
                                If Not gameStarted Then
                                    Exit Sub
                                End If
                            End If
                            frameCounter(charIndex) = 6
                        Else
                            frameProg(Index) = -1
                            Call getHurt(Index, charIndex)
                            If Not gameStarted Then
                                Exit Sub
                            End If
                        End If
                    ElseIf frameCounter(Index) = frameCounter(charIndex) Then 'both chars at same frame; head-on collision
                        frameProg(Index) = -1
                        frameProg(charIndex) = -1
                        If Not blnRecover(Index) Then
                            Call getHurt(Index, charIndex)
                            If Not gameStarted Then
                                Exit Sub
                            End If
                        End If
                        If Not blnRecover(charIndex) Then
                            Call getHurt(charIndex, Index)
                            If Not gameStarted Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                'target is jumping to character's tile current and character is moving backward, push target back to avoid a glitch
            ElseIf gameMode = 1 And curX(Index) = nextX(charIndex) And curY(Index) = nextY(charIndex) And frameProg(Index) < 0 Then
                'index character jumps back to its tile (set to reverse)
                If frameCounter(Index) = 9 Then
                    frameProg(Index) = -1
                End If
                Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), nextTile.x, nextTile.y)
            End If
        End If
        End With
    Next charIndex
    'if character has no target, jump to tile as normal
    If targIndex(Index) < 0 Then
        Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), tile(nextX(Index), nextY(Index)).x, tile(nextX(Index), nextY(Index)).y)
    End If
    'paint character sprite
    Call PaintCharSprite(Index, spriteX(Index), spriteY(Index), False)
    'character is idle at frame 10
    If frameCounter(Index) = 10 Then
        strState(Index) = "I"
    End If
    'if frame limit is reached, get jump completion or tick completion depending on mode
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
        End If
    Else 'if not reached, next frame
        frameCounter(Index) = frameCounter(Index) + frameProg(Index)
    End If
Else 'jump off edge
    Dim curTile As terrain
    curTile = tile(curX(Index), curY(Index))
    Dim altTile As terrain
    'get painting for jumping off edges in each direction
    If strDir(Index) = "L" Then
        Call getCharJumpAnim(Index, frameCounter(Index), tile(curX(Index), curY(Index)), tile(curX(Index), curY(Index)).x - 50, tile(curX(Index), curY(Index)).y - 75)
        If curY(Index) = 0 Then
            If spriteY(Index) < curTile.y + 50 Or (spriteX(Index) > -35 And spriteX(Index) < tile(mapWidth - 1, 0).x + 135 And spriteY(Index) < curTile.y + 25) Then
                Call PaintCharSprite(Index, spriteX(Index), spriteY(Index), False)
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
            Call PaintCharSprite(Index, spriteX(Index), spriteY(Index), False)
        End If
    ElseIf strDir(Index) = "U" Then
        Call getCharJumpAnim(Index, frameCounter(Index), curTile, curTile.x + 50, curTile.y - 75)
        If curY(Index) = 0 Then
            If spriteY(Index) < curTile.y + 50 Or (spriteX(Index) > -35 And spriteX(Index) < tile(mapWidth - 1, 0).x + 135 And spriteY(Index) < curTile.y + 25) Then
                Call PaintCharSprite(Index, spriteX(Index), spriteY(Index), False)
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
            Call PaintCharSprite(Index, spriteX(Index), spriteY(Index), False)
        End If
    ElseIf strDir(Index) = "R" Then
        Call getCharJumpAnim(Index, frameCounter(Index), curTile, curTile.x + 50, curTile.y + 75)
        If curX(Index) = mapWidth And curY(Index) > 0 Then
            Call clearTile(curTile, True, Index, "CharSide+X+Y")
        ElseIf curY(Index) = mapHeight - 1 Then
            Call clearTile(curTile, True, Index, "CharBottom+X+Y")
        End If
        Call PaintCharSprite(Index, spriteX(Index), spriteY(Index), False)
    ElseIf strDir(Index) = "D" Then
        Call getCharJumpAnim(Index, frameCounter(Index), curTile, curTile.x - 50, curTile.y + 75)
        If curX(Index) = 0 And (curY(Index) > 0 And curY(Index) < mapHeight - 1) Then
            Call clearTile(curTile, True, Index, "CharSide-X+Y")
        ElseIf curY(Index) = mapHeight - 1 Then
            Call clearTile(curTile, True, Index, "CharBottom-X+Y")
        End If
        Call PaintCharSprite(Index, spriteX(Index), spriteY(Index), False)
    End If
    'if edge jump frame limit reached (normal limit * 1.6), set character to go to next tile (tile jumped from if recovering, respawn tile if not)
    If frameCounter(Index) = frameLimit(Index) * 1.6 Then
        'set character to new coordinates
        spriteX(Index) = tile(nextX(Index), nextY(Index)).x + 25
        spriteY(Index) = tile(nextX(Index), nextY(Index)).y - 15
        'get jump complete or tick complete
        If gameMode = 0 Then
            Call getJumpComplete(Index)
            frameProg(Index) = 1
        ElseIf checkTickComplete And gameMode = 1 Then
            For getNext = 0 To 3
                Call getTickComplete(getNext)
            Next getNext
        ElseIf gameMode = 2 Then
            Call getTickComplete(Index)
        End If
    Else 'if not, next frame
        frameCounter(Index) = frameCounter(Index) + frameProg(Index)
    End If
End If
End Sub

'get character jump
Public Function getJump(ByVal Index As Integer, ByVal strDirJ As String, ByVal blnTile As Boolean)
With frmMain
    If frameCounter(Index) = 0 Then
        strDir(Index) = strDirJ
        blnEdgeJump(Index) = Not blnTile
        'if character can jump, set next tiles and direction change
        If (gameMode = 0 And blnPlayerMoveable(Index)) Or (gameMode <> 0 And blnMoveOnTick(Index)) Then
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
            'set char to move to respawn tile if jumping off edge
            ElseIf Not blnRecover(Index) Then
                Call setCharRespawn(Index, tile(curX(Index), curY(Index)))
            End If
            'if arcade mode, set character to begin jump animation
            If gameMode = 0 Then
                frameCounter(Index) = 1
            End If
        End If
    End If
End With
End Function

'get a random direction
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

'get jump completion
Public Sub getJumpComplete(ByVal Index As Integer)
Dim Pscore As Long
Dim q As Integer
With frmMain
'set previous tile coordinates
prevX(Index) = curX(Index)
prevY(Index) = curY(Index)
'if character is not facing the direction it jumped, character changes to the direction it was facing
If altDir(Index) <> "" Then
    strDir(Index) = altDir(Index)
    altDir(Index) = ""
End If

strState(Index) = "I"
frameCounter(Index) = 0

'if character has moved
If nextX(Index) <> curX(Index) Or nextY(Index) <> curY(Index) Then
    tile(curX(Index), curY(Index)).hasChar = False
End If
'if jump is the character bouncing off another, character can move and last tile has a character
If blnBounceJump(Index) Then
    blnBounceJump(Index) = False
    tile(prevX(Index), prevY(Index)).hasChar = True
    blnPlayerMoveable(Index) = True
'if not a bounce jump and mode is not turn based and character has no target, character can move
ElseIf gameMode <> 2 Or targIndex(Index) < 0 Then
    blnPlayerMoveable(Index) = True
End If
'if jumping off edge, its tile has no character and damage character
If blnEdgeJump(Index) Then
    tile(curX(Index), curY(Index)).hasChar = False
    Call getHurt(Index, Index)
    If Not gameStarted Then
        Exit Sub
    End If
End If

'if character is moving backward in arcade mode, set next coordinate to its last tile
If gameMode = 0 And frameProg(Index) <= 0 Then
    nextX(Index) = curX(Index)
    nextY(Index) = curY(Index)
End If

'set current coordinate to next coordinates
curX(Index) = nextX(Index)
curY(Index) = nextY(Index)

Dim inputTile As terrain
inputTile = tile(curX(Index), curY(Index))

'current tile has a character
tile(curX(Index), curY(Index)).hasChar = True
'player-oriented updates (sounds, add points, object scores, etc.)
If isPlayer(Index) Then
    If inputTile.hasObj Then
        Dim blnMulti As Boolean
        Dim blnLives As Boolean
        If inputTile.objType(0) = "Coin" Then
            'play coin sound
            If canPlay Then
                frmMain.mmcCoin.Command = "prev"
                frmMain.mmcCoin.Command = "open"
                frmMain.mmcCoin.Command = "play"
            End If
            If inputTile.objType(1) = "Y" Then
                Pscore = 100
            ElseIf inputTile.objType(1) = "R" Then
                Pscore = 250
            ElseIf inputTile.objType(1) = "B" Then
                Pscore = 500
            End If
        ElseIf inputTile.objType(0) = "Pow" Then
            'play scare power-up sound
            If inputTile.objType(1) = "Scare" Then
                Pscore = 200
            End If
        ElseIf inputTile.objType(0) = "Egg" Then
            If inputTile.objType(1) = "M" Then
                intMulti(Index) = intMulti(Index) + 1
                Pscore = 1000
                blnMulti = True
            ElseIf inputTile.objType(1) = "G" Then
                intLives(Index) = intLives(Index) + 1
                blnLives = True
            End If
        End If
    ElseIf blnBounceJump(Index) Then 'jumping on enemy
        Pscore = 1000
    End If
    If inputTile.terType = "G" Then
        Pscore = Pscore + 25
    End If
    If Not blnEdgeJump(Index) Then
        Call addScore(Index, Pscore)
    End If
    Call refreshLabels(Index, True, blnLives, blnMulti)
Else
End If
'collect object
If inputTile.hasObj Then
    If inputTile.objType(0) <> "Terrain" Then
        strState(Index) = "I"
        Call killObj(tile(curX(Index), curY(Index)))
        If inputTile.objType(0) = "Pow" Then
            Call getPowEffect(Index, inputTile.objType(1))
        End If
    Else
        'jump again on an ice tile
        If inputTile.objType(1) = "I" Then
            Call getJump(Index, strDir(Index), evalMove(Index, strDir(Index)))
            frameCounter(Index) = 1
        End If
    End If
End If
'as long as character is not bouncing on another, set last tile to clear the select image
If Not blnBounceJump(Index) Then
    blnClearPrevTile(Index) = True
End If
End With
End Sub

'return whether a given direction will take a given character to a tile or off the edge
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

'check and return whether all characters have jumped
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

'get a complete tick ,call jump complete for a given character
'also call tick for next character as long as it isn't player-controlled and able to move
Public Sub getTickComplete(ByVal Index As Integer)
With frmMain
If .tmrChar(Index).Enabled And blnMoveOnTick(Index) Then
    'if moving backward, set to move forward and set next tile to current tile
    If frameProg(Index) < 0 Then
        frameProg(Index) = 1
        nextX(Index) = curX(Index)
        nextY(Index) = curY(Index)
    End If
    Call getJumpComplete(Index)
    'if character has a target, get bounce jump and target can't move for the next turn
    If targIndex(Index) >= 0 Then
        blnBounceJump(Index) = True
        If Not blnRecover(targIndex(Index)) Then
            Call getHurt(targIndex(Index), Index)
            If Not gameStarted Then
                Exit Sub
            End If
        End If
        If gameMode = 2 Then
            'for turn-based mode
            intMoves(targIndex(Index)) = intMoves(targIndex(Index)) - 1 'target can't move the turn after being jumped on
            Call getTick(Index) 'get tick
        ElseIf curX(Index) = curX(targIndex(Index)) And curY(Index) = curY(targIndex(Index)) Then
            'for puzzle mode, get jump (if char is on top of target)
            Call getJump(Index, strDir(Index), evalMove(Index, strDir(Index)))
            frameCounter(Index) = 1
        End If
    'for turn based mode, next player-controlled character can move if fast enough, otherwise skip turn (get tick)
    ElseIf gameMode = 2 And Index < 3 Then
        blnMoveOnTick(Index) = False
        If isPlayer(Index + 1) And (intMoves(Index) <= intMoves(Index + 1) Or intMoveCount > 0) Then
            blnMoveOnTick(Index + 1) = True
        Else
            Call getTick(Index + 1)
        End If
    End If
End If
End With
End Sub

'get a tick for turn-based modes
Public Sub getTick(Optional Index As Integer)
Dim loopMin As Integer
Dim loopMax As Integer
If gameMode = 1 Then 'arcade mode uses one tick and executes all movements
    Index = -1
    loopMin = 0
    loopMax = 3
ElseIf gameMode = 2 Then 'turn-based mode uses one index per tick and goes through one character per call
    loopMin = Index
    loopMax = Index
    With frmMain
    'disable character move if not enabled
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
'check for highest move and set it
For moveCheck = 0 To 3
    If .tmrChar(moveCheck).Enabled Then
        If intMoves(moveCheck) > highestMove Then
            highestMove = intMoves(moveCheck)
        End If
    End If
Next moveCheck
'disable move for characters that are not set to move according to higher character speeds, enable otherwise (as long as they are enabled)
For moveCheck2 = loopMin To loopMax
    If intMoves(moveCheck2) < highestMove - intMoveCount Then
        blnMoveOnTick(moveCheck2) = False
        If gameMode = 2 And (intMoves(moveCheck2) = 0 Or (.tmrPow(moveCheck2).Tag = "Speed" And intMoves(moveCheck2) = 1)) Then
            intMoves(moveCheck2) = intMoves(moveCheck2) + 1
        End If
    ElseIf .tmrChar(moveCheck2).Enabled Then
        blnMoveOnTick(moveCheck2) = True
    End If
Next moveCheck2
For setmove = loopMin To loopMax
    'if not bounce jumping in turn-based or in puzzle, characters get next tile values
    If Not blnBounceJump(setmove) Or gameMode = 1 Then
        If isPlayer(setmove) Then 'if setMove is a player, call getJump (direction change and next values set)
            If gameMode = 1 Or blnMoveOnTick(setmove) Then
                Call getJump(setmove, strDir(setmove), evalMove(setmove, strDir(setmove)))
            End If
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(setmove)) Then
            Call cpuAI(setmove) 'if computer player, call AI (direction change and next values)
        End If
    Else 'if bounce jumping in turn-based mode, character set to jump
        blnMoveOnTick(setmove) = True
        Call getJump(setmove, strDir(setmove), evalMove(setmove, strDir(setmove)))
    End If
Next setmove
'in puzzle mode, characters can't move if they are being jumped on and are trying to jump towards an attacking character
If gameMode = 1 Then
    For moveCheck3 = loopMin To loopMax
        If blnMoveOnTick(moveCheck3) Then 'if character can move
            For checkTarg = 0 To 3
                If checkTarg <> moveCheck3 And .tmrChar(checkTarg).Enabled Then 'if index is different and checkTarg char is active
                    'if checkMove3 char and checkTarg char are moving toward eachother
                    If nextX(moveCheck3) = curX(checkTarg) And nextY(moveCheck3) = curY(checkTarg) And curX(moveCheck3) = nextX(checkTarg) And curY(moveCheck3) = nextY(checkTarg) Then
                        If curY(moveCheck3) < curY(checkTarg) Then 'if checkTarg char is lower on map
                            blnMoveOnTick(checkTarg) = False 'checkTarg char can no longer move
                        End If
                    End If
                End If
            Next checkTarg
        End If
    Next moveCheck3
End If
For getMove = loopMin To loopMax 'get movement (or not)
    If .tmrPow(getMove).Tag <> "" Then 'if corresponding pow timer has a tag (power-up activated), call a tick
        Call getPowTick(getMove)
    End If
    If blnMoveOnTick(getMove) Then 'if char can move, set it to start animation
        If Not blnBounceJump(getMove) Then
            frameCounter(getMove) = 1
        Else
            frameCounter(getMove) = 5
        End If
    Else 'if not, set next values to current values
        nextX(getMove) = curX(getMove)
        nextY(getMove) = curY(getMove)
        If gameMode = 2 Then
            blnPlayerMoveable(getMove) = True
            If Index < 3 Then
                If isPlayer(Index + 1) Then
                    'next player-controlled character can move if it is fast enough
                    If intMoves(Index + 1) >= highestMove - intMoveCount Then
                        blnMoveOnTick(Index + 1) = True
                    Else
                        'call a tick (skip char) if not fast enough
                        Call getTick(Index + 1)
                    End If
                Else
                    Call getTick(Index + 1) 'call tick for CPU players
                End If
            End If
        End If
    End If
Next getMove
'if all characters have made their moves
If gameMode = 1 Or (gameMode = 2 And Index = 3) Then
    intMoveCount = intMoveCount + 1
    If highestMove - intMoveCount <= 0 Then 'reset intMoveCount if it, subtracted from the highest move, is below default speed
        intMoveCount = 0
    End If
    If gameMode = 1 Or (gameMode = 2 And intMoveCount = 0) Then
        For T = 0 To tileCount - 1 'check each tile for object, if it has one, call an object timer tick
            If tile(getTileFromInt(True, T), getTileFromInt(False, T)).hasObj Then
                tile(getTileFromInt(True, T), getTileFromInt(False, T)).objTimer = tile(getTileFromInt(True, T), getTileFromInt(False, T)).objTimer + 1
            End If
        Next T
        Call .tmrObjEvent_Timer 'call an object to appear
    End If
    If gameMode = 2 Then 'for turn-based mode, get starting character if player 1 is not guarenteed to be moveable
        If intMoveCount = highestMove - 1 Or highestMove = 1 Then
            blnMoveOnTick(0) = True
        Else
            Dim startChar As Integer
            For startChar = 0 To 3
                If .tmrChar(startChar).Enabled Then
                    If intMoves(startChar) >= highestMove Then
                        If isPlayer(startChar) Then
                            blnMoveOnTick(startChar) = True
                            Exit Sub
                        Else
                            Call getTick(startChar)
                            Exit Sub
                        End If
                        Exit Sub
                    End If
                End If
            Next startChar
        End If
    End If
End If
End With
End Sub
