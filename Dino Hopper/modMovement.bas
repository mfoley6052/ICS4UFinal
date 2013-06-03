Attribute VB_Name = "modMovement"
Public Sub charAction(ByVal index As Integer, nextTile As terrain)
If frameCounter(index) = 1 Then
    strState(index) = "C"
ElseIf frameCounter(index) = 5 Then
    strState(index) = "J"
End If
If gameMode <> 0 Then
    targIndex(index) = -1
End If
If Not blnEdgeJump(index) Then
    If gameMode = 0 Then
        Call getCharJumpAnim(index, frameCounter(index), tile(curX(index), curY(index)), nextTile.X, nextTile.Y)
    ElseIf gameMode <> 0 Then
        For charIndex = 0 To 3
            With frmMain
            If charIndex <> index And .tmrChar(charIndex).Enabled Then
                If nextX(index) = nextX(charIndex) And nextY(index) = nextY(charIndex) Then
                    If curY(index) < curY(charIndex) Then
                        targIndex(index) = charIndex
                        Call getCharJumpAnim(index, frameCounter(index), tile(curX(index), curY(index)), nextTile.X, spriteY(targIndex(index)))
                    ElseIf curY(index) = curY(charIndex) Then
                        If frameCounter(index) = 9 Then
                            frameProg(index) = -1
                        End If
                        Call getCharJumpAnim(index, frameCounter(index), tile(curX(index), curY(index)), nextTile.X, nextTile.Y)
                    End If
                End If
            End If
            End With
        Next charIndex
    End If
    If targIndex(index) < 0 Then
        Call getCharJumpAnim(index, frameCounter(index), tile(curX(index), curY(index)), tile(nextX(index), nextY(index)).X, tile(nextX(index), nextY(index)).Y)
    End If
    Call PaintCharSprite(index, spriteX(index), spriteY(index))
    If frameCounter(index) = 10 Then
        strState(index) = "I"
    End If
    If (frameCounter(index) = frameLimit(index)) Or (frameProg(index) < 0 And frameCounter(index) <= 5) Then
        If gameMode = 0 Then
            Call getJumpComplete(index)
        ElseIf checkTickComplete Then
            Call getTickComplete
        End If
    Else
        frameCounter(index) = frameCounter(index) + frameProg(index)
    End If
Else 'jump off edge
    Dim curTile As terrain
    curTile = tile(curX(index), curY(index))
    Dim altTile As terrain
    If strDir(index) = "L" Then
        Call getCharJumpAnim(index, frameCounter(index), tile(curX(index), curY(index)), tile(curX(index), curY(index)).X - 50, tile(curX(index), curY(index)).Y - 75)
        If curY(index) = 0 Then
            If spriteY(index) < curTile.Y + 50 Or (spriteX(index) > -35 And spriteX(index) < tile(mapWidth - 1, 0).X + 135 And spriteY(index) < curTile.Y + 25) Then
                Call PaintCharSprite(index, spriteX(index), spriteY(index))
            End If
            Call clearTile(curTile, True, index, "CharTop+X-Y")
            If curX(index) > 0 Then
                altTile = tile(curTile.Xc - 1, curTile.Yc)
                If altTile.hasObj Then
                    Call PaintObj(altTile.objType(0), altTile.objType(1), altTile.objFrame, altTile.Xc, altTile.Yc, False)
                End If
            End If
        ElseIf curX(index) = 0 And curY(index) > 0 Then
            Call clearTile(curTile, True, index, "CharSide-X-Y")
            Call PaintCharSprite(index, spriteX(index), spriteY(index))
        End If
    ElseIf strDir(index) = "U" Then
        Call getCharJumpAnim(index, frameCounter(index), curTile, curTile.X + 50, curTile.Y - 75)
        If curY(index) = 0 Then
            If spriteY(index) < curTile.Y + 50 Or (spriteX(index) > -35 And spriteX(index) < tile(mapWidth - 1, 0).X + 135 And spriteY(index) < curTile.Y + 25) Then
                Call PaintCharSprite(index, spriteX(index), spriteY(index))
            End If
            Call clearTile(curTile, True, index, "CharTop+X-Y")
            If curX(index) < mapWidth - 1 Then
                altTile = tile(curTile.Xc + 1, curTile.Yc)
                If altTile.hasObj Then
                    Call PaintObj(altTile.objType(0), altTile.objType(1), altTile.objFrame, altTile.Xc, altTile.Yc, False)
                End If
            End If
        ElseIf curX(index) = mapWidth And curY(index) > 0 Then
            Call clearTile(curTile, True, index, "CharSide+X-Y")
            Call PaintCharSprite(index, spriteX(index), spriteY(index))
        End If
    ElseIf strDir(index) = "R" Then
        Call getCharJumpAnim(index, frameCounter(index), curTile, curTile.X + 50, curTile.Y + 75)
        If curX(index) = mapWidth And curY(index) > 0 Then
            Call clearTile(curTile, True, index, "CharSide+X+Y")
        ElseIf curY(index) = mapHeight - 1 Then
            Call clearTile(curTile, True, index, "CharBottom+X+Y")
        End If
        Call PaintCharSprite(index, spriteX(index), spriteY(index))
    ElseIf strDir(index) = "D" Then
        Call getCharJumpAnim(index, frameCounter(index), curTile, curTile.X - 50, curTile.Y + 75)
        If curX(index) = 0 And (curY(index) > 0 And curY(index) < mapHeight - 1) Then
            Call clearTile(curTile, True, index, "CharSide-X+Y")
        ElseIf curY(index) = mapHeight - 1 Then
            Call clearTile(curTile, True, index, "CharBottom-X+Y")
        End If
        Call PaintCharSprite(index, spriteX(index), spriteY(index))
    End If
    If frameCounter(index) = frameLimit(index) * 1.6 Then
        spriteX(index) = tile(curX(index), curY(index)).X + 25
        spriteY(index) = tile(curX(index), curY(index)).Y - 15
        If gameMode = 0 Then
            Call getJumpComplete(index)
        ElseIf checkTickComplete Then
            Call getTickComplete
        End If
    Else
        frameCounter(index) = frameCounter(index) + frameProg(index)
    End If
End If
End Sub

Public Function getJump(ByVal index As Integer, ByVal strDirJ As String, ByVal blnTile As Boolean)
With frmMain
    If frameCounter(index) = 0 Then
        strDir(index) = strDirJ
        blnEdgeJump(index) = Not blnTile
        If blnPlayerMoveable(index) Then
            blnPlayerMoveable(index) = False
            If blnTile Then
                If strDir(index) = "L" Then
                    'if y row is odd
                    If oddRow(curY(index)) Then
                        nextX(index) = curX(index) - 1
                    Else
                        nextX(index) = curX(index)
                    End If
                    nextY(index) = curY(index) - 1
                ElseIf strDir(index) = "U" Then
                    'if y row is even
                    If Not oddRow(curY(index)) Then
                        nextX(index) = curX(index) + 1
                    Else
                        nextX(index) = curX(index)
                    End If
                    nextY(index) = curY(index) - 1
                ElseIf strDir(index) = "R" Then
                    'if y row is even
                    If Not oddRow(curY(index)) Then
                        nextX(index) = curX(index) + 1
                    Else
                        nextX(index) = curX(index)
                    End If
                    nextY(index) = curY(index) + 1
                ElseIf strDir(index) = "D" Then
                    'if y row is odd
                    If oddRow(curY(index)) Then
                        nextX(index) = curX(index) - 1
                    Else
                        nextX(index) = curX(index)
                    End If
                    nextY(index) = curY(index) + 1
                End If
            ElseIf Not blnRecover(index) Then
                Call setCharRespawn(index, tile(curX(index), curY(index)))
            End If
            If gameMode <> 1 Then
                frameCounter(index) = 1
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
Public Sub getJumpComplete(ByVal index As Integer)
Dim pScore As Integer
Dim q As Integer
With frmMain
prevX(index) = curX(index)
prevY(index) = curY(index)

strState(index) = "I"
frameCounter(index) = 0
blnPlayerMoveable(index) = True

If Not blnBounceJump(index) Then
    tile(curX(index), curY(index)).hasChar = False
ElseIf blnBounceJump(index) And targIndex(index) >= 0 Then
    blnBounceJump(index) = False
End If
If blnEdgeJump(index) Then
    tile(curX(index), curY(index)).hasChar = False
    Call getHurt(index, index)
End If

curX(index) = nextX(index)
curY(index) = nextY(index)

Dim inputTile As terrain
inputTile = tile(curX(index), curY(index))

tile(curX(index), curY(index)).hasChar = True
If isPlayer(index) Then
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
                intMulti(index) = intMulti(index) + 1
                pScore = 1000
                blnMulti = True
            ElseIf inputTile.objType(1) = "G" Then
                intLives(index) = intLives(index) + 1
                blnLives = True
            End If
        End If
    ElseIf blnBounceJump(index) Then 'jumping on enemy
        pScore = 1000
    End If
    If inputTile.terType = "G" Then
        pScore = pScore + 25
    End If
    Call addScore(index, pScore)
    Call refreshLabels(True, blnLives, blnMulti)
End If
If inputTile.hasObj Then
    If inputTile.objType(0) <> "Terrain" Then
        strState(index) = "I"
        Call killObj(tile(curX(index), curY(index)))
        If inputTile.objType(0) = "Pow" Then
            Call getPowEffect(index, inputTile.objType(1))
        End If
    Else
        If inputTile.objType(1) = "I" Then
            Call getJump(index, strDir(index), evalMove(index, strDir(index)))
        End If
    End If
End If
blnClearPrevTile(index) = True
.lblTest.Caption = "(" & curX(0) & ", " & curY(0) & ") (" & nextX(0) & ", " & nextY(0) & ")"
.lblTest2.Caption = oddRow(curY(0)) & ", " & oddRow(nextY(0))
End With
End Sub

Public Function evalMove(ByVal index As Integer, ByVal strDirMove As String) As Boolean
If strDirMove = "L" Then
    If oddRow(curY(index)) Then
        If curX(index) = 0 Then
            evalMove = False
        ElseIf curY(index) > 0 Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY(index) > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "U" Then
    If oddRow(curY(index)) Then
        If curX(index) < mapWidth Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curX(index) < (mapWidth) And curY(index) > 0 Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "R" Then
    If oddRow(curY(index)) Then
        If curX(index) = mapWidth Then
            evalMove = False
        ElseIf curY(index) > 0 Then
            evalMove = True
        End If
    ElseIf curY(index) < (mapWidth - 1) And curX(index) < mapWidth Then
        evalMove = True
    Else
        evalMove = False
    End If
ElseIf strDirMove = "D" Then
    If oddRow(curY(index)) Then
        If curX(index) > 0 And curY(index) < mapHeight Then
            evalMove = True
        Else
            evalMove = False
        End If
    ElseIf curY(index) < (mapHeight - 1) Then
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

Public Sub getTickComplete()
With frmMain
For getNext = 0 To 3
    If .tmrChar(getNext).Enabled And blnMoveOnTick(getNext) Then
        If frameProg(getNext) < 0 Then
            frameProg(getNext) = 1
            nextX(getNext) = curX(getNext)
            nextY(getNext) = curY(getNext)
        End If
        Call getJumpComplete(getNext)
        If Not blnBounceJump(getNext) And targIndex(getNext) >= 0 Then
            blnBounceJump(getNext) = True
            If Not blnRecover(targIndex(getNext)) Then
                Call getHurt(targIndex(getNext), getNext)
            End If
            Call getJump(getNext, strDir(getNext), evalMove(getNext, strDir(getNext)))
            frameCounter(getNext) = 1
        End If
    End If
Next getNext
End With
End Sub

Public Sub getTick()
With frmMain
Static intMoveCount As Integer 'counter for moves to assign which characters move
Dim highestMove As Integer 'highest move count out of all characters
highestMove = 1 'set highest move to default
For moveCheck = 0 To 3
    If .tmrChar(moveCheck).Enabled Then 'check for highest move and set it
        If intMoves(moveCheck) > highestMove Then
            highestMove = intMoves(moveCheck)
        End If
    End If
Next moveCheck
For moveCheck2 = 0 To 3 'disable blnMove for characters that are not set to move according to higher character speeds
    If intMoves(moveCheck2) < highestMove - intMoveCount Then
        blnMoveOnTick(moveCheck2) = False
    ElseIf .tmrChar(moveCheck2).Enabled Then
        blnMoveOnTick(moveCheck2) = True
    End If
Next moveCheck2
For setMove = 0 To 3
    If isPlayer(setMove) Then 'if setMove is a player, call getJump (direction change and next values set)
        Call getJump(setMove, strDir(setMove), evalMove(setMove, strDir(setMove)))
    Else 'if computer player, call AI (direction change and next values)
        Call cpuAI(setMove)
    End If
Next setMove
For movecheck3 = 0 To 3
    If blnMoveOnTick(movecheck3) Then 'if character can move
        For checkTarg = 0 To 3
            If checkTarg <> movecheck3 And .tmrChar(checkTarg).Enabled Then 'if index is different and checkTarg char is active
                'if checkMove3 char and checkTarg char are moving toward eachother
                If nextX(movecheck3) = curX(checkTarg) And nextY(movecheck3) = curY(checkTarg) And curX(movecheck3) = nextX(checkTarg) And curY(movecheck3) = nextY(checkTarg) Then
                    If curY(movecheck3) < curY(checkTarg) Then 'if checkTarg char is lower on map
                        blnMoveOnTick(checkTarg) = False 'checkTarg char can no longer move
                    End If
                End If
            End If
        Next checkTarg
    End If
Next movecheck3
intMoveCount = intMoveCount + 1 'move count increases by one
If highestMove - intMoveCount = 1 Then 'reset intMoveCount if it, subtracted from the highest move, is the lowest speed
    intMoveCount = 0
End If
For getMove = 0 To 3 'get movement (or not)
    If .tmrPow(getMove).Tag <> "" Then 'if corresponding pow timer has a tag (power-up activated), call a tick
        Call getPowTick(getMove)
    End If
    If blnMoveOnTick(getMove) Then 'if char can move, set it to start animation
        frameCounter(getMove) = 1
    Else 'if not, set next values to current values
        nextX(getMove) = curX(getMove)
        nextY(getMove) = curY(getMove)
    End If
Next getMove
For T = 0 To tileCount - 1 'check each tile for object, if it has one, call an object timer tick
    If tile(getTileFromInt(True, T), getTileFromInt(False, T)).hasObj Then
        tile(getTileFromInt(True, T), getTileFromInt(False, T)).objTimer = tile(getTileFromInt(True, T), getTileFromInt(False, T)).objTimer + 1
    End If
Next T
Call .tmrObjEvent_Timer 'call an object to appear
End With
End Sub
