Attribute VB_Name = "modMovement"
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
If Not blnBounceJump(index) Then
    tile(curX(index), curY(index)).hasChar = False
End If
If (index = 0 And blnPlayerMoveable(index)) Or index > 0 Then
    curX(index) = nextX(index)
    curY(index) = nextY(index)
End If
Dim inputTile As terrain
inputTile = tile(curX(index), curY(index))
tile(curX(index), curY(index)).hasChar = True
If index = 0 Then
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

Public Sub getTick()
With frmMain
Static intMoveCount As Integer 'counter for moves to assign which characters move
Dim blnMove(0 To 3) As Boolean 'boolean for characters moving on this tick
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
        blnMove(moveCheck2) = False
    ElseIf .tmrChar(moveCheck2).Enabled Then
        blnMove(moveCheck2) = True
    End If
Next moveCheck2
For setMove = 0 To 3
    If isPlayer(setMove) Then 'if setMove is a player, call getJump (direction change and next values set)
        Call getJump(setMove, strDir(setMove), evalMove(setMove, strDir(setMove)))
    Else 'if computer player, call AI (direction change and next values)
        Call cpuAI(setMove)
    End If
Next setMove
For moveCheck3 = 0 To 3
    If blnMove(moveCheck3) Then 'if character can move
        For checkTarg = 0 To 3
            If checkTarg <> moveCheck3 And .tmrChar(checkTarg).Enabled Then 'if index is different and checkTarg char is active
                'if checkMove3 char and checkTarg char are moving toward eachother
                If nextX(moveCheck3) = curX(checkTarg) And nextY(moveCheck3) = curY(checkTarg) And curX(moveCheck3) = nextX(checkTarg) And curY(moveCheck3) = nextY(checkTarg) Then
                    If curY(moveCheck3) < curY(checkTarg) Then 'if checkTarg char is lower on map
                        blnMove(checkTarg) = False 'checkTarg char can no longer move
                    End If
                End If
            End If
        Next checkTarg
    End If
Next moveCheck3
intMoveCount = intMoveCount + 1 'move count increases by one
If highestMove - intMoveCount = 1 Then 'reset intMoveCount if it, subtracted from the highest move, is the lowest speed
    intMoveCount = 0
End If
For getMove = 0 To 3 'get movement (or not)
    If .tmrPow(getMove).Tag <> "" Then 'if corresponding pow timer has a tag (power-up activated), call a tick
        Call getPowTick(getMove)
    End If
    If blnMove(getMove) Then 'if char can move, set it to start animation
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
