Attribute VB_Name = "modMovement"
Public Function getJump(ByVal index As Integer, ByVal strDirJ As String, ByVal blnTile As Boolean)
With frmMain
    If frameCounter(index) = 0 Then
        strDir(index) = strDirJ
        blnEdgeJump(index) = Not blnTile
        If blnTile And blnPlayerMoveable(index) Then
            blnPlayerMoveable(index) = False
            If strDir(index) = "L" Then
                'if y row is odd
                If oddRow(curY(index)) Then
                    nextX(index) = curX(index) - 1
                End If
                nextY(index) = curY(index) - 1
            ElseIf strDir(index) = "U" Then
                'if y row is even
                If Not oddRow(curY(index)) Then
                    nextX(index) = curX(index) + 1
                End If
                nextY(index) = curY(index) - 1
            ElseIf strDir(index) = "R" Then
                'if y row is even
                If Not oddRow(curY(index)) Then
                    nextX(index) = curX(index) + 1
                End If
                nextY(index) = curY(index) + 1
            ElseIf strDir(index) = "D" Then
                'if y row is odd
                If oddRow(curY(index)) Then
                    nextX(index) = curX(index) - 1
                End If
                nextY(index) = curY(index) + 1
            End If
        End If
        If gameMode = 1 Then
            If tile(nextX(index), nextY(index)).hasChar Then
                If curY(index) < nextY(index) Then
                    frameCounter(index) = 1
                Else
                    nextX(index) = curX(index)
                    nextY(index) = curY(index)
                End If
            Else
                frameCounter(index) = 1
            End If
            If index = 0 Then
                Call getTick
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
    randDir = "R"
ElseIf temp = 3 Then
    randDir = "D"
Else
    randDir = "U"
End If
End Function
Public Sub getJumpComplete(ByVal index As Integer, ByVal blnBounce As Boolean)
Dim pScore As Integer
Dim q As Integer
With frmMain
prevX(index) = curX(index)
prevY(index) = curY(index)
If Not blnBounce Then
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
    ElseIf blnBounce Then 'jumping on enemy
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


Public Function evalMove(index As Integer, ByVal strDirMove As String) As Boolean
If strDirMove = "L" Then
    If oddRow(curY(index)) Then
        If curX(index) = 0 Then
            evalMove = False
        ElseIf curY(index) < mapWidth Then
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
Static intMoveCount As Integer
intMoveCount = intMoveCount + 1
If intMoveCount >= intMoves(0) Then
    intMoveCount = 0
    For C = 1 To 3
        If nextX(0) = curX(C) And nextY(0) = curY(C) And frameCounter(0) > 0 Then
            blnPlayerMoveable(C) = False
        End If
        For m = 1 To intMoves(C)
            If .tmrChar(C).Enabled Then
                Call cpuAI(C)
            End If
        Next m
        If .tmrPow(C).Tag <> "" Then
            Call getPowTick(C)
        End If
    Next C
End If
If .tmrPow(0).Tag <> "" Then
    Call getPowTick(0)
End If
For t = 0 To tileCount - 1
    If tile(getTileFromInt(True, t), getTileFromInt(False, t)).hasObj Then
        tile(getTileFromInt(True, t), getTileFromInt(False, t)).objTimer = tile(getTileFromInt(True, t), getTileFromInt(False, t)).objTimer + 1
    End If
Next t
Call .tmrObjEvent_Timer
End With
End Sub
