Attribute VB_Name = "modMovement"
Public Function getJump(ByVal Index As Integer, ByVal strDirJ As String)
With frmMain
    If .tmrFrame(Index).Enabled = False Then
        strDir(Index) = strDirJ
        'blnPlayerMoveable = False
        If strDir(Index) = "L" Then
            'if y row is odd
            If oddRow(curY(Index)) Then
                nextX(Index) = curX(Index) - 1
            End If
            nextY(Index) = curY(Index) - 1
        ElseIf strDir(Index) = "U" Then
            'if y row is even
            If Not oddRow(curY(Index)) Then
                nextX(Index) = curX(Index) + 1
            End If
            nextY(Index) = curY(Index) - 1
        ElseIf strDir(Index) = "R" Then
            'if y row is even
            If Not oddRow(curY(Index)) Then
                nextX(Index) = curX(Index) + 1
            End If
            nextY(Index) = curY(Index) + 1
        ElseIf strDir(Index) = "D" Then
            'if y row is odd
            If oddRow(curY(Index)) Then
                nextX(Index) = curX(Index) - 1
            End If
            nextY(Index) = curY(Index) + 1
        End If
        .tmrFrame(Index).Enabled = True
    End If
End With
End Function
Public Sub getJumpComplete(ByVal Index As Integer)
Dim pScore As Integer
Dim q As Integer
With frmMain
prevX(Index) = curX(Index)
prevY(Index) = curY(Index)
'For ai pathing
'If index = 0 Then
'    Do Until q = 2
'        q = UBound(pathStep)
'        pathStep(q).x = pathStep(q - 1).x
'        pathStep(q).y = pathStep(q - 1).y
'
'        q = q - 1
'    Loop
'    pathStep(1).x = curX(0)
'    pathStep(1).y = curY(0)
'End If
prevX(Index) = curX(Index)
prevY(Index) = curY(Index)
tile(curX(Index), curY(Index)).hasChar = False
If (Index = 0 And blnPlayerMoveable = True) Or Index > 0 Then
    curX(Index) = nextX(Index)
    curY(Index) = nextY(Index)
End If

If Index > 0 Then
    If curX(Index) = curX(0) And curY(Index) = curY(0) Then
        blnPlayerMoveable = False
        .tmrHurt(Index).Enabled = True
    End If
Else
End If
tile(curX(Index), curY(Index)).hasChar = True
If Index = 0 Then
    If tile(curX(0), curY(0)).hasObj Then
        Dim blnMulti As Boolean
        If tile(curX(0), curY(0)).objType(0) = "Coin" Then
            'play coin sound
            If tile(curX(0), curY(0)).objType(1) = "Y" Then
                pScore = 100
            ElseIf tile(curX(0), curY(0)).objType(1) = "R" Then
                pScore = 250
            ElseIf tile(curX(0), curY(0)).objType(1) = "B" Then
                pScore = 500
            End If
        ElseIf tile(curX(0), curY(0)).objType(0) = "Pow" Then
            'play scare power-up sound
            If tile(curX(0), curY(0)).objType(1) = "Scare" Then
                pScore = 200
            End If
        ElseIf tile(curX(0), curY(0)).objType(0) = "Egg" Then
            intMulti = tile(curX(0), curY(0)).objType(1)
            pScore = 1000
            blnMulti = True
        End If
    Else
        pScore = 10
    End If
    Call addScore(pScore)
    Call refreshLabels(True, False, blnMulti)
End If
If tile(curX(Index), curY(Index)).hasObj = True Then
    Call killObj(tile(curX(Index), curY(Index)))
End If
blnClearPrevTile(Index) = True
strState(Index) = "I"
.lblTest.Caption = "(" & curX(0) & ", " & curY(0) & ") (" & nextX(0) & ", " & nextY(0) & ")"
.lblTest2.Caption = oddRow(curY(0)) & ", " & oddRow(nextY(0))
End With
End Sub


Public Function evalMove(Index As Integer, ByVal strDirMove As String) As Boolean
If strDirMove = "L" Then
    If oddRow(curY(Index)) Then
        If curX(Index) = 0 Then
            evalMove = False
        ElseIf curY(Index) < mapWidth Then
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
