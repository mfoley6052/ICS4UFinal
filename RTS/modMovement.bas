Attribute VB_Name = "modMovement"
Public Function getJump(ByVal index As Integer, ByVal strDirJ As String)
With frmMain
If .tmrFrame(index).Enabled = False Then
    strDir(index) = strDirJ
    'blnPlayerMoveable = False
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
    .tmrFrame(index).Enabled = True
End If
End With
End Function
Public Sub getJumpComplete(ByVal index As Integer)
Dim pScore As Integer
With frmMain
prevX(index) = curX(index)
prevY(index) = curY(index)
tile(curX(index), curY(index)).hasChar = False
If (index = 0 And blnPlayerMoveable = True) Or index > 0 Then
    curX(index) = nextX(index)
    curY(index) = nextY(index)
End If
If index > 0 Then
    If curX(index) = curX(0) And curY(index) = curY(0) Then
        blnPlayerMoveable = False
        .tmrHurt(index).Enabled = True
    End If
Else
End If
tile(curX(index), curY(index)).hasChar = True
If index = 0 Then
    If tile(curX(0), curY(0)).coinEnabled = True Then
        'play coin sound
        If tile(curX(0), curY(0)).coinType = "Y" Then
            pScore = 100
        ElseIf tile(curX(0), curY(0)).coinType = "R" Then
            pScore = 250
        ElseIf tile(curX(0), curY(0)).coinType = "B" Then
            pScore = 500
        End If
    Else
        pScore = 10
    End If
    addScore (pScore)
End If
If tile(curX(index), curY(index)).coinEnabled = True Then
    tile(curX(index), curY(index)).coinEnabled = False
    tile(curX(index), curY(index)).coinTimer = 0
    coinTileCount = coinTileCount - 1
End If
blnClearPrevTile(index) = True
strState(index) = "I"
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
