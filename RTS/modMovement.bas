Attribute VB_Name = "modMovement"
Public Function getJump(ByVal index As Integer, ByVal strDirJ As String)
If tmrFrame(index).Enabled = False Then
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
    tmrFrame(index).Enabled = True
End If
End Function

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
