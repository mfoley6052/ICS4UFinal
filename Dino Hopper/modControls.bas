Attribute VB_Name = "modControls"
Public Sub keyHandler(ByVal keyCode As Integer, ByVal Shift As Integer)
Dim playerIndex As Integer
If keyCode = key(0) Or keyCode = key(1) Or keyCode = key(2) Or keyCode = key(3) Or keyCode = key(4) Then
    playerIndex = 0
ElseIf keyCode = key(5) Or keyCode = key(6) Or keyCode = key(7) Or keyCode = key(8) Or keyCode = key(9) Then
    playerIndex = 1
ElseIf keyCode = key(10) Or keyCode = key(11) Or keyCode = key(12) Or keyCode = key(13) Or keyCode = key(14) Then
    playerIndex = 2
End If
If blnPlayerMoveable(playerIndex) = True And frameCounter(playerIndex) = 0 Then
    If keyCode = key(0 + (5 * playerIndex)) Then 'Left
        If gameMode = 0 Then
            Call getJump(playerIndex, "L", evalMove(playerIndex, "L"))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            If strDir(playerIndex) = "L" Then
                Call getTick(playerIndex)
            Else
                strDir(playerIndex) = "L"
            End If
        End If
    ElseIf keyCode = key(1 + (5 * playerIndex)) Then 'Up
        If gameMode = 0 Then
            Call getJump(playerIndex, "U", evalMove(playerIndex, "U"))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            If strDir(playerIndex) = "U" Then
                Call getTick(playerIndex)
            Else
                strDir(playerIndex) = "U"
            End If
        End If
    ElseIf keyCode = key(2 + (5 * playerIndex)) Then 'Right
        If gameMode = 0 Then
            Call getJump(playerIndex, "R", evalMove(playerIndex, "R"))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            If strDir(playerIndex) = "R" Then
                Call getTick(playerIndex)
            Else
                strDir(playerIndex) = "R"
            End If
        End If
    ElseIf keyCode = key(3 + (5 * playerIndex)) Then 'Down
        If gameMode = 0 Then
            Call getJump(playerIndex, "D", evalMove(playerIndex, "D"))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            If strDir(playerIndex) = "D" Then
                Call getTick(playerIndex)
            Else
                strDir(playerIndex) = "D"
            End If
        End If
    ElseIf keyCode = key(4 + (5 * playerIndex)) Then 'Action
        If gameMode = 0 Then
            Call getJump(playerIndex, strDir(playerIndex), evalMove(playerIndex, strDir(0)))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            Call getTick(playerIndex)
        End If
    End If
End If
End Sub
