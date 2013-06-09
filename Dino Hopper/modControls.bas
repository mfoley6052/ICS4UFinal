Attribute VB_Name = "modControls"
Public Sub keyHandler(ByVal KeyCode As Integer, ByVal Shift As Integer)
Dim playerIndex As Integer
If KeyCode = key(0) Or KeyCode = key(1) Or KeyCode = key(2) Or KeyCode = key(3) Or KeyCode = key(4) Then
    playerIndex = 0
ElseIf KeyCode = key(5) Or KeyCode = key(6) Or KeyCode = key(7) Or KeyCode = key(8) Or KeyCode = key(9) Then
    playerIndex = 1
ElseIf KeyCode = key(10) Or KeyCode = key(11) Or KeyCode = key(12) Or KeyCode = key(13) Or KeyCode = key(14) Then
    playerIndex = 2
End If
If blnPlayerMoveable(playerIndex) = True And frameCounter(playerIndex) = 0 Then
    If KeyCode = key(0 + (5 * playerIndex)) Then 'Left
        If gameMode = 0 Then
            Call getJump(playerIndex, "L", evalMove(playerIndex, "L"))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            If strDir(playerIndex) = "L" Then
                Call getTick(playerIndex)
            Else
                strDir(playerIndex) = "L"
            End If
        End If
    ElseIf KeyCode = key(1 + (5 * playerIndex)) Then 'Up
        If gameMode = 0 Then
            Call getJump(playerIndex, "U", evalMove(playerIndex, "U"))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            If strDir(playerIndex) = "U" Then
                Call getTick(playerIndex)
            Else
                strDir(playerIndex) = "U"
            End If
        End If
    ElseIf KeyCode = key(2 + (5 * playerIndex)) Then 'Right
        If gameMode = 0 Then
            Call getJump(playerIndex, "R", evalMove(playerIndex, "R"))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            If strDir(playerIndex) = "R" Then
                Call getTick(playerIndex)
            Else
                strDir(playerIndex) = "R"
            End If
        End If
    ElseIf KeyCode = key(3 + (5 * playerIndex)) Then 'Down
        If gameMode = 0 Then
            Call getJump(playerIndex, "D", evalMove(playerIndex, "D"))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            If strDir(playerIndex) = "D" Then
                Call getTick(playerIndex)
            Else
                strDir(playerIndex) = "D"
            End If
        End If
    ElseIf KeyCode = key(4 + (5 * playerIndex)) Then 'Action
        If gameMode = 0 Then
            Call getJump(playerIndex, strDir(playerIndex), evalMove(playerIndex, strDir(0)))
        ElseIf gameMode = 1 Or (gameMode = 2 And blnMoveOnTick(playerIndex)) Then
            Call getTick(playerIndex)
        End If
    End If
End If
End Sub

Public Sub keyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 123 Then 'F12
    If frmDbg.Visible = False Then
        frmDbg.Visible = True
        frmDbg.Show
        frmGUI.SetFocus
    Else
        frmDbg.Visible = False
        frmDbg.Hide
    End If
ElseIf KeyCode = 122 Then 'F11
        If frmSettings.Visible = False Then
        frmSettings.Visible = True
    Else
        frmSettings.Visible = False
    End If
ElseIf KeyCode = 27 Then 'Esc
    frmPause.Show
End If
End Sub
