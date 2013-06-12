Attribute VB_Name = "modCollision"
Public Sub getHurt(ByVal Index As Integer, ByVal enemyIndex As Integer)
'MsgBox ("You got hurt.")
If blnRecover(Index) = False Then
    If isPlayer(Index) Then
        If intLives(Index) > 1 Then
            intLives(Index) = intLives(Index) - 1
        Else
            MsgBox ("Game Over")
            If numPlayers = 1 And numCPU = 0 Then
                playMode = "SOLO"
            ElseIf numPlayers = 1 Then
                playMode = "SP"
            Else
                playMode = "MP"
            End If
            For x = 0 To numPlayers
                Call WriteScore(InputBox("Please enter your name(3 character max): ", "HiScore Entry", "RST"), intScore(x))
            Next x
            frmStart.Show
            Call getGameEnd
            Exit Sub
        End If
        Call refreshLabels(Index, False, True, False)
    End If
    If gameMode = 0 Then
        If isPlayer(Index) Then
            blnPlayerMoveable(Index) = False
        Else
            frmMain.tmrCPUMove(Index).Enabled = False
        End If
        frmMain.tmrStun(Index).Enabled = True
    End If
    blnRecover(Index) = True
    frmMain.tmrChar(Index).Tag = "Recover"
    Call getPowEffect(Index, "Recover")
    blnPlayerMoveable(Index) = True
End If
End Sub
