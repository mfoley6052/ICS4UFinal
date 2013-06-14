Attribute VB_Name = "modCollision"
Public Sub getHurt(ByVal Index As Integer, ByVal enemyIndex As Integer)
'MsgBox ("You got hurt.")
If blnRecover(Index) = False Then
    If isPlayer(Index) Then
        If intLives(Index) > 1 Then
            intLives(Index) = intLives(Index) - 1
        Else
            Call getGameEnd
            frmMain.tmrGameOver.Enabled = True
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
