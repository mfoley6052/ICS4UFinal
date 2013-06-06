Attribute VB_Name = "modCollision"
Public Sub getHurt(ByVal Index As Integer, ByVal enemyIndex As Integer)
'MsgBox ("You got hurt.")
If blnRecover(Index) = False Then
    If isPlayer(Index) Then
        If intLives(Index) > 1 Then
            intLives(Index) = intLives(Index) - 1
        Else
            MsgBox ("Game Over")
            frmStart.Show
            frmMain.Hide
        End If
        Call refreshLabels(False, True, False)
    End If
    If gameMode = 0 Then
        If isPlayer(Index) Then
            blnPlayerMoveable(Index) = False
        Else
            frmMain.tmrCPUMove(Index).Enabled = False
        End If
        frmMain.tmrStun(Index).Enabled = True
        blnRecover(Index) = True
        Call getPowEffect(Index, "Recover")
    End If
    blnPlayerMoveable(Index) = True
End If
End Sub
