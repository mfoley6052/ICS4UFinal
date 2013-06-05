Attribute VB_Name = "modCollision"
Public Sub getHurt(ByVal index As Integer, ByVal enemyIndex As Integer)
'MsgBox ("You got hurt.")
If blnRecover(index) = False Then
    If isPlayer(index) Then
        If intLives(index) > 1 Then
            intLives(index) = intLives(index) - 1
        Else
            MsgBox ("Game Over")
            frmStart.Show
            frmMain.Hide
        End If
        Call refreshLabels(False, True, False)
    End If
    If gameMode = 0 Then
        blnRecover(index) = True
        Call getPowEffect(index, "Recover")
    End If
    blnPlayerMoveable(index) = True
End If
End Sub
