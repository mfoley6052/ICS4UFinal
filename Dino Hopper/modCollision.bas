Attribute VB_Name = "modCollision"
Public Sub getHurt(ByVal Index As Integer, ByVal enemyIndex As Integer)
'MsgBox ("You got hurt.")
If blnRecover(0) = False Then
    If intLives(0) > 1 Then
        intLives(0) = intLives(0) - 1
    Else
        MsgBox ("GameOver")
        frmStart.Show
        frmMain.Hide
    End If
    If gameMode <> 1 Then
        blnRecover(0) = True
        Call getPowEffect(Index, "Recover")
    End If
    Call refreshLabels(False, True, False)
End If
End Sub
