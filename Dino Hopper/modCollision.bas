Attribute VB_Name = "modCollision"
Public Sub getHurt(ByVal enemyIndex As Integer)
'MsgBox ("You got hurt.")
If intLives(0) > 1 Then
    intLives(0) = intLives(0) - 1
Else
    MsgBox ("GameOver")
    frmStart.Show
    frmMain.Hide
End If
Call refreshLabels(False, True, False)
End Sub
