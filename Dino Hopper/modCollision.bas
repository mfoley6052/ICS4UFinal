Attribute VB_Name = "modCollision"
Public Sub getHurt(ByVal enemyIndex As Integer)
MsgBox ("You got hurt.")
intLives(0) = intLives(0) - 1
Call refreshLabels(False, True, False)
End Sub
