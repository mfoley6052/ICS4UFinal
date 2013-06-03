Attribute VB_Name = "modCpuAi"

Public Function cpuAI(ByVal index As Integer)
Static counter As Long
Dim temp As String
Dim target As Integer
counter = counter + 1
temp = randDir
For X = 0 To 3
    If isPlayer(X) Then
        If Abs(curX(index) - curX(X)) + Abs(curY(index) - curY(X)) < Abs(curX(index) - curX(target)) + Abs(curY(index) - curY(target)) Then
            target = X
        End If
    End If
Next X
If curX(index) > curX(target) And curY(index) > curY(target) Then
    If evalMove(index, "L") Then
        Call getJump(index, "L", evalMove(index, "L"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) > curX(target) And curY(index) < curY(target) Then
    If evalMove(index, "D") Then
        Call getJump(index, "D", evalMove(index, "D"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) < curX(target) And curY(index) > curY(target) Then
    If evalMove(index, "U") Then
        Call getJump(index, "U", evalMove(index, "U"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) < curX(target) And curY(index) < curY(target) Then
    If evalMove(index, "R") Then
        Call getJump(index, "R", evalMove(index, "R"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curY(index) > curY(target) And curX(index) = curX(target) Then
    If evalMove(index, "L") Then
        Call getJump(index, "L", evalMove(index, "L"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curY(index) < curY(target) And curX(index) = curX(target) Then
    If evalMove(index, "R") Then
        Call getJump(index, "R", evalMove(index, "R"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) > curX(target) And curY(index) = curY(target) Then
    If evalMove(index, "D") Then
        Call getJump(index, "D", evalMove(index, "D"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) < curX(target) And curY(index) = curY(target) Then
    If evalMove(index, "U") Then
        Call getJump(index, "U", evalMove(index, "U"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
Else
    Call getJump(index, temp, evalMove(index, temp))
End If
If gameMode <> 1 Then
    If counter Mod 5 = 0 And frmMain.tmrChar(index).Interval > 200 Then
        frmMain.tmrChar(index).Interval = Int(frmMain.tmrChar(index).Interval - 50)
    End If
Else
    If counter Mod 30 = 0 Then
        intMoves(index) = 2
    End If
End If
End Function
