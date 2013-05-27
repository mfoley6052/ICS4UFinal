Attribute VB_Name = "modCpuAi"

Public Function cpuAI(ByVal index As Integer)
Dim temp As String
temp = randDir
If curX(index) > curX(0) And curY(index) > curY(0) Then
    If evalMove(index, "L") Then
        Call getJump(index, "L", evalMove(index, "L"))
    Else
         Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) > curX(0) And curY(index) < curY(0) Then
    If evalMove(index, "D") Then
        Call getJump(index, "D", evalMove(index, "D"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) < curX(0) And curY(index) > curY(0) Then
    If evalMove(index, "U") Then
         Call getJump(index, "U", evalMove(index, "U"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) < curX(0) And curY(index) < curY(0) Then
    If evalMove(index, "R") Then
        Call getJump(index, "R", evalMove(index, "R"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curY(index) > curY(0) And curX(index) = curX(0) Then
    If evalMove(index, "L") Then
        Call getJump(index, "L", evalMove(index, "L"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curY(index) < curY(0) And curX(index) = curX(0) Then
    If evalMove(index, "R") Then
        Call getJump(index, "R", evalMove(index, "R"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) > curX(0) And curY(index) = curY(0) Then
    If evalMove(index, "D") Then
        Call getJump(index, "D", evalMove(index, "D"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
ElseIf curX(index) < curX(0) And curY(index) = curY(0) Then
    If evalMove(index, "U") Then
        Call getJump(index, "U", evalMove(index, "U"))
    Else
        Call getJump(index, temp, evalMove(index, temp))
    End If
Else
    Call getJump(index, temp, evalMove(index, temp))
End If
End Function
