Attribute VB_Name = "modCpuAi"

Public Function cpuAI(ByVal index As Integer)
If curX(index) > curX(0) And curY(index) > curY(0) Then
    Call getJump(index, "L", evalMove(index, "L"))
ElseIf curX(index) > curX(0) And curY(index) < curY(0) Then
    Call getJump(index, "D", evalMove(index, "D"))
ElseIf curX(index) < curX(0) And curY(index) > curY(0) Then
    Call getJump(index, "U", evalMove(index, "U"))
ElseIf curX(index) < curX(0) And curY(index) < curY(0) Then
    Call getJump(index, "R", evalMove(index, "R"))
ElseIf curY(index) > curY(0) And curX(index) = curX(0) Then
    Call getJump(index, "L", evalMove(index, "U"))
ElseIf curY(index) < curY(0) And curX(index) = curX(0) Then
    Call getJump(index, "R", evalMove(index, "D"))
ElseIf curX(index) > curX(0) And curY(index) = curY(0) Then
    Call getJump(index, "D", evalMove(index, "D"))
ElseIf curX(index) < curX(0) And curY(index) = curY(0) Then
    Call getJump(index, "U", evalMove(index, "D"))
Else
    Dim temp As String
    temp = randDir
    Call getJump(index, temp, evalMove(index, temp))
End If
End Function
