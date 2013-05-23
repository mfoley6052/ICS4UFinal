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
Else
    Dim dirRand As String
    dirRand = randDir
    Call getJump(index, dirRand, evalMove(index, dirRand))
End If
End Function
