Attribute VB_Name = "modCpuAi"

Public Function cpuAI(ByVal index As Integer)
If curX(index) > curX(0) And curY(index) > curY(0) Then
    Call getJump(index, "L")
ElseIf curX(index) > curX(0) And curY(index) < curY(0) Then
    Call getJump(index, "D")
ElseIf curX(index) < curX(0) And curY(index) > curY(0) Then
    Call getJump(index, "U")
ElseIf curX(index) < curX(0) And curY(index) < curY(0) Then
    Call getJump(index, "R")
Else
    Call getJump(index, RandDir)
End If
End Function
