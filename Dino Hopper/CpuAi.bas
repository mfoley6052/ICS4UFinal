Attribute VB_Name = "modCpuAi"

Public Function cpuAI(ByVal Index As Integer)
Dim tempRand As Integer
Dim hasBC As Boolean
Static stepCount As Integer

If intScore < 1000 Then
    ReDim Preserve pathStep(1 To 5)
ElseIf intScore < 2000 Then
    ReDim Preserve pathStep(1 To 10)
ElseIf intScore < 3000 Then
    ReDim Preserve pathStep(1 To 15)
ElseIf intScore < 4000 Then
    ReDim Preserve pathStep(1 To 20)
ElseIf intScore < 5000 Then
    ReDim Preserve pathStep(1 To 30)
End If
If Index = 1 Then
'Lead the way
    For q = LBound(pathStep) To UBound(pathStep)
        If pathStep(q).x = curX(Index) And pathStep(q).y = curY(Index) Then
            hasBC = True
        End If
    Next q
    If hasBC Then
        'If stepCount < UBound(pathStep) Then
         '   nextX(Index) = pathStep(stepCount).x
          '  nextY(Index) = pathStep(stepCount).y
       ' End If
    Else
        Randomize Timer
        tempRand = randInt(1, 4)
    End If
Else
'Follow Leader with breadcrumbs
End If
End Function
