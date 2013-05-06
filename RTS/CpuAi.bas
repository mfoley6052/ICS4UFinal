Attribute VB_Name = "modCpuAi"

Public Function cpuAI(ByVal index As Integer)
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
If index = 0 Then
'Lead the way
    For q = LBound(pathStep) To UBound(pathStep)
        If pathStep(q).x = curX(index) And pathStep(q).y = curY(index) Then
            hasBC = True
        End If
    Next q
    If hasBC Then
        If stepCount < UBound(pathStep) Then
            nextX(index) = pathStep(stepCount).x
            nextY(index) = pathStep(stepCount).y
        End If
    Else
        Randomize Timer
        tempRand = randInt(1, 4)
    End If
Else
'Follow Leader with breadcrumbs
End If
End Function

Public Function nearestCoin(ByVal index As Integer) As Long
smallestX = 0
smallestY = 0
For x = 0 To mapWidth
    For y = 0 To mapHeight
    'if the tile has a coin then check if it is closer than the closest one, if it is make it the closest one
        If tile(x, y).coinEnabled = True Then
            If Abs(curX(index) - tile(x, y).x) < smallestX And Abs(curY(index) - tile(x, y).y) < smallestY Then
                smallestX = Abs(curX(index) - tile(x, y).x)
                smallestY = Abs(curY(index) - tile(x, y).y)
            End If
        End If
    Next y
Next x
End Function
