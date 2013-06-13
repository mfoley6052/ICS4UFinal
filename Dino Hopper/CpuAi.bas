Attribute VB_Name = "modCpuAi"

Public Function cpuAI(ByVal Index As Integer)

Static counter As Long
Dim temp As String
Dim target As Integer
counter = counter + 1
temp = randDir
If numPlayers <> 1 Then
    For x = 0 To 3
        If isPlayer(x) Then
            If Abs(curX(Index) - curX(x)) + Abs(curY(Index) - curY(x)) < Abs(curX(Index) - curX(target)) + Abs(curY(Index) - curY(target)) Then
                target = x
            End If
        End If
    Next x
End If
If Not blnBounceJump(Index) Then
    If curX(Index) > curX(target) And curY(Index) > curY(target) Then
        If evalMove(Index, "L") Then
            If Not isScared Then
                Call getJump(Index, "L", evalMove(Index, "L"))
            Else
                Call getJump(Index, "R", evalMove(Index, "R"))
            End If
        Else
            Call getJump(Index, temp, evalMove(Index, temp))
        End If
    ElseIf curX(Index) > curX(target) And curY(Index) < curY(target) Then
        If evalMove(Index, "D") Then
            If Not isScared Then
                Call getJump(Index, "D", evalMove(Index, "D"))
            Else
                Call getJump(Index, "U", evalMove(Index, "U"))
            End If
        Else
            Call getJump(Index, temp, evalMove(Index, temp))
        End If
    ElseIf curX(Index) < curX(target) And curY(Index) > curY(target) Then
        If evalMove(Index, "U") Then
            If Not isScared Then
                Call getJump(Index, "U", evalMove(Index, "U"))
            Else
                Call getJump(Index, "D", evalMove(Index, "D"))
            End If
        Else
            Call getJump(Index, temp, evalMove(Index, temp))
        End If
    ElseIf curX(Index) < curX(target) And curY(Index) < curY(target) Then
        If evalMove(Index, "R") Then
            If Not isScared Then
                Call getJump(Index, "R", evalMove(Index, "R"))
            Else
                Call getJump(Index, "L", evalMove(Index, "L"))
            End If
        Else
            Call getJump(Index, temp, evalMove(Index, temp))
        End If
    ElseIf curY(Index) > curY(target) And curX(Index) = curX(target) Then
        If evalMove(Index, "L") Then
            If Not isScared Then
                Call getJump(Index, "L", evalMove(Index, "L"))
            Else
                Call getJump(Index, "R", evalMove(Index, "R"))
            End If
        Else
            Call getJump(Index, temp, evalMove(Index, temp))
        End If
    ElseIf curY(Index) < curY(target) And curX(Index) = curX(target) Then
        If evalMove(Index, "R") Then
            If Not isScared Then
                Call getJump(Index, "R", evalMove(Index, "R"))
            Else
                Call getJump(Index, "L", evalMove(Index, "L"))
            End If
        Else
            Call getJump(Index, temp, evalMove(Index, temp))
        End If
    ElseIf curX(Index) > curX(target) And curY(Index) = curY(target) Then
        If evalMove(Index, "D") Then
            If Not isScared Then
                Call getJump(Index, "D", evalMove(Index, "D"))
            Else
                Call getJump(Index, "U", evalMove(Index, "U"))
            End If
        Else
            Call getJump(Index, temp, evalMove(Index, temp))
        End If
    ElseIf curX(Index) < curX(target) And curY(Index) = curY(target) Then
        If evalMove(Index, "U") Then
            If Not isScared Then
                Call getJump(Index, "U", evalMove(Index, "U"))
            Else
                Call getJump(Index, "D", evalMove(Index, "D"))
            End If
        Else
            Call getJump(Index, temp, evalMove(Index, temp))
        End If
    Else
        Call getJump(Index, temp, evalMove(Index, temp))
    End If
Else
    Call getJump(Index, strDir(Index), evalMove(Index, strDir(Index)))
End If
If gameMode = 0 Then
    If counter Mod 5 = 0 And frmMain.tmrChar(Index).Interval > 200 Then
        frmMain.tmrChar(Index).Interval = Int(frmMain.tmrChar(Index).Interval - 50)
    End If
Else
    If counter Mod 30 = 0 Then
        intMoves(Index) = 2
    End If
End If
End Function
