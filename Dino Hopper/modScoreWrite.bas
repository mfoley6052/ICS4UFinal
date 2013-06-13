Attribute VB_Name = "modScoreWrite"
Public gMode As String
Public playerChoice As String
Public playMode As String
Public score(9) As Record
Public Type Record
    score As Long
    nam As String * 3
End Type
Public Sub WriteScore(ByVal playerName As String, ByVal Pscore As Long)
Dim dump As String
Dim count As Integer
Dim sorted As Boolean
Dim upper As Long
Dim temp As Record
playerChoice = numPlayers
If playMode <> "SOLO" Then
    Open App.Path & "\Scores\" & playMode & "\" & gMode & "\" & playerChoice & ".sav" For Input As #1
Else
    Open App.Path & "\Scores\" & playMode & "\" & playerChoice & ".sav" For Input As #1
End If
    Do Until EOF(1)
        Line Input #1, dump
        count = count + 1
    Loop
Close #1
    If count <= 18 Then
        If playMode <> "SOLO" Then
            Open App.Path & "\Scores\" & playMode & "\" & gMode & "\" & playerChoice & ".sav" For Append As #1
        Else
            Open App.Path & "\Scores\" & playMode & "\" & playerChoice & ".sav" For Append As #1
        End If
            Print #1, playerName
            Print #1, Pscore
        Close #1
    Else
        'Find the lowest score, compare it to the score to be added, and then choose which one to keep
        If playMode <> "SOLO" Then
            Open App.Path & "\Scores\" & playMode & "\" & gMode & "\" & playerChoice & ".sav" For Input As #1
        Else
            Open App.Path & "\Scores\" & playMode & "\" & playerChoice & ".sav" For Input As #1
        End If
        For x = 0 To 9
            If Not EOF(1) Then
                Line Input #1, dump
                score(x).nam = dump
                Line Input #1, dump
                score(x).score = dump
            End If
        Next x
        Close #1
        upper = UBound(score) - 1
        Do Until sorted
            sorted = True
            For y = LBound(score) To upper 'Step - 1
                If score(y + 1).score < score(y).score Then
                    temp.score = score(y).score
                    temp.nam = score(y).nam
                    score(y).score = score(y + 1).score
                    score(y).nam = score(y + 1).nam
                    score(y + 1).score = temp.score
                    score(y + 1).nam = temp.nam
                    sorted = False
                End If
            Next y
            upper = upper - 1
        Loop
        If score(0).score < Pscore Then
            Dim temporary() As Record
            'BUG:Deletes any scores that share the name of the lowest score
            'Delete the low score AND the name of the player that got that score, then add the new score
            If playMode <> "SOLO" Then
                Open App.Path & "\Scores\" & playMode & "\" & gMode & "\" & playerChoice & ".sav" For Input As #1
            Else
                Open App.Path & "\Scores\" & playMode & "\" & playerChoice & ".sav" For Input As #1
            End If
                count = 0
                Do Until EOF(1)
                    Line Input #1, dump
                    If Not IsNumeric(dump) Then
                        If dump <> score(0).nam Then
                            ReDim Preserve temporary(count) As Record
                            temporary(count).nam = dump
                        End If
                    Else
                        If dump <> score(0).score Then
                            ReDim Preserve temporary(count) As Record
                            temporary(count).score = dump
                            count = count + 1
                        End If
                    End If
                Loop
            Close #1
            If playMode <> "SOLO" Then
                Open App.Path & "\Scores\" & playMode & "\" & gMode & "\" & playerChoice & ".sav" For Output As #1
            Else
                Open App.Path & "\Scores\" & playMode & "\" & playerChoice & ".sav" For Output As #1
            End If
                For x = 0 To UBound(temporary)
                    Print #1, temporary(x).nam
                    Print #1, temporary(x).score
                Next x
                Print #1, playerName
                Print #1, Pscore
            Close #1
        Else
            MsgBox ("Score didnt get into top 10")
        End If
    End If
End Sub

