Attribute VB_Name = "modCollision"
Public Sub getHurt(ByVal Index As Integer, ByVal enemyIndex As Integer)
frmMain.mmcHit.Command = "prev"
frmMain.mmcHit.Command = "open"
frmMain.mmcHit.Command = "play"
If blnRecover(Index) = False Then
    If isPlayer(Index) Then
        'check if all lives lost
        If intLives(Index) > 1 Then
            intLives(Index) = intLives(Index) - 1
        Else
            frmMain.tmrChar(Index).Enabled = False
            frmGUI.lblLives(Index).Visible = False
            frmGUI.lblMulti(Index).Visible = False
            frmGUI.lblScore(Index).Visible = False
            'repaint the tile that the character was on
            Call clearTile(tile(prevX(Index), prevY(Index)), True, Index, "SelPXPY")
            'Clear egg from gui
            With frmGUI
            .PaintPicture frmMain.picEggMask(0).Image, .lblLives(Index).Left - 24, .lblLives(Index).Top - (frmMain.picEggG(0).Height - .lblLives(Index).Height), 22, 30, 0, 0, 22, 30, vbSrcAnd
            .PaintPicture .picCyan.Image, .lblLives(Index).Left - 24, .lblLives(Index).Top - (frmMain.picEggG(0).Height - .lblLives(Index).Height), 22, 30, .picCyan.Width - (.picCyan.Width - frmMain.picEggG(0).Width), .picCyan.Height - (.picCyan.Height - frmMain.picEggG(0).Height), 22, 30, vbSrcPaint
            End With
            Dim blnGameOver
            blnGameOver = True
            'check for game over
            For checkPlayers = 0 To numPlayers - 1
                If frmMain.tmrChar(checkPlayers).Enabled Then
                    blnGameOver = False
                    With frmGUI
                    .PaintPicture frmMain.picEggMask(0).Image, .lblLives(checkPlayers).Left - 24, .lblLives(checkPlayers).Top - (frmMain.picEggG(0).Height - .lblLives(checkPlayers).Height), 22, 30, 0, 0, 22, 30, vbSrcAnd
                    .PaintPicture frmMain.picEggG(0).Image, .lblLives(checkPlayers).Left - 24, .lblLives(checkPlayers).Top - (frmMain.picEggG(0).Height - .lblLives(checkPlayers).Height), 22, 30, 0, 0, 22, 30, vbSrcPaint
                    End With
                End If
            Next checkPlayers
            'game over case
            If blnGameOver Then
                frmMain.tmrGameOver = blnGameOver
                Call getGameEnd
                Exit Sub
            End If
        End If
        Call refreshLabels(Index, False, True, False)
    End If
    'stun char in arcade mode
    If gameMode = 0 Then
        If isPlayer(Index) Then
            blnPlayerMoveable(Index) = False
        Else
            frmMain.tmrCPUMove(Index).Enabled = False
        End If
        frmMain.tmrStun(Index).Enabled = True
    End If
    blnRecover(Index) = True
    frmMain.tmrChar(Index).Tag = "Recover"
    Call getPowEffect(Index, "Recover")
    blnPlayerMoveable(Index) = True
End If
End Sub
