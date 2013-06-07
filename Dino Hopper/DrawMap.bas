Attribute VB_Name = "modDrawMap"
'Draw Map Procedure
'Needs picture boxes with visible false, autodraw true, scalemode pixel:
'picScene holds the picture of the tile
'picMask holds the mask
'form named as frmMain
Public mapHeight As Integer
Public mapWidth As Integer
Public tileCount As Integer
Public tile() As terrain

Public Function DrawMap(ByVal mapNum As Integer) As Boolean
'If mapNum = 1 Then
 '   frmMain.picScene.Picture = LoadPicture(App.Path & "\Images\Terrain\1\1Tile.gif")
'End If
    Erase tile
    Dim rand As Integer
    Dim xStart As Integer
    Dim yStart As Integer
    yStart = 87
    Randomize Timer
    frmDbg.lstMap.Clear
    mapHeight = 7
    mapWidth = 7
    For Y = yStart To ((mapHeight - 1) * 75) + yStart Step 75
        For X = (50 - xStart) To (mapWidth * 100) + xStart Step 100
            'rand = Int(Rnd() * 2)
            If ((Y - yStart) + 75) Mod 150 = 0 Then
                xStart = 0
            Else
                xStart = 100 / 2
            End If
            ReDim Preserve tile(mapWidth, mapHeight) As terrain
            Dim currentTile As terrain
            currentTile = tile(Int((X + (50 - xStart)) / 100), Int((Y - yStart) / 75))
            currentTile.Xc = X \ 100
            currentTile.Yc = (Y - yStart) \ 75
            currentTile.X = X
            currentTile.Y = Y
            currentTile.hasObj = False
            rand = randInt(1, 100)
            If rand >= 25 Then
                currentTile.terType = "G"
                If rand > 66 Then
                    rand = 1
                Else
                    rand = 0
                End If
            ElseIf rand < 25 Then
                currentTile.terType = "D"
                If rand > 10 Then
                    rand = 2
                Else
                    rand = 3
                End If
            End If
            With frmMain
            If currentTile.Yc = 0 Then
                currentTile.cutoff = ""
                Set currentTile.picMask = .picMask
                Set currentTile.picTile = .picTile(rand)
            ElseIf oddRow(currentTile.Yc) Then
                If currentTile.Xc = 0 Then
                    currentTile.cutoff = "L"
                    Set currentTile.picMask = .picMaskL
                    Set currentTile.picTile = .picTileL(rand)
                ElseIf currentTile.Xc = mapWidth Then
                    currentTile.cutoff = "R"
                    Set currentTile.picMask = .picMaskR
                    Set currentTile.picTile = .picTileR(rand)
                Else
                    currentTile.cutoff = "LR"
                    Set currentTile.picMask = .picMaskLR
                    Set currentTile.picTile = .picTileLR(rand)
                End If
            Else
                currentTile.cutoff = "LR"
                Set currentTile.picMask = .picMaskLR
                Set currentTile.picTile = .picTileLR(rand)
            End If
            End With
            tile(Int((X + (50 - xStart)) / 100), Int((Y - yStart) / 75)) = currentTile
        Next X
    Next Y
frmMain.tmrTileAnimDelay.Enabled = True
tileCount = ((mapWidth * mapHeight) + (Int(mapHeight / 2)))
End Function
