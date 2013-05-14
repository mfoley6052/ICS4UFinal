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
    yStart = 50
    Randomize Timer
    frmDbg.lstMap.Clear
    mapHeight = 7
    mapWidth = 7
    For y = yStart To ((mapHeight - 1) * 75) + yStart Step 75
        For x = (50 - xStart) To (mapWidth * 100) + xStart Step 100
            'rand = Int(Rnd() * 2)
            If ((y - yStart) + 75) Mod 150 = 0 Then
                xStart = 0
            Else
                xStart = 100 / 2
            End If
            ReDim Preserve tile(mapWidth, mapHeight) As terrain
            Dim currentTile As terrain
            currentTile = tile(Int((x + (50 - xStart)) / 100), Int((y - yStart) / 75))
            currentTile.Xc = x \ 100
            currentTile.Yc = (y - yStart) \ 75
            currentTile.x = x
            currentTile.y = y
            currentTile.hasObj = False
            rand = randInt(1, 100)
            If rand >= 25 Then
                currentTile.terType = 0
                If rand > 66 Then
                    rand = 1
                Else
                    rand = 0
                End If
            ElseIf rand < 25 Then
                currentTile.terType = 1
                If rand > 10 Then
                    rand = 2
                Else
                    rand = 3
                End If
            End If
            With frmMain
            If currentTile.Yc = 0 Then
                Set currentTile.picMask = .picMask
                Set currentTile.picTile = .picTile(rand)
            ElseIf oddRow(currentTile.Yc) Then
                If currentTile.Xc = 0 Then
                    Set currentTile.picMask = .picMaskL
                    Set currentTile.picTile = .picTileL(rand)
                ElseIf currentTile.Xc = mapWidth Then
                    Set currentTile.picMask = .picMaskR
                    Set currentTile.picTile = .picTileR(rand)
                Else
                    Set currentTile.picMask = .picMaskLR
                    Set currentTile.picTile = .picTileLR(rand)
                End If
            Else
                Set currentTile.picMask = .picMaskLR
                Set currentTile.picTile = .picTileLR(rand)
            End If
            End With
            tile(Int((x + (50 - xStart)) / 100), Int((y - yStart) / 75)) = currentTile
        Next x
    Next y
frmMain.tmrTileAnimDelay.Enabled = True
tileCount = ((mapWidth * mapHeight) + (Int(mapHeight / 2)))
End Function
