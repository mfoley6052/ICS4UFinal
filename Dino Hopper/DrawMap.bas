Attribute VB_Name = "modDrawMap"
'Draw Map Procedure
'Needs picture boxes with visible false, autodraw true, scalemode pixel:
'picScene holds the picture of the tile
'picMask holds the mask
'form named as frmMain
Public mapHeight As Integer
Public mapWidth As Integer
Public tileCount As Integer
Public tile(7, 7) As terrain

'draw the game map
Public Sub DrawMap(ByVal intWidth As Integer, ByVal intHeight As Integer)
Dim rand As Integer
Dim xStart As Integer
Dim yStart As Integer
yStart = 87
Randomize Timer
mapHeight = intWidth
mapWidth = intHeight
'for each tile column; get position with mapHeight and starting y value
For y = yStart To ((mapHeight - 1) * 75) + yStart Step 75
    'for each tile in a row; get position with mapWidth and starting x value
    For x = (50 - xStart) To (mapWidth * 100) + xStart Step 100
        'for every odd row, tiles are translated 50 pixels to the right
        If ((y - yStart) + 75) Mod 150 = 0 Then
            xStart = 0
        Else
            xStart = 100 / 2
        End If
        Dim currentTile As terrain
        currentTile = tile(Int((x + (50 - xStart)) / 100), Int((y - yStart) / 75))
        currentTile.Xc = x \ 100
        currentTile.Yc = (y - yStart) \ 75
        currentTile.x = x
        currentTile.y = y
        currentTile.hasObj = False
        'randomly generate the type of tile (dirt gives no points for characters, grass gives some)
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
        'get tile pictures for current tile
        '(when tile is touching another tile, different picture is used where top pixels of corner are cut off so tiles don't overlap)
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
        tile(Int((x + (50 - xStart)) / 100), Int((y - yStart) / 75)) = currentTile
    Next x
Next y
'start tile animation when tiles are defined
frmMain.tmrTileAnimDelay.Enabled = True
'set number of tiles
tileCount = ((mapWidth * mapHeight) + (Int(mapHeight / 2)))
End Sub
