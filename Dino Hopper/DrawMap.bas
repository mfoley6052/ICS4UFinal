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
    Randomize Timer
    frmDbg.lstMap.Clear
    mapHeight = 7
    mapWidth = 7
    For y = 0 To (mapHeight - 1) * (frmMain.picScene(0).Height * 0.75) Step 0.75 * frmMain.picScene(0).Height
        For x = (50 - xStart) To (mapWidth * frmMain.picScene(0).Width) + xStart Step frmMain.picScene(0).Width
            'rand = Int(Rnd() * 2)
            If (y + 75) Mod 150 = 0 Then
                xStart = 0
            Else
                xStart = frmMain.picScene(0).Width / 2
            End If
            ReDim Preserve tile(mapWidth, mapHeight) As terrain

            tile(Int((x + (50 - xStart)) / frmMain.picScene(0).Width), Int(y / (frmMain.picScene(0).Height * 0.75))).Xc = x \ 100
            tile(Int((x + (50 - xStart)) / frmMain.picScene(0).Width), Int(y / (frmMain.picScene(0).Height * 0.75))).Yc = y \ 75
            tile(Int((x + (50 - xStart)) / frmMain.picScene(0).Width), Int(y / (frmMain.picScene(0).Height * 0.75))).x = x
            tile(Int((x + (50 - xStart)) / frmMain.picScene(0).Width), Int(y / (frmMain.picScene(0).Height * 0.75))).y = y
            tile(Int((x + (50 - xStart)) / frmMain.picScene(0).Width), Int(y / (frmMain.picScene(0).Height * 0.75))).hasObj = False

            'frmDbg.lstMap.AddItem ("(" & Tile(Int((X - xStart) / 100), Int(Y / 25)).X & "," & Tile(Int((X - xStart) / 100), Int(Y / 25)).Y & ")" & vbTab)
        Next x
    Next y
frmMain.tmrTileAnimDelay.Enabled = True
tileCount = ((mapWidth * mapHeight) + (Int(mapHeight / 2)))
End Function
