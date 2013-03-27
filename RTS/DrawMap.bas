Attribute VB_Name = "Module1"
'Draw Map Procedure
'Needs picture boxes with visible false, autodraw true, scalemode pixel:
'picScene holds the picture of the tile
'picMask holds the mask
'form named as frmMain
Public mapHeight As Integer
Public mapWidth As Integer
Public Tile() As terrain

Public Function DrawMap(ByVal mapNum As Integer) As Boolean
'If mapNum = 1 Then
 '   frmMain.picScene.Picture = LoadPicture(App.Path & "\Images\Terrain\1\1Tile.gif")
'End If
    Erase Tile
    Dim rand As Integer
    Dim xStart As Integer
    Randomize Timer
    frmDbg.lstMap.Clear
    mapHeight = 5
    mapWidth = 5
    For Y = 0 To (mapHeight - 1) * (frmMain.picScene(0).Height * 0.75) Step 0.75 * frmMain.picScene(0).Height
        For X = (50 - xStart) To (mapWidth * frmMain.picScene(0).Width) + xStart Step frmMain.picScene(0).Width
            'rand = Int(Rnd() * 2)
            If (Y + 75) Mod 150 = 0 Then
                xStart = 0
            Else
                xStart = frmMain.picScene(0).Width / 2
            End If
            frmMain.PaintPicture frmMain.picMask.Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
            frmMain.PaintPicture frmMain.picScene(0).Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            frmMain.picBackground.PaintPicture frmMain.picScene(0).Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
            
            ReDim Preserve Tile(mapWidth, mapHeight) As terrain
            
            'If rand = 0 Then
                'Tile(Int((X - xStart) / 100), Int(Y / 25)).pic = "0" '0 = grass
            'ElseIf rand = 1 Then
                'Tile(Int((X - xStart) / 100), Int(Y / 25)).pic = "0A" '0 = grass
            'End If
            Tile(Int((X + (50 - xStart)) / frmMain.picScene(0).Width), Int(Y / (frmMain.picScene(0).Height * 0.75))).X = X
            Tile(Int((X + (50 - xStart)) / frmMain.picScene(0).Width), Int(Y / (frmMain.picScene(0).Height * 0.75))).Y = Y
            'frmDbg.lstMap.AddItem ("(" & Tile(Int((X - xStart) / 100), Int(Y / 25)).X & "," & Tile(Int((X - xStart) / 100), Int(Y / 25)).Y & ")" & vbTab)
        Next X
    Next Y
frmMain.tmrTileAnim.Enabled = True
End Function
