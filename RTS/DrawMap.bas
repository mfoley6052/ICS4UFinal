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
    mapHeight = 2 * (frmMain.ScaleHeight / (frmMain.picScene(0).Height / 2))
    mapWidth = 2 * (frmMain.ScaleWidth / frmMain.picScene(0).Width)
    For Y = 0 To frmMain.ScaleHeight Step (frmMain.picScene(0).Height / 4)
        If (Y + 50) Mod 50 = 0 Then
            xStart = 0
        Else
            xStart = 50
        End If
        For X = xStart To frmMain.ScaleWidth + xStart Step frmMain.picScene(0).Width
            rand = Int(Rnd() * 2)
            frmMain.PaintPicture frmMain.picMask.Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
            frmMain.PaintPicture frmMain.picScene(rand).Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            frmMain.picBackground.PaintPicture frmMain.picMask.Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
            frmMain.picBackground.PaintPicture frmMain.picScene(rand).Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            
            ReDim Preserve Tile(mapWidth, mapHeight) As terrain
            
            If rand = 0 Then
                Tile(Int((X - xStart) / 100), Int(Y / 25)).pic = "0" '0 = grass
            ElseIf rand = 1 Then
                Tile(Int((X - xStart) / 100), Int(Y / 25)).pic = "0A" '0 = grass
            End If
            Tile(Int((X - xStart) / 100), Int(Y / 25)).X = X
            Tile(Int((X - xStart) / 100), Int(Y / 25)).Y = Y
            Tile(Int((X - xStart) / 100), Int(Y / 25)).selectable = True
            frmDbg.lstMap.AddItem ("(" & Tile(Int((X - xStart) / 100), Int(Y / 25)).X & "," & Tile(Int((X - xStart) / 100), Int(Y / 25)).Y & ")" & vbTab)
        Next X
    Next Y
End Function
