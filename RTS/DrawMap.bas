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
    Randomize Timer
    frmDbg.lstMap.Clear
    
    mapHeight = (frmMain.ScaleHeight / (frmMain.picScene(0).Height / 2))
    mapWidth = 2 * (frmMain.ScaleWidth / frmMain.picScene(0).Width)
    For Y = 0 To frmMain.ScaleHeight Step (frmMain.picScene(0).Height / 2)
        For X = 0 To frmMain.ScaleWidth Step frmMain.picScene(0).Width
            rand = Int(Rnd() * 2)
            
            If (X / 2) Mod 20 = 10 Or (X / 2) Mod 20 = 0 Then
                frmMain.PaintPicture frmMain.picMask.Picture, X + 50, Y - 25, 100, 100, 0, 0, 100, 100, vbSrcAnd
                frmMain.PaintPicture frmMain.picScene(rand).Picture, X + 50, Y - 25, 100, 100, 0, 0, 100, 100, vbSrcPaint
                frmMain.picBackground.PaintPicture frmMain.picMask.Picture, X + 50, Y - 25, 100, 100, 0, 0, 100, 100, vbSrcAnd
                frmMain.picBackground.PaintPicture frmMain.picScene(rand).Picture, X + 50, Y - 25, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
            
            frmMain.PaintPicture frmMain.picMask.Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
            frmMain.PaintPicture frmMain.picScene(rand).Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            frmMain.picBackground.PaintPicture frmMain.picMask.Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
            frmMain.picBackground.PaintPicture frmMain.picScene(rand).Picture, X, Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ReDim Preserve Tile(mapWidth, mapHeight) As terrain
            
            If rand = 0 Then
                Tile(Int(X / 100), Int(Y / 50)).pic = "0" '0 = grass
            ElseIf rand = 1 Then
                Tile(Int(X / 100), Int(Y / 50)).pic = "0A" '0 = grass
            End If
            Tile(Int(X / 100), Int(Y / 50)).X = X
            Tile(Int(X / 100), Int(Y / 50)).Y = Y
            'Top half of odd tile
            If Y Mod 50 < 25 Then
                'Left half of odd tile
                If X Mod 100 < 50 Then
                    If 25 - Int((X Mod 50) / 2) >= Y Mod 50 Then
                    Tile(Int(X / 100), Int(Y / 50)).Y = Y '- 25
                    Tile(Int(X / 100), Int(Y / 50)).X = X '- 50
                    End If
                'Right half of odd tile
                ElseIf X Mod 100 >= 50 Then
                    If Int((X Mod 50) / 2) >= Y Mod 50 Then
                     Tile(Int(X / 100), Int(Y / 50)).Y = Y '- 25
                     Tile(Int(X / 100), Int(Y / 50)).X = X '+ 50
        
                    End If
                End If
            'Bottom half of even tile
            ElseIf Y Mod 50 >= 25 Then
                'Left half of odd tile
                If X Mod 100 < 50 Then
                    If 25 + Int((X Mod 50) / 2) < Y Mod 50 Then
                        Tile(Int(X / 100), Int(Y / 50)).Y = Y + 25
                        Tile(Int(X / 100), Int(Y / 50)).X = X - 50
                    End If
                'Right half of odd tile
                ElseIf X Mod 100 >= 50 Then
                    If 50 - Int((X Mod 50) / 2) < Y Mod 50 Then
                        Tile(Int(X / 100), Int(Y / 50)).Y = Y + 25
                        Tile(Int(X / 100), Int(Y / 50)).X = X + 50
                    End If
                End If
            End If
  
            Tile(Int(X / 100), Int(Y / 50)).selectable = True
            frmDbg.lstMap.AddItem ("(" & Tile(Int(X / 100), Int(Y / 50)).X & "," & Tile(Int(X / 100), Int(Y / 50)).Y & ")" & vbTab)
        Next X
    Next Y
End Function
