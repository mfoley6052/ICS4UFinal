Attribute VB_Name = "modPaint"
Public Sub clearTile(tileInput As terrain, ByVal bypassForObj As Boolean, Optional Index As Integer, Optional callID As String)
With frmMain
Dim tilePic As Object
Set tilePic = tileInput.picTile 'Set tilePic as tile's tile pic
'if object not enabled on tile or bypassForObj is true (doesn't paint if bypassForObj is true and object is enabled on tile)
If Not tileInput.hasObj Or Not bypassForObj Then
    If Not bypassForObj Then 'if object is caller of function
        If tileInput.hasObj And Mid(callID, 1, 5) = "ObjXY" Then
            'paint over object
            If tileInput.objType(0) <> "Terrain" Then
                Call clearTileTop(tileInput, True, callID) 'paint top of tile with object mask
            ElseIf tileInput.objType(0) = "Terrain" Then
                'paint over full tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf Mid(callID, 1, 6) = "ObjTop" Then 'if call is for top of tile
            Call clearTileTop(tileInput, False, callID) 'paint top of tile or top left or right according to callID
        Else 'no object on tile
            If callID = "ObjOddX-Y" Or callID = "ObjEven+X-Y" Then
                'paint over bottom left half of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 25, 50, 75, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 0, 25, 50, 75, vbSrcCopy
                .PaintPicture .picMaskSides.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
            ElseIf callID = "ObjOdd-X-Y" Or callID = "ObjEvenX-Y" Then
                'paint over bottom right half of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 50, 25, 50, 75, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 50, 25, 50, 75, vbSrcCopy
                .PaintPicture .picMaskSides.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 50, 0, 50, 75, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
            End If
        End If
    Else 'called by character
        If tileInput.hasChar Then 'if character is on tile
            If (tileInput.Xc = curX(Index) And tileInput.Yc = curY(Index)) Then 'if character on tile is character that called clear
                Dim tileAlt As terrain
                If tileInput.Yc = 0 Then '(x, 0)
                    If frameCounter(Index) > 0 And strDir(Index) = "L" Then
                        If callID = "SelXY" Then
                            Call clearVoid(tileInput, True, True) 'paint spacer
                            If tileInput.Xc > 0 Then
                                tileAlt = tile(curX(Index) - 1, curY(Index))
                                Call clearVoid(tileAlt, False, True, callID) 'paint right spacer
                            Else
                                tileAlt = tile(curX(Index), curY(Index) + 1)
                                Call clearVoid(tileAlt, True, False, callID) 'paint left spacer
                                .PaintPicture picBG, tileInput.x - 50, tileInput.y - 99, 50, 49, tileInput.x - 50, tileInput.y - 99, 50, 49, vbSrcCopy
                                .PaintPicture picBG, tileInput.x - 50, tileInput.y - 50, 50, 125, tileInput.x - 50, tileInput.y - 50, 50, 125, vbSrcCopy
                            End If
                        Else
                            If tileInput.Xc > 0 Then
                                tileAlt = tile(curX(Index) - 1, curY(Index))
                            Else
                                tileAlt = tile(curX(Index), curY(Index) + 1)
                            End If
                            Set tilePic = tileAlt.picTile
                            .picBackground.PaintPicture tilePic.Image, tileAlt.x, tileAlt.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                            .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                            .PaintPicture tileAlt.picMask.Image, tileAlt.x, tileAlt.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                            .PaintPicture .picBuffer.Image, tileAlt.x, tileAlt.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                        End If
                    ElseIf frameCounter(Index) > 0 And strDir(Index) = "U" Then
                        If callID = "SelXY" Then
                            Call clearVoid(tileInput, True, True, callID) 'paint spacer
                            If tileInput.Xc < mapWidth - 1 Then
                                tileAlt = tile(curX(Index) + 1, curY(Index))
                                Call clearVoid(tileAlt, True, False, callID) 'paint left spacer
                            Else
                                tileAlt = tile(curX(Index), curY(Index) + 1)
                                Call clearVoid(tileAlt, False, True, callID) 'paint right spacer
                                .PaintPicture picBG, tileInput.x + 100, tileInput.y - 99, 50, 49, tileInput.x + 100, tileInput.y - 99, 50, 49, vbSrcCopy
                                .PaintPicture picBG, tileInput.x + 100, tileInput.y - 50, 50, 125, tileInput.x + 100, tileInput.y - 50, 50, 125, vbSrcCopy
                            End If
                        Else
                            If tileInput.Xc < mapWidth - 1 Then
                                tileAlt = tile(curX(Index) + 1, curY(Index))
                            Else
                                tileAlt = tile(curX(Index), curY(Index) + 1)
                            End If
                            Set tilePic = tileAlt.picTile
                            .picBackground.PaintPicture tilePic.Image, tileAlt.x, tileAlt.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                            .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                            .PaintPicture tileAlt.picMask.Image, tileAlt.x, tileAlt.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                            .PaintPicture .picBuffer.Image, tileAlt.x, tileAlt.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                        End If
                    Else
                        Call clearVoid(tileInput, True, True, callID) 'paint spacer
                    End If
                ElseIf tileInput.Yc = mapHeight - 1 Then '(x, mapHeight - 1)
                    If frameCounter(Index) > 0 And strDir(Index) = "R" Then
                        If callID = "CharBottom+X+Y" Then
                            If tileInput.Xc = mapWidth - 1 Then
                                .PaintPicture picBG, tileInput.x + 100, tileInput.y + 25, 50, 100, tileInput.x + 100, tileInput.y + 25, 50, 100, vbSrcCopy
                            Else
                                tileAlt = tile(curX(Index) + 1, curY(Index))
                                Set tilePic = tileAlt.picTile
                                'paint over bottom left half of tile
                                .picBackground.PaintPicture tilePic.Image, tileAlt.x, tileAlt.y + 50, 50, 50, 0, 50, 50, 50, vbSrcCopy
                                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 74, 0, 26, 50, 74, vbSrcCopy
                                .PaintPicture tileAlt.picMask.Image, tileAlt.x, tileAlt.y + 26, 50, 74, 0, 26, 50, 74, vbSrcAnd
                                .PaintPicture .picBuffer.Image, tileAlt.x, tileAlt.y + 26, 50, 74, 0, 0, 50, 74, vbSrcPaint
                            End If
                        Else
                            .PaintPicture picBG, tileInput.x + 50, tileInput.y + 76, 100, 49, tileInput.x + 50, tileInput.y + 76, 100, 49, vbSrcCopy
                        End If
                    ElseIf frameCounter(Index) > 0 And strDir(Index) = "D" Then
                        If callID = "CharBottom-X+Y" Then
                            If tileInput.Xc = 0 Then
                                .PaintPicture picBG, tileInput.x - 50, tileInput.y + 25, 50, 100, tileInput.x - 50, tileInput.y + 25, 50, 100, vbSrcCopy
                            Else
                                tileAlt = tile(curX(Index) - 1, curY(Index))
                                Set tilePic = tileAlt.picTile
                                'paint over bottom right half of tile
                                .picBackground.PaintPicture tilePic.Image, tileAlt.x + 50, tileAlt.y + 50, 50, 50, 50, 50, 50, 50, vbSrcCopy
                                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 74, 50, 26, 50, 74, vbSrcCopy
                                .PaintPicture tileAlt.picMask.Image, tileAlt.x + 50, tileAlt.y + 26, 50, 74, 50, 26, 50, 74, vbSrcAnd
                                .PaintPicture .picBuffer.Image, tileAlt.x + 50, tileAlt.y + 26, 50, 74, 0, 0, 50, 74, vbSrcPaint
                            End If
                        Else
                            .PaintPicture picBG, tileInput.x - 50, tileInput.y + 76, 100, 49, tileInput.x - 50, tileInput.y + 76, 100, 49, vbSrcCopy
                        End If
                    End If
                ElseIf oddRow(tileInput.Yc) Then '(x, odd)
                    If tileInput.Xc = 0 Then 'if first column
                        If frameCounter(Index) > 0 And strDir(Index) = "L" Then 'if jumping off left side edge
                            If callID <> "CharSide-X-Y" Then
                                Call clearVoid(tileInput, True, False) 'paint left spacer
                                .PaintPicture picBG, tileInput.x, tileInput.y - 74, 50, 24, tileInput.x, tileInput.y - 74, 50, 24, vbSrcCopy
                                If tileInput.Yc = 1 Then 'tilerow is 1
                                    .PaintPicture picBG, tileInput.x, tileInput.y - 125, 50, 51, tileInput.x, tileInput.y - 125, 50, 51, vbSrcCopy
                                End If
                            ElseIf callID = "CharSide-X-Y" Then
                                If tileInput.Yc > 1 Then 'tilerow greater than 1 (tile above in jump)
                                    tileAlt = tile(curX(Index), curY(Index) - 2)
                                    Set tilePic = tileAlt.picTile
                                    'paint over bottom left half of tile
                                    .picBackground.PaintPicture tilePic.Image, tileAlt.x, tileAlt.y + 50, 50, 50, tileAlt.x, tileAlt.y + 50, 50, 50, vbSrcCopy
                                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 0, 50, 50, 50, vbSrcCopy
                                    .PaintPicture tileAlt.picMask.Image, tileAlt.x, tileAlt.y + 50, 50, 50, 0, 50, 50, 50, vbSrcAnd
                                    .PaintPicture .picBuffer.Image, tileAlt.x, tileAlt.y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
                                Else 'tilerow is 1
                                    .PaintPicture picBG, tileInput.x, tileInput.y - 124, 50, 49, tileInput.x, tileInput.y - 124, 50, 49, vbSrcCopy
                                End If
                            End If
                        ElseIf frameCounter(Index) > 0 And strDir(Index) = "D" Then
                            If callID <> "CharSide-X+Y" Then
                                Call clearVoid(tileInput, True, False, callID) 'paint left spacer
                                .PaintPicture picBG, tileInput.x, tileInput.y + 76, 50, 49, tileInput.x, tileInput.y + 76, 50, 49, vbSrcCopy
                            End If
                        Else
                            Call clearVoid(tileInput, True, False, callID) 'paint left spacer
                        End If
                    ElseIf tileInput.Xc = mapWidth Then 'if last column
                        If frameCounter(Index) > 0 And strDir(Index) = "U" Then 'if jumping off right side edge
                            If callID <> "CharSide+X-Y" Then 'if jumping off right edge
                                Call clearVoid(tileInput, False, True, callID) 'paint right spacer
                                .PaintPicture picBG, tileInput.x + 50, tileInput.y - 74, 50, 24, tileInput.x + 50, tileInput.y - 74, 50, 24, vbSrcCopy
                                If tileInput.Yc = 1 Then 'tilerow is 1
                                    .PaintPicture picBG, tileInput.x + 50, tileInput.y - 125, 50, 51, tileInput.x + 50, tileInput.y - 125, 50, 51, vbSrcCopy
                                End If
                            ElseIf callID = "CharSide+X-Y" Then
                                If tileInput.Yc > 1 Then
                                    tileAlt = tile(curX(Index), curY(Index) - 2)
                                    Set tilePic = tileAlt.picTile
                                    'paint over bottom right half of tile
                                    .picBackground.PaintPicture tilePic.Image, tileAlt.x + 50, tileAlt.y + 50, 50, 50, tileAlt.x + 50, tileAlt.y + 50, 50, 50, vbSrcCopy
                                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 50, 50, 50, 50, vbSrcCopy
                                    .PaintPicture tileAlt.picMask.Image, tileAlt.x + 50, tileAlt.y + 50, 50, 50, 50, 50, 50, 50, vbSrcAnd
                                    .PaintPicture .picBuffer.Image, tileAlt.x + 50, tileAlt.y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
                                Else
                                    .PaintPicture picBG, tileInput.x + 50, tileInput.y - 124, 50, 49, tileInput.x + 50, tileInput.y - 124, 50, 49, vbSrcCopy
                                End If
                            End If
                        ElseIf frameCounter(Index) > 0 And strDir(Index) = "R" Then
                            If callID <> "CharSide+X+Y" Then
                                Call clearVoid(tileInput, False, True, callID) 'paint right spacer
                                .PaintPicture picBG, tileInput.x + 50, tileInput.y + 76, 50, 49, tileInput.x + 50, tileInput.y + 76, 50, 49, vbSrcCopy
                            End If
                        Else
                            Call clearVoid(tileInput, False, True) 'paint right spacer
                        End If
                    End If
                End If
                Set tilePic = tileInput.picTile
                If frameCounter(Index) = 0 Or (curX(Index) <> nextX(Index) And curY(Index) <> nextY(Index)) Or callID = "SelXY" Then
                    If Not charTouchingTile(Index, tileInput) Or frameCounter(Index) > 0 Then 'if not touching a char or frame is greater than 0
                        'paint full tile
                        .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                        .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                        .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                        .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                    Else 'if touching a char, only paint over top of tile
                        Call clearTileTop(tileInput, False, callID)
                    End If
                End If
            End If
        Else 'if not character on tile
            'if coordinates match previous tile of character
            If (tileInput.Xc = prevX(Index) And tileInput.Yc = prevY(Index)) Then 'paint top of tile if previous
                Call clearTileTop(tileInput, False, callID)
            End If
            If callID = "CharEven+X-Y" Or callID = "CharOddX-Y" Or callID = "CharNOddNXN-Y" Or callID = "CharNEvenN+XN-Y" Or callID = "CharNOddNXN+Y" Or callID = "CharNEvenN+XN+Y" Then 'paint over bottom left of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 25, 50, 75, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 0, 25, 50, 75, vbSrcCopy
                .PaintPicture .picMaskSides.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
            ElseIf callID = "CharEvenX-Y" Or callID = "CharOdd-X-Y" Or callID = "CharNOddN-XN+Y" Or callID = "CharNEvenNXN+Y" Or callID = "CharNOddN-XN-Y" Or callID = "CharNEvenNXN-Y" Then 'paint over bottom right of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 50, 25, 50, 75, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 50, 25, 50, 75, vbSrcCopy
                .PaintPicture .picMaskSides.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 50, 0, 50, 75, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
            ElseIf callID = "CharTopNXNY" Or callID = "SelPXPY" Then 'paint over top of tile
                Call clearTileTop(tileInput, False, callID)
            ElseIf callID = "CharNXNY" Then
                If strDir(Index) = "L" Then
                    .picBackground.PaintPicture tilePic.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 50, 25, 50, 75, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 50, 25, 50, 75, vbSrcCopy
                    .PaintPicture .picMaskSides.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 50, 0, 50, 75, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
                ElseIf strDir(Index) = "U" Then
                    .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 25, 50, 75, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 0, 25, 50, 75, vbSrcCopy
                    .PaintPicture .picMaskSides.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
                End If
                Call clearTileTop(tileInput, False, callID)
            Else 'paint over full tile
                callID = callID
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    End If
ElseIf bypassForObj And tileInput.hasObj Then
    'paint tile
    .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
    .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
    .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    Call PaintObj(tileInput.objType(0), tileInput.objType(1), tileInput.objFrame, tileInput.Xc, tileInput.Yc, False)
End If
End With
End Sub

Public Sub clearVoid(tileInput As terrain, ByVal blnL As Boolean, ByVal blnR As Boolean, Optional callID As String) 'clear empty spots on map
With frmMain
    If blnL And blnR Then
        .PaintPicture picBG, tileInput.x, tileInput.y - 99, 100, 25, tileInput.x, tileInput.y - 99, 100, 25, vbSrcCopy
        .PaintPicture picBG, tileInput.x, tileInput.y - 74, 100, 98, tileInput.x, tileInput.y - 74, 100, 98, vbSrcCopy
    ElseIf blnL Then
        If tileInput.Yc = 0 Then
            .PaintPicture picBG, tileInput.x, tileInput.y - 99, 50, 49, tileInput.x, tileInput.y - 99, 50, 49, vbSrcCopy
        End If
        .PaintPicture picBG, tileInput.x, tileInput.y - 50, 50, 74, tileInput.x, tileInput.y - 50, 50, 74, vbSrcCopy
    ElseIf blnR Then
        If tileInput.Yc = 0 Then
            .PaintPicture picBG, tileInput.x + 50, tileInput.y - 99, 50, 49, tileInput.x + 50, tileInput.y - 99, 50, 49, vbSrcCopy
        End If
        .PaintPicture picBG, tileInput.x + 50, tileInput.y - 50, 50, 74, tileInput.x + 50, tileInput.y - 50, 50, 74, vbSrcCopy
    End If
End With
End Sub

Public Sub clearTileTop(tileInput As terrain, ByVal blnObjMask As Boolean, Optional callID As String)
Dim tilePic As Object
Set tilePic = tileInput.picTile
With frmMain
If Not blnObjMask Then
    If tileInput.Yc = 0 Then
        .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
        .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
        .PaintPicture .picMaskTop.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
    ElseIf tileInput.Xc = 0 And oddRow(tileInput.Yc) Then
        If callID = "ObjTopLeftXY" Then
            .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 50, 50, 0, 0, 50, 50, vbSrcCopy
            .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 0, 0, 50, 50, vbSrcCopy
        Else
            .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
            .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
            .PaintPicture .picMaskTopLR.Image, tileInput.x + 50, tileInput.y, 50, 50, 50, 0, 50, 50, vbSrcAnd
        End If
        .PaintPicture .picMaskTopLR.Image, tileInput.x, tileInput.y, 50, 50, 0, 0, 50, 50, vbSrcAnd
    ElseIf tileInput.Xc = mapWidth Then
        If callID = "ObjTopRightXY" Then
            .picBackground.PaintPicture tilePic.Image, tileInput.x + 50, tileInput.y, 50, 50, 50, 0, 50, 50, vbSrcCopy
            .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 50, 0, 50, 50, vbSrcCopy
        Else
            .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
            .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
            .PaintPicture .picMaskTopLR.Image, tileInput.x, tileInput.y, 50, 50, 0, 0, 50, 50, vbSrcAnd
        End If
        .PaintPicture .picMaskTopLR.Image, tileInput.x + 50, tileInput.y, 50, 50, 50, 0, 50, 50, vbSrcAnd
    Else
        .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
        .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
        .PaintPicture .picMaskTopLR.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
    End If
    If callID = "ObjTopLeftXY" Or callID = "ObjTopRightXY" Then
        If callID = "ObjTopLeftXY" Then
            .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 50, 50, 0, 0, 50, 50, vbSrcPaint
        ElseIf callID = "ObjTopRightXY" Then
            .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y, 50, 50, 50, 0, 50, 50, vbSrcPaint
        End If
    Else
        .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcPaint
    End If
Else
    'causes glitch; needs fixing
    .picBackground.PaintPicture tilePic.Image, tileInput.x + tileInput.objXOffset, tileInput.y, 100 - (tileInput.objXOffset * 2), 50, tileInput.objXOffset, 0, 100 - (tileInput.objXOffset * 2), 50, vbSrcCopy
    .picBuffer.PaintPicture tilePic.Image, 0, 0, 100 - (tileInput.objXOffset * 2), 50, tileInput.objXOffset, 0, 100 - (tileInput.objXOffset * 2), 50, vbSrcCopy
    .PaintPicture tileInput.objMask.Image, tileInput.x + tileInput.objXOffset, tileInput.y, 100 - (tileInput.objXOffset * 2), 50, 0, 0, 100 - (tileInput.objXOffset * 2), 50, vbSrcAnd
    .PaintPicture .picBuffer.Image, tileInput.x + tileInput.objXOffset, tileInput.y, 100 - (tileInput.objXOffset * 2), 50, 0, 0, 100 - (tileInput.objXOffset * 2), 50, vbSrcPaint
    If tileInput.objSparkle And tileInput.objFrame >= 13 Then
        .PaintPicture .picSparkleMask(tileInput.objFrame - 13).Image, tileInput.x + tileInput.objXOffset, tileInput.y + (tileInput.objYOffset + 3), 18, 35, 0, 0, 18, 35, vbSrcAnd
        .PaintPicture .picBuffer.Image, tileInput.x + tileInput.objXOffset, tileInput.y + (tileInput.objYOffset + 3), 18, 35, tileInput.x + tileInput.objXOffset, tileInput.y + (tileInput.objYOffset + 3), 18, 35, vbSrcPaint
    End If
End If
End With
End Sub

Public Function PaintObj(ByVal strObjType As String, ByVal strType As String, ByVal intFrame As Integer, ByVal intObjX As Integer, ByVal intObjY As Integer, ByVal killObj As Boolean)
Dim intXOffset As Integer
Dim intYOffset As Integer
Dim intFrameOffset As Integer
If Not killObj Then 'if object has not expired
    If strObjType = "Coin" Then
        tile(intObjX, intObjY).objXOffset = 41
        tile(intObjX, intObjY).objYOffset = 0
        If intFrame > 13 Then
            intFrameOffset = -14
        Else
            intFrameOffset = 0
        End If
    ElseIf strObjType = "Pow" Then
        If strType = "Scare" Then
            tile(intObjX, intObjY).objXOffset = 35
            tile(intObjX, intObjY).objYOffset = 0
            If intFrame > 4 Then
                intFrameOffset = (intFrame - (9 - intFrame)) * -1 'frames 5 - 9 will play in reverse
            Else
                intFrameOffset = 0
            End If
        ElseIf strType = "Speed" Then
            tile(intObjX, intObjY).objXOffset = 41
            tile(intObjX, intObjY).objYOffset = 0
            intFrameOffset = 0
        ElseIf strType = "Freeze" Then
            tile(intObjX, intObjY).objXOffset = 32
            tile(intObjX, intObjY).objYOffset = 0
            If intFrame < 14 Then
                intFrameOffset = -intFrame
            Else
                intFrameOffset = -14
            End If
        End If
    ElseIf strObjType = "Egg" Then
        tile(intObjX, intObjY).objXOffset = 39
        tile(intObjX, intObjY).objYOffset = 0
        'multipliers are for gold eggs, where 3 animation cycles represents 1 frame cycle
        If intFrame - (8 * Int(intFrame / 8)) = 6 Or intFrame - (8 * Int(intFrame / 8)) = 7 Then
            'frames 6 and 7 will both use 2nd sprite (egg sprites play at half speed)
            intFrameOffset = -4
        Else
            intFrameOffset = 0
        End If
    End If
End If
intXOffset = tile(intObjX, intObjY).objXOffset
intYOffset = tile(intObjX, intObjY).objYOffset
If (gameMode = 0 And tile(intObjX, intObjY).objTimer > 0) Or (gameMode <> 0 And (intFrame > 0)) Or killObj Then
    Call clearVoid(tile(intObjX, intObjY), checkClearVoid(tile(intObjX, intObjY), True, False), checkClearVoid(tile(intObjX, intObjY), False, True), "ObjXY")
    If checkClearVoid(tile(intObjX, intObjY), True, False) Then
        Call clearTile(tile(intObjX, intObjY), False, -1, "ObjTopLeftXY")
    End If
    If checkClearVoid(tile(intObjX, intObjY), False, True) Then
        Call clearTile(tile(intObjX, intObjY), False, -1, "ObjTopRightXY")
    End If
    Call clearTile(tile(intObjX, intObjY), False, -1, "ObjXY")
End If
With frmMain
If Not killObj Then 'if object has not expired
    If strObjType = "Coin" Then
        'set coin pictures
        Set tile(intObjX, intObjY).objMask = .picCoinMask(intFrame + intFrameOffset)
        If strType = "Y" Then
            Set tile(intObjX, intObjY).picObj = .picCoinY(intFrame + intFrameOffset)
        ElseIf strType = "R" Then
            Set tile(intObjX, intObjY).picObj = .picCoinR(intFrame + intFrameOffset)
        ElseIf strType = "B" Then
            Set tile(intObjX, intObjY).picObj = .picCoinB(intFrame + intFrameOffset)
        End If
    ElseIf strObjType = "Pow" Then
        'set power-up pictures
        If strType = "Scare" Then
            Set tile(intObjX, intObjY).objMask = .picPowScareMask(intFrame + intFrameOffset)
            Set tile(intObjX, intObjY).picObj = .picPowScare(intFrame + intFrameOffset)
        ElseIf strType = "Speed" Then
            Set tile(intObjX, intObjY).objMask = .picPowSpeedMask(intFrame + intFrameOffset)
            Set tile(intObjX, intObjY).picObj = .picPowSpeed(intFrame + intFrameOffset)
        ElseIf strType = "Freeze" Then
            Set tile(intObjX, intObjY).objMask = .picPowFreezeMask
            Set tile(intObjX, intObjY).picObj = .picPowFreeze(intFrame + intFrameOffset)
        End If
    ElseIf strObjType = "Egg" Then
        'set egg pictures
        If strType = "M" Then
            Set tile(intObjX, intObjY).objMask = .picEggMask((intFrame \ 2) + (intFrameOffset \ 2))
            Set tile(intObjX, intObjY).picObj = .picEgg((intFrame \ 2) + (intFrameOffset \ 2))
        ElseIf strType = "G" Then
            Set tile(intObjX, intObjY).objMask = .picEggMask(((intFrame - (8 * Int(intFrame / 8))) \ 2) + (intFrameOffset \ 2))
            Set tile(intObjX, intObjY).picObj = .picEggG(((intFrame - (8 * Int(intFrame / 8))) \ 2) + (intFrameOffset \ 2))
        End If
    End If
End If
If intObjY > 0 Then
    If oddRow(intObjY) Then
        If intObjX <= mapWidth - 1 Then
            Call clearTile(tile(intObjX, intObjY - 1), False, -1, "ObjOddX-Y")
            'Call clearVoid(tile(intObjX, intObjY - 1), checkClearVoid(tile(intObjX, intObjY - 1), True, False), checkClearVoid(tile(intObjX, intObjY - 1), False, True))
        End If
        If intObjX > 0 Then
            Call clearTile(tile(intObjX - 1, intObjY - 1), False, -1, "ObjOdd-X-Y")
            'Call clearVoid(tile(intObjX - 1, intObjY - 1), checkClearVoid(tile(intObjX - 1, intObjY - 1), True, False), False)
        End If
    Else
        If intObjX <= mapWidth Then
            Call clearTile(tile(intObjX, intObjY - 1), False, -1, "ObjEvenX-Y")
            'Call clearVoid(tile(intObjX, intObjY - 1), checkClearVoid(tile(intObjX, intObjY - 1), True, False), checkClearVoid(tile(intObjX, intObjY - 1), False, True))
        End If
        If intObjX < mapWidth Then
            Call clearTile(tile(intObjX + 1, intObjY - 1), False, -1, "ObjEven+X-Y")
            'Call clearVoid(tile(intObjX + 1, intObjY - 1), False, checkClearVoid(tile(intObjX + 1, intObjY - 1), False, True))
        End If
    End If
Else
    Call clearVoid(tile(intObjX, intObjY), True, True, "ObjTopXY")
    Call clearTile(tile(intObjX, intObjY), False, -1, "ObjTopXY")
End If
For charPaint = 0 To 3
    If frmMain.tmrChar(charPaint).Enabled And nextX(charPaint) = intObjX And nextY(charPaint) = intObjY Then
        Call PaintCharSprite(charPaint, spriteX(charPaint), spriteY(charPaint), True)
    End If
Next charPaint
If Not killObj And strObjType <> "Terrain" Then
    If paintMask(tile(intObjX, intObjY), -1) Then 'if mask is to be painted, paint it
        .PaintPicture tile(intObjX, intObjY).objMask.Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
    End If
    'paint object
    .PaintPicture tile(intObjX, intObjY).picObj.Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    'for sparkling objects, paint sparkle effect
    If strObjType = "Coin" Or (strObjType = "Egg" And strType = "G") Then
        If intFrame > 11 And intFrame < 19 Then
            tile(intObjX, intObjY).objSparkle = True
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picSparkleMask(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 3), 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picSparkle(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 3), 100, 100, 0, 0, 100, 100, vbSrcPaint
        Else
            tile(intObjX, intObjY).objSparkle = False
        End If
    Else
        tile(intObjX, intObjY).objSparkle = False
    End If
End If
End With
End Function
Public Function PaintSelector(ByVal Index As Integer, ByVal imgIndex As Integer) As Integer
'paint over last sel
If blnClearPrevTile(Index) = True Then
    Call clearTile(tile(prevX(Index), prevY(Index)), True, Index, "SelPXPY") 'clear (prevX, prevY)
    blnClearPrevTile(Index) = False
End If
'paint over frame
Call clearTile(tile(curX(Index), curY(Index)), True, Index, "SelXY")
'paint sel
With frmMain
    If isPlayer(Index) Then
        .PaintPicture .picSelMask(imgIndex).Image, tile(curX(Index), curY(Index)).x, tile(curX(Index), curY(Index)).y, 100, 100, 0, 0, 100, 100, vbSrcAnd
        .PaintPicture .picSel(imgIndex).Image, tile(curX(Index), curY(Index)).x, tile(curX(Index), curY(Index)).y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    Else
        .PaintPicture .picSelMask(imgIndex + 5).Image, tile(curX(Index), curY(Index)).x, tile(curX(Index), curY(Index)).y, 100, 100, 0, 0, 100, 100, vbSrcAnd
        .PaintPicture .picSel(imgIndex + 5).Image, tile(curX(Index), curY(Index)).x, tile(curX(Index), curY(Index)).y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    End If
End With
End Function

Public Sub PaintCharSprite(ByVal Index As Integer, ByVal charX As Integer, ByVal charY As Integer, ByVal byPassClear As Boolean)
With frmMain
If Not byPassClear Then
    'clear tiles character may be touching
    If curY(Index) > 0 Then
        If oddRow(curY(Index)) Then 'odd row
            If curX(Index) <= mapWidth - 1 Then
                Call clearTile(tile(curX(Index), curY(Index) - 1), True, Index, "CharOddX-Y") 'clear (curX, curY - 1)
            End If
            If curX(Index) > 0 Then  'if column is less than last column
                Call clearTile(tile(curX(Index) - 1, curY(Index) - 1), True, Index, "CharOdd-X-Y") 'clear (curX - 1, curY - 1)
            End If
        Else 'even row
            If curX(Index) <= mapWidth Then 'if column is greater than first column
                Call clearTile(tile(curX(Index), curY(Index) - 1), True, Index, "CharEvenX-Y") 'clear (curX, curY - 1)
                If curX(Index) < mapWidth Then
                    Call clearTile(tile(curX(Index) + 1, curY(Index) - 1), True, Index, "CharEven+X-Y") 'clear (curX + 1, curY - 1)
                End If
            End If
        End If
    End If
    If (curX(Index) <> nextX(Index) Or curY(Index) <> nextY(Index)) Then
        Call clearTile(tile(nextX(Index), nextY(Index)), True, Index, "CharNXNY") 'clear (nextX, nextY)
        If nextY(Index) > 0 Then
            If strDir(Index) = "L" Then
                If oddRow(nextY(Index)) Then
                    If nextX(Index) <= mapWidth - 1 Then
                        Call clearTile(tile(nextX(Index), nextY(Index) - 1), True, Index, "CharNOddNXN-Y") 'clear (nextX, nextY - 1)
                    End If
                    If nextX(Index) > 0 Then
                        Call clearTile(tile(nextX(Index) - 1, nextY(Index) - 1), True, Index, "CharNOddN-XN-Y") 'clear (nextX - 1, nextY - 1)
                    End If
                Else
                    If nextY(Index) <= mapWidth Then
                        Call clearTile(tile(nextX(Index), nextY(Index) - 1), True, Index, "CharNEvenNXN-Y") 'clear (nextX, nextY - 1)
                    End If
                    If nextX(Index) < mapWidth - 1 Then
                        Call clearTile(tile(nextX(Index) + 1, nextY(Index) - 1), True, Index, "CharNEvenN+XN-Y") 'clear (nextX + 1, nextY - 1)
                    End If
                End If
            ElseIf strDir(Index) = "U" Then
                If oddRow(nextY(Index)) Then
                    If nextX(Index) < mapWidth Then
                        Call clearTile(tile(nextX(Index), nextY(Index) - 1), True, Index, "CharNOddNXN-Y") 'clear (nextX, nextY - 1)
                    End If
                    If nextX(Index) > 0 Then
                        Call clearTile(tile(nextX(Index) - 1, nextY(Index) - 1), True, Index, "CharNOddN-XN-Y") 'clear (nextX - 1, nextY - 1)
                    End If
                Else
                    If nextX(Index) <= mapWidth Then
                        Call clearTile(tile(nextX(Index), nextY(Index) - 1), True, Index, "CharNEvenNXN-Y") 'clear (nextX, nextY - 1)
                    End If
                    If nextX(Index) < mapWidth Then
                        Call clearTile(tile(nextX(Index) + 1, nextY(Index) - 1), True, Index, "CharNEvenN+XN-Y") 'clear (nextX + 1, nextY - 1)
                    End If
                End If
            ElseIf strDir(Index) = "R" Then
                If oddRow(nextY(Index)) Then
                    If (nextY(Index) < mapHeight - 1 And nextX(Index) < mapWidth) Or (nextY(Index) <> curY(Index) And nextX(Index) < mapWidth) Then
                        Call clearTile(tile(nextX(Index), nextY(Index) - 1), True, Index, "CharNOddNXN+Y") 'clear (nextX, nextY - 1)
                    End If
                Else
                    If nextX(Index) < mapWidth And nextY(Index) <> curY(Index) Then
                        Call clearTile(tile(nextX(Index) + 1, nextY(Index) - 1), True, Index, "CharNEvenN+XN+Y") 'clear (nextX + 1, nextY - 1)
                    End If
                End If
            ElseIf strDir(Index) = "D" Then
                If oddRow(nextY(Index)) Then
                    If nextX(Index) > 0 Then
                        Call clearTile(tile(nextX(Index) - 1, nextY(Index) - 1), True, Index, "CharNOddN-XN+Y") 'clear (nextX - 1, nextY - 1)
                    End If
                Else
                    If nextY(Index) < mapHeight - 1 Or nextY(Index) <> curY(Index) Then
                        Call clearTile(tile(nextX(Index), nextY(Index) - 1), True, Index, "CharNEvenNXN+Y") 'clear (nextX, nextY - 1)
                    End If
                End If
            End If
        Else
            If strDir(Index) = "L" Then
                Call clearVoid(tile(nextX(Index), nextY(Index)), True, True)
                Call clearTile(tile(nextX(Index), nextY(Index)), True, Index, "CharTopNXNY") 'clear top (nextX, nextY)
            ElseIf strDir(Index) = "U" Then
                Call clearVoid(tile(nextX(Index), nextY(Index)), True, True)
                Call clearTile(tile(nextX(Index), nextY(Index)), True, Index, "CharTopNXNY") 'clear top (nextX, nextY)
            End If
        End If
    End If
End If
Static counterC As Integer
If strState(Index) = "I" Then
    If strDir(Index) = "L" Then
        If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
            .PaintPicture .picCharMaskIL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If Index = 0 Then
            .PaintPicture .picP1IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 1 Then
            .PaintPicture .picP2IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 2 Then
            .PaintPicture .picP3IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 3 Then
            .PaintPicture .picP4IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "U" Then
        If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
            .PaintPicture .picCharMaskIU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If Index = 0 Then
            .PaintPicture .picP1IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 1 Then
            .PaintPicture .picP2IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 2 Then
            .PaintPicture .picP3IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 3 Then
            .PaintPicture .picP4IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "R" Then
        If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
            .PaintPicture .picCharMaskIR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If Index = 0 Then
            .PaintPicture .picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 1 Then
            .PaintPicture .picP2IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 2 Then
            .PaintPicture .picP3IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 3 Then
            .PaintPicture .picP4IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "D" Then
        If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
            .PaintPicture .picCharMaskID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If Index = 0 Then
            .PaintPicture .picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 1 Then
            .PaintPicture .picP2ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 2 Then
            .PaintPicture .picP3ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index = 3 Then
            .PaintPicture .picP4ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    End If
Else
    If strState(Index) = "C" Then
        If counterC < 3 Then
            counterC = counterC + 1
        End If
        Dim frameC As Integer
        frameC = Int((counterC + 1) / 2)
        If strDir(Index) = "L" Then
            If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
                .PaintPicture .picCharMaskCL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If Index = 0 Then
                .PaintPicture .picP1CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 1 Then
                .PaintPicture .picP2CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 2 Then
                .PaintPicture .picP3CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 3 Then
                .PaintPicture .picP4CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "U" Then
            If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
                .PaintPicture .picCharMaskCU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If Index = 0 Then
                .PaintPicture .picP1CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 1 Then
                .PaintPicture .picP2CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 2 Then
                .PaintPicture .picP3CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 3 Then
                .PaintPicture .picP4CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "R" Then
            If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
                .PaintPicture .picCharMaskCR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If Index = 0 Then
                .PaintPicture .picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 1 Then
                .PaintPicture .picP2CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 2 Then
                .PaintPicture .picP3CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 3 Then
                .PaintPicture .picP4CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "D" Then
            If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
                .PaintPicture .picCharMaskCD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If Index = 0 Then
                .PaintPicture .picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 1 Then
                .PaintPicture .picP2CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 2 Then
                .PaintPicture .picP3CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 3 Then
                .PaintPicture .picP4CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    ElseIf strState(Index) = "J" Then
        counterC = 0
        If strDir(Index) = "L" Then
            If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
                .PaintPicture .picCharMaskJL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If Index = 0 Then
                .PaintPicture .picP1JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 1 Then
                .PaintPicture .picP2JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 2 Then
                .PaintPicture .picP3JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 3 Then
                .PaintPicture .picP4JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "U" Then
            If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
                .PaintPicture .picCharMaskJU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If Index = 0 Then
                .PaintPicture .picP1JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 1 Then
                .PaintPicture .picP2JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 2 Then
                .PaintPicture .picP3JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 3 Then
                .PaintPicture .picP4JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "R" Then
            If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
                .PaintPicture .picCharMaskJR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If Index = 0 Then
                .PaintPicture .picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 1 Then
                .PaintPicture .picP2JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 2 Then
                .PaintPicture .picP3JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 3 Then
                .PaintPicture .picP4JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "D" Then
            If Not blnRecover(Index) Or paintMask(tile(curX(Index), curY(Index)), Index) Then
                .PaintPicture .picCharMaskJD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If Index = 0 Then
                .PaintPicture .picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 1 Then
                .PaintPicture .picP2JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 2 Then
                .PaintPicture .picP3JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index = 3 Then
                .PaintPicture .picP4JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    End If
End If
End With
End Sub

Public Sub getTileAnim(ByVal intFrame As Integer, tileInput As terrain)
With frmMain
    If ((tileInput.y - 400) + intFrame * 50) >= 0 Then
        .picBuffer.PaintPicture .picBackground.Image, 0, 0, 100, 100, tileInput.x, (tileInput.y - 400) + intFrame * 50, 100, 100, vbSrcCopy
        .PaintPicture .picBuffer.Image, tileInput.x, (tileInput.y - 400) + (intFrame - 1) * 50, 100, 100, 0, 0, 100, 100, vbSrcCopy
    End If
    'paint tile mask with new y
    .PaintPicture tileInput.picMask.Image, tileInput.x, (tileInput.y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcAnd
    'paint tile with new y
    .PaintPicture tileInput.picTile.Image, tileInput.x, (tileInput.y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcPaint
    If intFrame >= 8 And tileInput.Xc = tile(0, 0).Xc And tileInput.Yc = tile(0, 0).Yc Then
        Call gameStart
        .tmrTileAnim.Enabled = False
    End If
End With
End Sub
