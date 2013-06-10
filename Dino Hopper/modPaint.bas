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
                'paint top of tile
                Call clearTileTop(tileInput)
            ElseIf tileInput.objType(0) = "Terrain" Then
                'paint over full tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf callID = "ObjTopXY" Then 'if call is for top of tile
            Call clearTileTop(tileInput) 'paint top of tile
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
                                Call clearVoid(tileAlt, False, True) 'paint right spacer
                            Else
                                tileAlt = tile(curX(Index), curY(Index) + 1)
                                Call clearVoid(tileAlt, True, False) 'paint left spacer
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
                            Call clearVoid(tileInput, True, True) 'paint spacer
                            If tileInput.Xc < mapWidth - 1 Then
                                tileAlt = tile(curX(Index) + 1, curY(Index))
                                Call clearVoid(tileAlt, True, False) 'paint left spacer
                            Else
                                tileAlt = tile(curX(Index), curY(Index) + 1)
                                Call clearVoid(tileAlt, False, True) 'paint right spacer
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
                        Call clearVoid(tileInput, True, True) 'paint spacer
                    End If
                ElseIf tileInput.Yc = mapHeight - 1 Then '(x, mapHeight - 1)
                    If frameCounter(Index) > 0 And strDir(Index) = "R" Then
                        If callID = "CharBottom+X+Y" Then
                            If tileInput.Xc = mapWidth - 1 Then
                                .PaintPicture picBG, tileInput.x + 50, tileInput.y, 50, 100, tileInput.x + 50, tileInput.y, 50, 100, vbSrcCopy
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
                                .PaintPicture picBG, tileInput.x - 50, tileInput.y, 50, 100, tileInput.x - 50, tileInput.y, 50, 100, vbSrcCopy
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
                                Call clearVoid(tileInput, True, False) 'paint left spacer
                                .PaintPicture picBG, tileInput.x, tileInput.y + 76, 50, 49, tileInput.x, tileInput.y + 76, 50, 49, vbSrcCopy
                            End If
                        Else
                            Call clearVoid(tileInput, True, False) 'paint left spacer
                        End If
                    ElseIf tileInput.Xc = mapWidth Then 'if last column
                        If frameCounter(Index) > 0 And strDir(Index) = "U" Then 'if jumping off right side edge
                            If callID <> "CharSide+X-Y" Then 'if jumping off right edge
                                Call clearVoid(tileInput, False, True) 'paint right spacer
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
                                Call clearVoid(tileInput, False, True) 'paint right spacer
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
                        Call clearTileTop(tileInput)
                    End If
                End If
            End If
        Else 'if not character on tile
            'if coordinates match previous tile of character
            If (tileInput.Xc = prevX(Index) And tileInput.Yc = prevY(Index)) Then 'paint top of tile if previous
                Call clearTileTop(tileInput)
            End If
            If callID = "CharEven+X-Y" Or callID = "CharOddX-Y" Then 'paint over bottom left of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 25, 50, 75, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 0, 25, 50, 75, vbSrcCopy
                .PaintPicture .picMaskSides.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
            ElseIf callID = "CharEvenX-Y" Or callID = "CharOdd-X-Y" Or callID = "CharNOddN-XN+Y" Or callID = "CharNEvenNXN+Y" Then 'paint over bottom right of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 50, 25, 50, 75, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 50, 25, 50, 75, vbSrcCopy
                .PaintPicture .picMaskSides.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 50, 0, 50, 75, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
            ElseIf callID = "CharTopNXNY" Or callID = "SelPXPY" Then 'paint over top of tile
                Call clearTileTop(tileInput)
            Else 'paint over full tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    End If
ElseIf bypassForObj And tileInput.hasObj Then
    'paint over bottom half of tile
    .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y + 25, 100, 75, 0, 25, 100, 75, vbSrcCopy
    .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 75, 0, 25, 100, 75, vbSrcCopy
    .PaintPicture .picMaskSides.Image, tileInput.x, tileInput.y + 25, 100, 75, 0, 0, 100, 75, vbSrcAnd
    .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 25, 100, 75, 0, 0, 100, 75, vbSrcPaint
    Call PaintObj(tileInput.objType(0), tileInput.objType(1), tileInput.objFrame, tileInput.Xc, tileInput.Yc, False)
End If
End With
End Sub

Public Sub clearVoid(tileInput As terrain, ByVal blnL As Boolean, ByVal blnR As Boolean) 'clear empty spots on map
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

Public Sub clearTileTop(tileInput As terrain)
Dim tilePic As Object
Set tilePic = tileInput.picTile
With frmMain
.picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
.picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
If tileInput.Yc = 0 Then
    .PaintPicture .picMaskTop.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
ElseIf tileInput.Xc = 0 And oddRow(tileInput.Yc) Then
    .PaintPicture .picMaskTop.Image, tileInput.x, tileInput.y, 50, 50, 0, 0, 50, 50, vbSrcAnd
    .PaintPicture .picMaskTopLR.Image, tileInput.x + 50, tileInput.y, 50, 50, 50, 0, 50, 50, vbSrcAnd
ElseIf tileInput.Xc = mapWidth Then
    .PaintPicture .picMaskTopLR.Image, tileInput.x, tileInput.y, 50, 50, 0, 0, 50, 50, vbSrcAnd
    .PaintPicture .picMaskTop.Image, tileInput.x + 50, tileInput.y, 50, 50, 50, 0, 50, 50, vbSrcAnd
Else
    .PaintPicture .picMaskTopLR.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
End If
.PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcPaint
End With
End Sub

Public Function PaintObj(ByVal strObjType As String, ByVal strType As String, ByVal intFrame As Integer, ByVal intObjX As Integer, ByVal intObjY As Integer, ByVal killObj As Boolean)
Dim intXOffset As Integer
Dim intYOffset As Integer
Dim intFrameOffset As Integer
If Not killObj Then 'if object has not expired
    If strObjType = "Coin" Then
        intXOffset = 41
        intYOffset = -1
        If intFrame > 13 Then
            intFrameOffset = -14
        Else
            intFrameOffset = 0
        End If
    ElseIf strObjType = "Pow" Then
        If strType = "Scare" Then
            intXOffset = 35
            intYOffset = -1
            If intFrame > 4 Then
                intFrameOffset = (intFrame - (9 - intFrame)) * -1 'frames 5 - 9 will play in reverse
            Else
                intFrameOffset = 0
            End If
        ElseIf strType = "Speed" Then
            intXOffset = 41
            intYOffset = -1
            intFrameOffset = 0
        ElseIf strType = "Freeze" Then
            intXOffset = 32
            intYOffset = -1
            If intFrame < 14 Then
                intFrameOffset = -intFrame
            Else
                intFrameOffset = -14
            End If
        End If
    ElseIf strObjType = "Egg" Then
        intXOffset = 39
        intYOffset = -1
        'multipliers are for gold eggs, where 3 animation cycles represents 1 frame cycle
        If intFrame - (8 * Int(intFrame / 8)) = 6 Or intFrame - (8 * Int(intFrame / 8)) = 7 Then
            'frames 6 and 7 will both use 2nd sprite (egg sprites play at half speed)
            intFrameOffset = -4
        Else
            intFrameOffset = 0
        End If
    End If
End If
Call clearTile(tile(intObjX, intObjY), False, -1, "ObjXY")
Call clearVoid(tile(intObjX, intObjY), checkClearVoid(tile(intObjX, intObjY), True, False), checkClearVoid(tile(intObjX, intObjY), False, True))
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
    Call clearVoid(tile(intObjX, intObjY), True, True)
    Call clearTile(tile(intObjX, intObjY), False, -1, "ObjTopXY")
End If
If Not killObj Then 'if object has not expired
    If strObjType = "Coin" Then
        'paint coin
        With frmMain
        If paintMask(tile(intObjX, intObjY), -1) Then
            .PaintPicture .picCoinMask(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If strType = "Y" Then
            .PaintPicture .picCoinY(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "R" Then
            .PaintPicture .picCoinR(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "B" Then
            .PaintPicture .picCoinB(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
        'paint sparkle
        If intFrame > 11 And intFrame < 19 Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picSparkleMask(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picSparkle(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
        End With
    ElseIf strObjType = "Pow" Then
        'paint power-up
        With frmMain
        If strType = "Scare" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picPowScareMask(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picPowScare(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "Speed" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picPowSpeedMask(intFrame).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picPowSpeed(intFrame).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "Freeze" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picPowFreezeMask.Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picPowFreeze(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
        End With
    ElseIf strObjType = "Egg" Then
        'paint egg
        With frmMain
        If strType = "M" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picEggMask((intFrame \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picEgg((intFrame \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "G" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picEggMask(((intFrame - (8 * Int(intFrame / 8))) \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picEggG(((intFrame - (8 * Int(intFrame / 8))) \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
            'paint sparkle
            If intFrame > 11 And intFrame < 19 Then
                If paintMask(tile(intObjX, intObjY), -1) Then
                    .PaintPicture .picSparkleMask(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcAnd
                End If
                .PaintPicture .picSparkle(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
        End With
    End If
End If
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

Public Sub PaintCharSprite(ByVal Index As Integer, ByVal charX As Integer, ByVal charY As Integer)
With frmMain
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
