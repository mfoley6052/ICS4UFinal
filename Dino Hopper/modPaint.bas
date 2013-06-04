Attribute VB_Name = "modPaint"
Option Explicit

Public Sub clearTile(tileInput As terrain, ByVal bypassForObj As Boolean, Optional index As Integer, Optional callID As String)
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
                .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                .PaintPicture .picMaskTop.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcPaint
            ElseIf tileInput.objType(0) = "Terrain" Then
                'paint over full tile
                .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf callID = "ObjTopXY" Then 'if call is for top of tile
            'paint top of tile
            .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcCopy
            .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
            .PaintPicture .picMaskTop.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcAnd
            .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcPaint
        Else 'no object on tile
            If callID = "ObjOddX-Y" Or callID = "ObjEven+X-Y" Then
                'paint over bottom left half of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y + 25, 50, 75, 0, 25, 50, 75, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 0, 25, 50, 75, vbSrcCopy
                .PaintPicture .picMaskSides.Image, tileInput.X, tileInput.Y + 25, 50, 75, 0, 0, 50, 75, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
            ElseIf callID = "ObjOdd-X-Y" Or callID = "ObjEvenX-Y" Then
                'paint over bottom right half of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.X + 50, tileInput.Y + 25, 50, 75, 50, 25, 50, 75, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 50, 25, 50, 75, vbSrcCopy
                .PaintPicture .picMaskSides.Image, tileInput.X + 50, tileInput.Y + 25, 50, 75, 50, 0, 50, 75, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.X + 50, tileInput.Y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
            End If
        End If
    Else 'called by character
        If tileInput.hasChar Then 'if character is on tile
            If (tileInput.Xc = curX(index) And tileInput.Yc = curY(index)) Then 'if character on tile is character that called clear
                Dim tileAlt As terrain
                If tileInput.Yc = 0 Then '(x, 0)
                    If frameCounter(index) > 0 And strDir(index) = "L" Then
                        If callID = "SelXY" Then
                            Call clearVoid(tileInput, True, True) 'paint spacer
                            If tileInput.Xc > 0 Then
                                tileAlt = tile(curX(index) - 1, curY(index))
                                Call clearVoid(tileAlt, False, True) 'paint right spacer
                            Else
                                tileAlt = tile(curX(index), curY(index) + 1)
                                Call clearVoid(tileAlt, True, False) 'paint left spacer
                                .PaintPicture picBG, tileInput.X - 50, tileInput.Y - 50, 50, 125, tileInput.X - 50, tileInput.Y - 50, 50, 125, vbSrcCopy
                            End If
                        Else
                            If tileInput.Xc > 0 Then
                                tileAlt = tile(curX(index) - 1, curY(index))
                            Else
                                tileAlt = tile(curX(index), curY(index) + 1)
                            End If
                            Set tilePic = tileAlt.picTile
                            .picBackground.PaintPicture tilePic.Image, tileAlt.X, tileAlt.Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                            .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                            .PaintPicture tileAlt.picMask.Image, tileAlt.X, tileAlt.Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                            .PaintPicture .picBuffer.Image, tileAlt.X, tileAlt.Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                        End If
                    ElseIf frameCounter(index) > 0 And strDir(index) = "U" Then
                        If callID = "SelXY" Then
                            Call clearVoid(tileInput, True, True) 'paint spacer
                            If tileInput.Xc < mapWidth - 1 Then
                                tileAlt = tile(curX(index) + 1, curY(index))
                                Call clearVoid(tileAlt, True, False) 'paint left spacer
                            Else
                                tileAlt = tile(curX(index), curY(index) + 1)
                                Call clearVoid(tileAlt, False, True) 'paint right spacer
                                .PaintPicture picBG, tileInput.X + 100, tileInput.Y - 50, 50, 125, tileInput.X + 100, tileInput.Y - 50, 50, 125, vbSrcCopy
                            End If
                        Else
                            If tileInput.Xc < mapWidth - 1 Then
                                tileAlt = tile(curX(index) + 1, curY(index))
                            Else
                                tileAlt = tile(curX(index), curY(index) + 1)
                            End If
                            Set tilePic = tileAlt.picTile
                            .picBackground.PaintPicture tilePic.Image, tileAlt.X, tileAlt.Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                            .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                            .PaintPicture tileAlt.picMask.Image, tileAlt.X, tileAlt.Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                            .PaintPicture .picBuffer.Image, tileAlt.X, tileAlt.Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                        End If
                    Else
                        Call clearVoid(tileInput, True, True) 'paint spacer
                    End If
                ElseIf tileInput.Yc = mapHeight - 1 Then '(x, mapHeight - 1)
                    If frameCounter(index) > 0 And strDir(index) = "R" Then
                        If callID = "CharBottom+X+Y" Then
                            If tileInput.Xc = mapWidth - 1 Then
                                .PaintPicture picBG, tileInput.X + 50, tileInput.Y, 50, 100, tileInput.X + 50, tileInput.Y, 50, 100, vbSrcCopy
                            Else
                                tileAlt = tile(curX(index) + 1, curY(index))
                                Set tilePic = tileAlt.picTile
                                'paint over bottom left half of tile
                                .picBackground.PaintPicture tilePic.Image, tileAlt.X, tileAlt.Y + 50, 50, 50, 0, 50, 50, 50, vbSrcCopy
                                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 74, 0, 26, 50, 74, vbSrcCopy
                                .PaintPicture tileAlt.picMask.Image, tileAlt.X, tileAlt.Y + 26, 50, 74, 0, 26, 50, 74, vbSrcAnd
                                .PaintPicture .picBuffer.Image, tileAlt.X, tileAlt.Y + 26, 50, 74, 0, 0, 50, 74, vbSrcPaint
                            End If
                        Else
                            .PaintPicture picBG, tileInput.X + 50, tileInput.Y + 76, 100, 49, tileInput.X + 50, tileInput.Y + 76, 100, 49, vbSrcCopy
                        End If
                    ElseIf frameCounter(index) > 0 And strDir(index) = "D" Then
                        If callID = "CharBottom-X+Y" Then
                            If tileInput.Xc = 0 Then
                                .PaintPicture picBG, tileInput.X - 50, tileInput.Y, 50, 100, tileInput.X - 50, tileInput.Y, 50, 100, vbSrcCopy
                            Else
                                tileAlt = tile(curX(index) - 1, curY(index))
                                Set tilePic = tileAlt.picTile
                                'paint over bottom right half of tile
                                .picBackground.PaintPicture tilePic.Image, tileAlt.X + 50, tileAlt.Y + 50, 50, 50, 50, 50, 50, 50, vbSrcCopy
                                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 74, 50, 26, 50, 74, vbSrcCopy
                                .PaintPicture tileAlt.picMask.Image, tileAlt.X + 50, tileAlt.Y + 26, 50, 74, 50, 26, 50, 74, vbSrcAnd
                                .PaintPicture .picBuffer.Image, tileAlt.X + 50, tileAlt.Y + 26, 50, 74, 0, 0, 50, 74, vbSrcPaint
                            End If
                        Else
                            .PaintPicture picBG, tileInput.X - 50, tileInput.Y + 76, 100, 49, tileInput.X - 50, tileInput.Y + 76, 100, 49, vbSrcCopy
                        End If
                    End If
                ElseIf oddRow(tileInput.Yc) Then '(x, odd)
                    If tileInput.Xc = 0 Then 'if first column
                        If frameCounter(index) > 0 And strDir(index) = "L" Then 'if jumping off left side edge
                            If callID <> "CharSide-X-Y" Then
                                Call clearVoid(tileInput, True, False) 'paint left spacer
                                .PaintPicture picBG, tileInput.X, tileInput.Y - 74, 50, 24, tileInput.X, tileInput.Y - 74, 50, 24, vbSrcCopy
                                If tileInput.Yc = 1 Then 'tilerow is 1
                                    .PaintPicture picBG, tileInput.X, tileInput.Y - 125, 50, 51, tileInput.X, tileInput.Y - 125, 50, 51, vbSrcCopy
                                End If
                            ElseIf callID = "CharSide-X-Y" Then
                                If tileInput.Yc > 1 Then 'tilerow greater than 1 (tile above in jump)
                                    tileAlt = tile(curX(index), curY(index) - 2)
                                    Set tilePic = tileAlt.picTile
                                    'paint over bottom left half of tile
                                    .picBackground.PaintPicture tilePic.Image, tileAlt.X, tileAlt.Y + 50, 50, 50, tileAlt.X, tileAlt.Y + 50, 50, 50, vbSrcCopy
                                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 0, 50, 50, 50, vbSrcCopy
                                    .PaintPicture tileAlt.picMask.Image, tileAlt.X, tileAlt.Y + 50, 50, 50, 0, 50, 50, 50, vbSrcAnd
                                    .PaintPicture .picBuffer.Image, tileAlt.X, tileAlt.Y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
                                Else 'tilerow is 1
                                    .PaintPicture picBG, tileInput.X, tileInput.Y - 124, 50, 49, tileInput.X, tileInput.Y - 124, 50, 49, vbSrcCopy
                                End If
                            End If
                        ElseIf frameCounter(index) > 0 And strDir(index) = "D" Then
                            If callID <> "CharSide-X+Y" Then
                                Call clearVoid(tileInput, True, False) 'paint left spacer
                                .PaintPicture picBG, tileInput.X, tileInput.Y + 76, 50, 49, tileInput.X, tileInput.Y + 76, 50, 49, vbSrcCopy
                            End If
                        Else
                            Call clearVoid(tileInput, True, False) 'paint left spacer
                        End If
                    ElseIf tileInput.Xc = mapWidth Then 'if last column
                        If frameCounter(index) > 0 And strDir(index) = "U" Then 'if jumping off right side edge
                            If callID <> "CharSide+X-Y" Then 'if jumping off right edge
                                Call clearVoid(tileInput, False, True) 'paint right spacer
                                .PaintPicture picBG, tileInput.X + 50, tileInput.Y - 74, 50, 24, tileInput.X + 50, tileInput.Y - 74, 50, 24, vbSrcCopy
                                If tileInput.Yc = 1 Then 'tilerow is 1
                                    .PaintPicture picBG, tileInput.X + 50, tileInput.Y - 125, 50, 51, tileInput.X + 50, tileInput.Y - 125, 50, 51, vbSrcCopy
                                End If
                            ElseIf callID = "CharSide+X-Y" Then
                                If tileInput.Yc > 1 Then
                                    tileAlt = tile(curX(index), curY(index) - 2)
                                    Set tilePic = tileAlt.picTile
                                    'paint over bottom right half of tile
                                    .picBackground.PaintPicture tilePic.Image, tileAlt.X + 50, tileAlt.Y + 50, 50, 50, tileAlt.X + 50, tileAlt.Y + 50, 50, 50, vbSrcCopy
                                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 50, 50, 50, 50, vbSrcCopy
                                    .PaintPicture tileAlt.picMask.Image, tileAlt.X + 50, tileAlt.Y + 50, 50, 50, 50, 50, 50, 50, vbSrcAnd
                                    .PaintPicture .picBuffer.Image, tileAlt.X + 50, tileAlt.Y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
                                Else
                                    .PaintPicture picBG, tileInput.X + 50, tileInput.Y - 124, 50, 49, tileInput.X + 50, tileInput.Y - 124, 50, 49, vbSrcCopy
                                End If
                            End If
                        ElseIf frameCounter(index) > 0 And strDir(index) = "R" Then
                            If callID <> "CharSide+X+Y" Then
                                Call clearVoid(tileInput, False, True) 'paint right spacer
                                .PaintPicture picBG, tileInput.X + 50, tileInput.Y + 76, 50, 49, tileInput.X + 50, tileInput.Y + 76, 50, 49, vbSrcCopy
                            End If
                        Else
                            Call clearVoid(tileInput, False, True) 'paint right spacer
                        End If
                    End If
                End If
                Set tilePic = tileInput.picTile
                If frameCounter(index) = 0 Or (curX(index) <> nextX(index) And curY(index) <> nextY(index)) Or callID = "SelXY" Then
                    If Not charTouchingTile(index, tileInput) Or frameCounter(index) > 0 Then 'if not touching a char or frame is greater than 0
                        'paint full tile
                        .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                        .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                        .PaintPicture tileInput.picMask.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                        .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                    Else 'if touching a char, only paint over top of tile
                        .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                        .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                        .PaintPicture .picMaskTopLR.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                        .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcPaint
                    End If
                End If
            End If
        Else 'if not character on tile
            'if coordinates match previous tile of character
            If (tileInput.Xc = prevX(index) And tileInput.Yc = prevY(index)) Then
                'paint over full tile
                .picBackground.PaintPicture tilePic, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            Else 'if not previous tile
                If callID = "CharEven+X-Y" Or callID = "CharOddX-Y" Then 'paint over bottom left of tile
                    .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y + 25, 50, 75, 0, 25, 50, 75, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 0, 25, 50, 75, vbSrcCopy
                    .PaintPicture .picMaskSides.Image, tileInput.X, tileInput.Y + 25, 50, 75, 0, 0, 50, 75, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
                ElseIf callID = "CharEvenX-Y" Or callID = "CharOdd-X-Y" Or callID = "CharNOddN-XN+Y" Or callID = "CharNEvenNXN+Y" Then 'paint over bottom right of tile
                    .picBackground.PaintPicture tilePic.Image, tileInput.X + 50, tileInput.Y + 25, 50, 75, 50, 25, 50, 75, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 75, 50, 25, 50, 75, vbSrcCopy
                    .PaintPicture .picMaskSides.Image, tileInput.X + 50, tileInput.Y + 25, 50, 75, 50, 0, 50, 75, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.X + 50, tileInput.Y + 25, 50, 75, 0, 0, 50, 75, vbSrcPaint
                ElseIf callID = "CharTopNXNY" Then
                    .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .PaintPicture .picMaskTop.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y, 100, 50, 0, 0, 100, 50, vbSrcPaint
                Else 'paint over full tile
                    .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                End If
            End If
        End If
    End If
ElseIf bypassForObj And tileInput.hasObj Then
    'paint over bottom half of tile
    .picBackground.PaintPicture tilePic.Image, tileInput.X, tileInput.Y + 25, 100, 75, 0, 25, 100, 75, vbSrcCopy
    .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 75, 0, 25, 100, 75, vbSrcCopy
    .PaintPicture .picMaskSides.Image, tileInput.X, tileInput.Y + 25, 100, 75, 0, 0, 100, 75, vbSrcAnd
    .PaintPicture .picBuffer.Image, tileInput.X, tileInput.Y + 25, 100, 75, 0, 0, 100, 75, vbSrcPaint
    Call PaintObj(tileInput.objType(0), tileInput.objType(1), tileInput.objFrame, tileInput.Xc, tileInput.Yc, False)
End If
End With
End Sub

Public Sub clearVoid(tileInput As terrain, ByVal blnL As Boolean, ByVal blnR As Boolean) 'clear empty spots on map
With frmMain
    If blnL And blnR Then
        .PaintPicture picBG, tileInput.X, tileInput.Y - 74, 100, 98, tileInput.X, tileInput.Y - 74, 100, 98, vbSrcCopy
    ElseIf blnL Then
        .PaintPicture picBG, tileInput.X, tileInput.Y - 50, 50, 74, tileInput.X, tileInput.Y - 50, 50, 74, vbSrcCopy
    ElseIf blnR Then
        .PaintPicture picBG, tileInput.X + 50, tileInput.Y - 50, 50, 74, tileInput.X + 50, tileInput.Y - 50, 50, 74, vbSrcCopy
    End If
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
'Call clearVoid(tile(intObjX, intObjY), checkClearVoid(tile(intObjX, intObjY), True, False), checkClearVoid(tile(intObjX, intObjY), False, True))
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
            .PaintPicture .picCoinMask(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If strType = "Y" Then
            .PaintPicture .picCoinY(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "R" Then
            .PaintPicture .picCoinR(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "B" Then
            .PaintPicture .picCoinB(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
        'paint sparkle
        If intFrame > 11 And intFrame < 19 Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picSparkleMask(intFrame - 12).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picSparkle(intFrame - 12).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
        End With
    ElseIf strObjType = "Pow" Then
        'paint power-up
        With frmMain
        If strType = "Scare" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picPowScareMask(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picPowScare(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "Speed" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picPowSpeedMask(intFrame).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picPowSpeed(intFrame).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "Freeze" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picPowFreezeMask.Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picPowFreeze(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
        End With
    ElseIf strObjType = "Egg" Then
        'paint egg
        With frmMain
        If strType = "M" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picEggMask((intFrame \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picEgg((intFrame \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "G" Then
            If paintMask(tile(intObjX, intObjY), -1) Then
                .PaintPicture .picEggMask(((intFrame - (8 * Int(intFrame / 8))) \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picEggG(((intFrame - (8 * Int(intFrame / 8))) \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
            'paint sparkle
            If intFrame > 11 And intFrame < 19 Then
                If paintMask(tile(intObjX, intObjY), -1) Then
                    .PaintPicture .picSparkleMask(intFrame - 12).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcAnd
                End If
                .PaintPicture .picSparkle(intFrame - 12).Image, tile(intObjX, intObjY).X + intXOffset, tile(intObjX, intObjY).Y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
        End With
    End If
End If
End Function
Public Function PaintSelector(ByVal index As Integer, ByVal imgIndex As Integer) As Integer
'paint over last sel
If blnClearPrevTile(index) = True Then
    Call clearTile(tile(prevX(index), prevY(index)), True, index, "SelPXPY") 'clear (prevX, prevY)
    blnClearPrevTile(index) = False
End If
'paint over frame
Call clearTile(tile(curX(index), curY(index)), True, index, "SelXY")
'paint sel
With frmMain
    If isPlayer(index) Then
        .PaintPicture .picSelMask(imgIndex).Image, tile(curX(index), curY(index)).X, tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
        .PaintPicture .picSel(imgIndex).Image, tile(curX(index), curY(index)).X, tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    Else
        .PaintPicture .picSelMask(imgIndex + 5).Image, tile(curX(index), curY(index)).X, tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcAnd
        .PaintPicture .picSel(imgIndex + 5).Image, tile(curX(index), curY(index)).X, tile(curX(index), curY(index)).Y, 100, 100, 0, 0, 100, 100, vbSrcPaint
    End If
End With
End Function

Public Sub PaintCharSprite(ByVal index As Integer, ByVal charX As Integer, ByVal charY As Integer)
With frmMain
'clear tiles character may be touching
If curY(index) > 0 Then
    If oddRow(curY(index)) Then 'odd row
        If curX(index) <= mapWidth - 1 Then
            Call clearTile(tile(curX(index), curY(index) - 1), True, index, "CharOddX-Y") 'clear (curX, curY - 1)
        End If
        If curX(index) > 0 Then  'if column is less than last column
            Call clearTile(tile(curX(index) - 1, curY(index) - 1), True, index, "CharOdd-X-Y") 'clear (curX - 1, curY - 1)
        End If
    Else 'even row
        If curX(index) <= mapWidth Then 'if column is greater than first column
            Call clearTile(tile(curX(index), curY(index) - 1), True, index, "CharEvenX-Y") 'clear (curX, curY - 1)
            If curX(index) < mapWidth Then
                Call clearTile(tile(curX(index) + 1, curY(index) - 1), True, index, "CharEven+X-Y") 'clear (curX + 1, curY - 1)
            End If
        End If
    End If
End If
If (curX(index) <> nextX(index) Or curY(index) <> nextY(index)) Then
    Call clearTile(tile(nextX(index), nextY(index)), True, index, "CharNXNY") 'clear (nextX, nextY)
    If nextY(index) > 0 Then
        If strDir(index) = "L" Then
            If oddRow(nextY(index)) Then
                If nextX(index) <= mapWidth - 1 Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index, "CharNOddNXN-Y") 'clear (nextX, nextY - 1)
                End If
                If nextX(index) > 0 Then
                    Call clearTile(tile(nextX(index) - 1, nextY(index) - 1), True, index, "CharNOddN-XN-Y") 'clear (nextX - 1, nextY - 1)
                End If
            Else
                If nextY(index) <= mapWidth Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index, "CharNEvenNXN-Y") 'clear (nextX, nextY - 1)
                End If
                If nextX(index) < mapWidth - 1 Then
                    Call clearTile(tile(nextX(index) + 1, nextY(index) - 1), True, index, "CharNEvenN+XN-Y") 'clear (nextX + 1, nextY - 1)
                End If
            End If
        ElseIf strDir(index) = "U" Then
            If oddRow(nextY(index)) Then
                If nextX(index) < mapWidth Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index, "CharNOddNXN-Y") 'clear (nextX, nextY - 1)
                End If
                If nextX(index) > 0 Then
                    Call clearTile(tile(nextX(index) - 1, nextY(index) - 1), True, index, "CharNOddN-XN-Y") 'clear (nextX - 1, nextY - 1)
                End If
            Else
                If nextX(index) <= mapWidth Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index, "CharNEvenNXN-Y") 'clear (nextX, nextY - 1)
                End If
                If nextX(index) < mapWidth Then
                    Call clearTile(tile(nextX(index) + 1, nextY(index) - 1), True, index, "CharNEvenN+XN-Y") 'clear (nextX + 1, nextY - 1)
                End If
            End If
        ElseIf strDir(index) = "R" Then
            If oddRow(nextY(index)) Then
                If (nextY(index) < mapHeight - 1 And nextX(index) < mapWidth) Or (nextY(index) <> curY(index) And nextX(index) < mapWidth) Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index, "CharNOddNXN+Y") 'clear (nextX, nextY - 1)
                End If
            Else
                If nextX(index) < mapWidth And nextY(index) <> curY(index) Then
                    Call clearTile(tile(nextX(index) + 1, nextY(index) - 1), True, index, "CharNEvenN+XN+Y") 'clear (nextX + 1, nextY - 1)
                End If
            End If
        ElseIf strDir(index) = "D" Then
            If oddRow(nextY(index)) Then
                If nextX(index) > 0 Then
                    Call clearTile(tile(nextX(index) - 1, nextY(index) - 1), True, index, "CharNOddN-XN+Y") 'clear (nextX - 1, nextY - 1)
                End If
            Else
                If nextY(index) < mapHeight - 1 Or nextY(index) <> curY(index) Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index, "CharNEvenNXN+Y") 'clear (nextX, nextY - 1)
                End If
            End If
        End If
    Else
        If strDir(index) = "L" Then
            Call clearVoid(tile(nextX(index), nextY(index)), True, True)
            Call clearTile(tile(nextX(index), nextY(index)), True, index, "CharTopNXNY") 'clear top (nextX, nextY)
        ElseIf strDir(index) = "U" Then
            Call clearVoid(tile(nextX(index), nextY(index)), True, True)
            Call clearTile(tile(nextX(index), nextY(index)), True, index, "CharTopNXNY") 'clear top (nextX, nextY)
        End If
    End If
End If
Static counterC As Integer
If strState(index) = "I" Then
    If strDir(index) = "L" Then
        If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
            .PaintPicture .picCharMaskIL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If index = 0 Then
            .PaintPicture .picP1IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 1 Then
            .PaintPicture .picP2IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 2 Then
            .PaintPicture .picP3IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 3 Then
            .PaintPicture .picP4IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "U" Then
        If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
            .PaintPicture .picCharMaskIU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If index = 0 Then
            .PaintPicture .picP1IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 1 Then
            .PaintPicture .picP2IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 2 Then
            .PaintPicture .picP3IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 3 Then
            .PaintPicture .picP4IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "R" Then
        If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
            .PaintPicture .picCharMaskIR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If index = 0 Then
            .PaintPicture .picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 1 Then
            .PaintPicture .picP2IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 2 Then
            .PaintPicture .picP3IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 3 Then
            .PaintPicture .picP4IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "D" Then
        If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
            .PaintPicture .picCharMaskID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        End If
        If index = 0 Then
            .PaintPicture .picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 1 Then
            .PaintPicture .picP2ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 2 Then
            .PaintPicture .picP3ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 3 Then
            .PaintPicture .picP4ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    End If
Else
    If strState(index) = "C" Then
        If counterC < 3 Then
            counterC = counterC + 1
        End If
        Dim frameC As Integer
        frameC = Int((counterC + 1) / 2)
        If strDir(index) = "L" Then
            If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
                .PaintPicture .picCharMaskCL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If index = 0 Then
                .PaintPicture .picP1CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 1 Then
                .PaintPicture .picP2CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 2 Then
                .PaintPicture .picP3CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 3 Then
                .PaintPicture .picP4CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(index) = "U" Then
            If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
                .PaintPicture .picCharMaskCU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If index = 0 Then
                .PaintPicture .picP1CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 1 Then
                .PaintPicture .picP2CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 2 Then
                .PaintPicture .picP3CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 3 Then
                .PaintPicture .picP4CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(index) = "R" Then
            If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
                .PaintPicture .picCharMaskCR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If index = 0 Then
                .PaintPicture .picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 1 Then
                .PaintPicture .picP2CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 2 Then
                .PaintPicture .picP3CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 3 Then
                .PaintPicture .picP4CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(index) = "D" Then
            If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
                .PaintPicture .picCharMaskCD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If index = 0 Then
                .PaintPicture .picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 1 Then
                .PaintPicture .picP2CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 2 Then
                .PaintPicture .picP3CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 3 Then
                .PaintPicture .picP4CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    ElseIf strState(index) = "J" Then
        counterC = 0
        If strDir(index) = "L" Then
            If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
                .PaintPicture .picCharMaskJL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If index = 0 Then
                .PaintPicture .picP1JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 1 Then
                .PaintPicture .picP2JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 2 Then
                .PaintPicture .picP3JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 3 Then
                .PaintPicture .picP4JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(index) = "U" Then
            If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
                .PaintPicture .picCharMaskJU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If index = 0 Then
                .PaintPicture .picP1JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 1 Then
                .PaintPicture .picP2JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 3 Then
                .PaintPicture .picP3JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 4 Then
                .PaintPicture .picP4JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(index) = "R" Then
            If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
                .PaintPicture .picCharMaskJR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If index = 0 Then
                .PaintPicture .picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 1 Then
                .PaintPicture .picP2JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 2 Then
                .PaintPicture .picP3JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 3 Then
                .PaintPicture .picP4JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(index) = "D" Then
            If Not blnRecover(index) Or paintMask(tile(curX(index), curY(index)), index) Then
                .PaintPicture .picCharMaskJD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            If index = 0 Then
                .PaintPicture .picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 1 Then
                .PaintPicture .picP2JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 2 Then
                .PaintPicture .picP3JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 3 Then
                .PaintPicture .picP4JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    End If
End If
End With
End Sub

Public Sub getTileAnim(ByVal intFrame As Integer, tileInput As terrain)
With frmMain
    If ((tileInput.Y - 400) + intFrame * 50) >= 0 Then
        .picBuffer.PaintPicture .picBackground.Image, 0, 0, 100, 100, tileInput.X, (tileInput.Y - 400) + intFrame * 50, 100, 100, vbSrcCopy
        .PaintPicture .picBuffer.Image, tileInput.X, (tileInput.Y - 400) + (intFrame - 1) * 50, 100, 100, 0, 0, 100, 100, vbSrcCopy
    End If
    'paint tile mask with new y
    .PaintPicture tileInput.picMask.Image, tileInput.X, (tileInput.Y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcAnd
    'paint tile with new y
    .PaintPicture tileInput.picTile.Image, tileInput.X, (tileInput.Y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcPaint
    If intFrame >= 8 And tileInput.Xc = tile(0, 0).Xc And tileInput.Yc = tile(0, 0).Yc Then
        Call gameStart
        .tmrTileAnim.Enabled = False
    End If
End With
End Sub
