Attribute VB_Name = "modPaint"
Option Explicit

Public Sub clearTile(ByVal intX As Integer, ByVal intY As Integer, ByVal bypassForCoin As Boolean, Optional index As Integer, Optional intCoinXOffset As Integer, Optional intCoinYOffset As Integer)
With frmMain
    'if coin not enabled on tile or bypassForCoin is true (doesn't paint if bypassForCoin is true and coin is enabled on tile)
    If Not tile(intX, intY).coinEnabled Or Not bypassForCoin Then
        'if coin is enabled on tile
        If tile(intX, intY).coinEnabled Then
            'paint over coin
            .picBackground.PaintPicture frmMain.picScene(0).Image, tile(intX, intY).x + intCoinXOffset, tile(intX, intY).y, 18, 35, intCoinXOffset, 0, 18, 35, vbSrcCopy
            .picBuffer.PaintPicture .picScene(0).Image, intCoinXOffset, 0, 18, 35, intCoinXOffset, 0, 18, 35, vbSrcCopy
            .PaintPicture .picMask.Image, tile(intX, intY).x + intCoinXOffset, tile(intX, intY).y, 18, 35, intCoinXOffset, 0, 18, 35, vbSrcAnd
            .PaintPicture .picBuffer.Image, tile(intX, intY).x + intCoinXOffset, tile(intX, intY).y, 18, 35, intCoinXOffset, 0, 18, 35, vbSrcPaint
        Else
            'if character is on tile
            If tile(intX, intY).hasChar Then
                'if character on tile is character that called clear
                If (intX = curX(index) And intY = curY(index)) Then
                    '(x, 0)
                    If intY = 0 Then
                        'paint spacer
                        Call clearVoid(intX, intY, True, True)
                    '(x, odd)
                    ElseIf oddRow(intY) Then
                        If intX = 0 Then 'if first column
                            'paint right spacer
                            Call clearVoid(intX, intY, True, False)
                        'last column
                        ElseIf intX = mapWidth Then
                            'paint left spacer
                            Call clearVoid(intX, intY, False, True)
                        End If
                    End If
                    If Not tileTouchingChar(intX, intY) Then 'if not touching a char, paint full tile
                        'paint over tile
                        .picBackground.PaintPicture .picScene(0).Image, tile(intX, intY).x, tile(intX, intY).y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                        .picBuffer.PaintPicture .picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                        .PaintPicture .picMask.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                        .PaintPicture .picBuffer.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                    Else 'if touching a char
                        'only paint over top half of tile
                        .picBackground.PaintPicture .picScene(0).Image, tile(intX, intY).x, tile(intX, intY).y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                        .picBuffer.PaintPicture .picScene(0).Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                        .PaintPicture .picMask.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                        .PaintPicture .picBuffer.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 50, 0, 0, 100, 50, vbSrcPaint
                    End If
                End If
            Else 'if no character on tile
                If Not tileTouchingChar(intX, intY) Or (intX = prevX(index) And intY = prevY(index)) Then 'if not touching a char, paint full tile
                    'paint over tile
                    .picBackground.PaintPicture .picScene(0).Image, tile(intX, intY).x, tile(intX, intY).y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .picBuffer.PaintPicture .picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .PaintPicture .picMask.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                Else 'if touching a char
                    'only paint over top half of tile
                    .picBackground.PaintPicture .picScene(0).Image, tile(intX, intY).x, tile(intX, intY).y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .picBuffer.PaintPicture .picScene(0).Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .PaintPicture .picMask.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 50, 0, 0, 100, 50, vbSrcPaint
                End If
            End If
        End If
    ElseIf bypassForCoin And tile(intX, intY).coinEnabled Then
        'paint over bottom half of tile
        .picBackground.PaintPicture .picScene(0).Image, tile(intX, intY).x, tile(intX, intY).y + 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
        .picBuffer.PaintPicture .picScene(0).Image, 0, 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
        .PaintPicture .picMask.Image, tile(intX, intY).x, tile(intX, intY).y + 50, 100, 50, 0, 50, 100, 50, vbSrcAnd
        .PaintPicture .picBuffer.Image, tile(intX, intY).x, tile(intX, intY).y + 50, 100, 50, 0, 50, 100, 50, vbSrcPaint
    End If
End With
End Sub
Public Sub clearVoid(ByVal intX As Integer, ByVal intY As Integer, ByVal blnL As Boolean, ByVal blnR As Boolean) 'clear empty spots on map
With frmMain
    If blnL And blnR Then
        .picBackground.PaintPicture .picSpacer.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 24, 0, 0, 100, 24, vbSrcCopy
        .picBuffer.PaintPicture .picSpacer.Image, 0, 0, 100, 24, 0, 0, 100, 24, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 24, 0, 0, 100, 24, vbSrcAnd
        .PaintPicture .picBuffer.Image, tile(intX, intY).x, tile(intX, intY).y, 100, 24, 0, 0, 100, 24, vbSrcPaint
    ElseIf blnL Then
        .picBackground.PaintPicture .picSpacer.Image, tile(intX, intY).x, tile(intX, intY).y, 50, 24, 0, 0, 50, 24, vbSrcCopy
        .picBuffer.PaintPicture .picSpacer.Image, 0, 0, 50, 24, 0, 0, 50, 24, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, tile(intX, intY).x, tile(intX, intY).y, 50, 24, 0, 0, 50, 24, vbSrcAnd
        .PaintPicture .picBuffer.Image, tile(intX, intY).x, tile(intX, intY).y, 50, 24, 0, 0, 50, 24, vbSrcPaint
    ElseIf blnR Then
        .picBackground.PaintPicture .picSpacer.Image, tile(intX, intY).x + 50, tile(intX, intY).y, 50, 24, 50, 0, 50, 24, vbSrcCopy
        .picBuffer.PaintPicture .picSpacer.Image, 50, 0, 50, 24, 50, 0, 50, 24, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, tile(intX, intY).x + 50, tile(intX, intY).y, 50, 24, 50, 0, 50, 24, vbSrcAnd
        .PaintPicture .picBuffer.Image, tile(intX, intY).x + 50, tile(intX, intY).y, 50, 24, 50, 0, 50, 24, vbSrcPaint
    End If
End With
End Sub

Public Function PaintCoin(ByVal strType As String, ByVal intFrame As Integer, ByVal intCoinX As Integer, ByVal intCoinY As Integer)
Dim intXOffset As Integer
Dim intYOffset As Integer
Dim intFrameOffset As Integer
intXOffset = 41
intYOffset = -1
If intFrame > 13 Then
    intFrameOffset = -14
End If
Call clearTile(intCoinX, intCoinY, False, -1, intXOffset, intYOffset)
'paint coin
With frmMain
    .PaintPicture .picCoinMask(intFrame + intFrameOffset).Image, tile(intCoinX, intCoinY).x + intXOffset, tile(intCoinX, intCoinY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
    If strType = "Y" Then
        .PaintPicture .picCoinY(intFrame + intFrameOffset).Image, tile(intCoinX, intCoinY).x + intXOffset, tile(intCoinX, intCoinY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    ElseIf strType = "R" Then
        .PaintPicture .picCoinR(intFrame + intFrameOffset).Image, tile(intCoinX, intCoinY).x + intXOffset, tile(intCoinX, intCoinY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    ElseIf strType = "B" Then
        .PaintPicture .picCoinB(intFrame + intFrameOffset).Image, tile(intCoinX, intCoinY).x + intXOffset, tile(intCoinX, intCoinY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    End If
    'paint sparkle
    If intFrame > 12 And intFrame < 20 Then
        .PaintPicture .picSparkleMask(intFrame - 13).Image, tile(intCoinX, intCoinY).x + intXOffset, tile(intCoinX, intCoinY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcAnd
        .PaintPicture .picSparkle(intFrame - 13).Image, tile(intCoinX, intCoinY).x + intXOffset, tile(intCoinX, intCoinY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcPaint
    End If
End With
End Function

Public Function PaintSelector(ByVal index As Integer, ByVal imgIndex As Integer) As Integer
'paint over last sel
If blnClearPrevTile(index) = True Then
    Call clearTile(prevX(index), prevY(index), True, index) 'clear (prevX, prevY)
    blnClearPrevTile(index) = False
End If
'paint over frame
Call clearTile(curX(index), curY(index), True, index)
'paint sel
With frmMain
    .PaintPicture .picSelMask(imgIndex + 5 * index).Image, tile(curX(index), curY(index)).x, tile(curX(index), curY(index)).y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    .PaintPicture .picSel(imgIndex + 5 * index).Image, tile(curX(index), curY(index)).x, tile(curX(index), curY(index)).y, 100, 100, 0, 0, 100, 100, vbSrcPaint
End With
Call PaintCharSprite(index, spriteX(index), spriteY(index))
End Function

Public Sub PaintCharSprite(ByVal index As Integer, ByVal charX As Integer, ByVal charY As Integer)
'clear tiles character may be touching
If curY(index) > 0 Then
    If (oddRow((curY(index) - 1)) And curX(index) < mapWidth) Or (Not oddRow((curY(index) - 1)) And curX(index) < mapWidth - 1) Then
        Call clearTile(curX(index), curY(index) - 1, True, index) 'clear (curX, curY - 1)
    End If
    If oddRow((curY(index) - 1)) Then 'odd row
        If curX(index) > 0 Then 'if column is greater than first column
            Call clearTile(curX(index) + 1, curY(index) - 1, True, index) 'clear (curX - 1, curY - 1)
        End If
    Else 'even row
        If curX(index) < mapWidth - 1 Then 'if column is less than last column
            Call clearTile(curX(index) - 1, curY(index) - 1, True, index) 'clear (curX + 1, curY - 1)
        End If
    End If
End If
If (curX(index) <> nextX(index) Or curY(index) <> nextY(index)) Then
    Call clearTile(nextX(index), nextY(index), True, index) 'clear (nextX, nextY)
End If
Static counterC As Integer
With frmMain
    If strState(index) = "I" Then
        If strDir(index) = "L" Then
            .PaintPicture .picCharMaskIL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If index = 0 Then
                .PaintPicture .picP1IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index > 0 Then
                .PaintPicture .picP1IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(index) = "U" Then
            .PaintPicture .picCharMaskIU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If index = 0 Then
                .PaintPicture .picP1IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index > 0 Then
                .PaintPicture .picP1IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(index) = "R" Then
            .PaintPicture .picCharMaskIR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If index = 0 Then
                .PaintPicture .picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index > 0 Then
                .PaintPicture .picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(index) = "D" Then
            .PaintPicture .picCharMaskID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If index = 0 Then
                .PaintPicture .picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index > 0 Then
                .PaintPicture .picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    Else
        If strState(index) = "C" Then
            If counterC < 3 Then
                counterC = counterC + 1
            End If
            Dim frameC As Integer
            frameC = Int((counterC + 1) / 2) 'Doesnt this just make frameC = counterC?
            If strDir(index) = "L" Then
                .PaintPicture .picCharMaskCL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
                If index = 0 Then
                    .PaintPicture .picP1CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                ElseIf index > 0 Then
                    .PaintPicture .picP1CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                End If
            ElseIf strDir(index) = "U" Then
                .PaintPicture .picCharMaskCU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
                If index = 0 Then
                    .PaintPicture .picP1CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                ElseIf index > 0 Then
                    .PaintPicture .picP1CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                End If
            ElseIf strDir(index) = "R" Then
                .PaintPicture .picCharMaskCR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
                If index = 0 Then
                    .PaintPicture .picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                ElseIf index > 0 Then
                    .PaintPicture .picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                End If
            ElseIf strDir(index) = "D" Then
                .PaintPicture .picCharMaskCD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
                If index = 0 Then
                    .PaintPicture .picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                ElseIf index > 0 Then
                    .PaintPicture .picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                End If
            End If
            ElseIf strState(index) = "J" Then
            counterC = 0
            If strDir(index) = "L" Then
                .PaintPicture .picCharMaskJL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
                If index = 0 Then
                    .PaintPicture .picP1JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                ElseIf index > 0 Then
                    .PaintPicture .picP1JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                End If
            ElseIf strDir(index) = "U" Then
                .PaintPicture .picCharMaskJU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
                If index = 0 Then
                    .PaintPicture .picP1JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                ElseIf index > 0 Then
                    .PaintPicture .picP1JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                End If
            ElseIf strDir(index) = "R" Then
                .PaintPicture .picCharMaskJR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
                If index = 0 Then
                    .PaintPicture .picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                ElseIf index > 0 Then
                    .PaintPicture .picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                End If
            ElseIf strDir(index) = "D" Then
                .PaintPicture .picCharMaskJD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
                If index = 0 Then
                    .PaintPicture .picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                ElseIf index > 0 Then
                    .PaintPicture .picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
                End If
            End If
        End If
    End If
End With
End Sub
Public Function getTileAnim(ByVal intFrame As Integer, ByVal intX As Integer, ByVal intY As Integer)
With frmMain
    .picBuffer.PaintPicture .picBackground.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
    .PaintPicture .picBuffer.Image, tile(intX, intY).x, (tile(intX, intY).y - 400) + (intFrame - 1) * 50, 100, 100, 0, 0, 100, 100, vbSrcCopy
    'paint tile mask with new y
    .PaintPicture .picMask.Picture, tile(intX, intY).x, (tile(intX, intY).y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcAnd
    'paint tile with new y
    .PaintPicture .picScene(0).Picture, tile(intX, intY).x, (tile(intX, intY).y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcPaint
    If intFrame >= 8 And intX = 0 And intY = 0 Then
        Call gameStart
        .tmrTileAnim.Enabled = False
    End If
End With
End Function
