Attribute VB_Name = "modPaint"
Option Explicit

Public Sub clearTile(TileInput As terrain, ByVal bypassForObj As Boolean, Optional index As Integer)
With frmMain
'if object not enabled on tile or bypassForObj is true (doesn't paint if bypassForObj is true and object is enabled on tile)
If Not TileInput.hasObj Or Not bypassForObj Then
    If TileInput.hasObj Then 'if object is enabled on tile
        'paint over object
        If TileInput.objType(0) = "Coin" Then
            'paint terrain over coin
            intObjxOffset = 41
            .picBackground.PaintPicture .picScene(0).Image, TileInput.x + intObjxOffset, TileInput.y, 18, 35, intObjxOffset, 0, 18, 35, vbSrcCopy
            .picBuffer.PaintPicture .picScene(0).Image, intObjxOffset, 0, 18, 35, intObjxOffset, 0, 18, 35, vbSrcCopy
            .PaintPicture .picMask.Image, TileInput.x + intObjxOffset, TileInput.y, 18, 35, intObjxOffset, 0, 18, 35, vbSrcAnd
            .PaintPicture .picBuffer.Image, TileInput.x + intObjxOffset, TileInput.y, 18, 35, intObjxOffset, 0, 18, 35, vbSrcPaint
        ElseIf TileInput.objType(0) = "Pow" Then
            If TileInput.objType(1) = "Scare" Then
                'paint terrain over scare power-up
                intObjxOffset = 35
                .picBackground.PaintPicture .picScene(0).Image, TileInput.x + intObjxOffset, TileInput.y, 30, 30, intObjxOffset, 0, 30, 30, vbSrcCopy
                .picBuffer.PaintPicture .picScene(0).Image, intObjxOffset, 0, 30, 30, intObjxOffset, 0, 30, 30, vbSrcCopy
                .PaintPicture .picMask.Image, TileInput.x + intObjxOffset, TileInput.y, 30, 30, intObjxOffset, 0, 30, 30, vbSrcAnd
                .PaintPicture .picBuffer.Image, TileInput.x + intObjxOffset, TileInput.y, 30, 30, intObjxOffset, 0, 30, 30, vbSrcPaint
            End If
        End If
    ElseIf Not bypassForObj And Not TileInput.hasObj Then 'called by object but no object on tile
        'paint bottom 30 of the tile
        .picBackground.PaintPicture .picScene(0).Image, TileInput.x, TileInput.y + 70, 100, 30, 0, 70, 100, 30, vbSrcCopy
        .picBuffer.PaintPicture .picScene(0).Image, 0, 0, 100, 30, 0, 70, 100, 30, vbSrcCopy
        .PaintPicture .picMask.Image, TileInput.x, TileInput.y + 70, 100, 30, 0, 70, 100, 30, vbSrcAnd
        .PaintPicture .picBuffer.Image, TileInput.x, TileInput.y + 70, 100, 30, 0, 0, 100, 30, vbSrcPaint
    Else 'called by character
        If TileInput.hasChar Then 'if character is on tile
            If (TileInput.Xc = curX(index) And TileInput.Yc = curY(index)) Then 'if character on tile is character that called clear
                If TileInput.Yc = 0 Then '(x, 0)
                    Call clearVoid(TileInput, True, True) 'paint spacer
                ElseIf oddRow(TileInput.Yc) Then '(x, odd)
                    If TileInput.Xc = 0 Then 'if first column
                        Call clearVoid(TileInput, True, False) 'paint right spacer
                    ElseIf TileInput.Xc = mapWidth Then 'if last column
                        Call clearVoid(TileInput, False, True) 'paint left spacer
                    End If
                End If
                If Not tileTouchingChar(TileInput) Then 'if not touching a char, paint full tile
                    'paint over tile
                    .picBackground.PaintPicture .picScene(0).Image, TileInput.x, TileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .picBuffer.PaintPicture .picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .PaintPicture .picMask.Image, TileInput.x, TileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                    .PaintPicture .picBuffer.Image, TileInput.x, TileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                Else 'if touching a char
                    'only paint over top half of tile
                    .picBackground.PaintPicture .picScene(0).Image, TileInput.x, TileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .picBuffer.PaintPicture .picScene(0).Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .PaintPicture .picMask.Image, TileInput.x, TileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                    .PaintPicture .picBuffer.Image, TileInput.x, TileInput.y, 100, 50, 0, 0, 100, 50, vbSrcPaint
                End If
            End If
        Else 'if not character on tile
            'if not touching character or coordinates match previous tile of character
            If Not tileTouchingChar(TileInput) Or (TileInput.Xc = prevX(index) And TileInput.Yc = prevY(index)) Then
                'paint over tile
                .picBackground.PaintPicture .picScene(0).Image, TileInput.x, TileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture .picScene(0).Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture .picMask.Image, TileInput.x, TileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, TileInput.x, TileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            Else 'if touching a char
                'only paint over top half of tile
                .picBackground.PaintPicture .picScene(0).Image, TileInput.x, TileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                .picBuffer.PaintPicture .picScene(0).Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                .PaintPicture .picMask.Image, TileInput.x, TileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                .PaintPicture .picBuffer.Image, TileInput.x, TileInput.y, 100, 50, 0, 0, 100, 50, vbSrcPaint
            End If
        End If
    End If
ElseIf bypassForObj And TileInput.hasObj Then
    'paint over bottom half of tile
    .picBackground.PaintPicture .picScene(0).Image, TileInput.x, TileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
    .picBuffer.PaintPicture .picScene(0).Image, 0, 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
    .PaintPicture .picMask.Image, TileInput.x, TileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcAnd
    .PaintPicture .picBuffer.Image, TileInput.x, TileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcPaint
End If
End With
End Sub

Public Sub clearVoid(TileInput As terrain, ByVal blnL As Boolean, ByVal blnR As Boolean) 'clear empty spots on map
With frmMain
    If blnL And blnR Then
        .picBackground.PaintPicture .picSpacer.Image, TileInput.x, TileInput.y, 100, 24, 0, 0, 100, 24, vbSrcCopy
        .picBuffer.PaintPicture .picSpacer.Image, 0, 0, 100, 24, 0, 0, 100, 24, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, TileInput.x, TileInput.y, 100, 24, 0, 0, 100, 24, vbSrcAnd
        .PaintPicture .picBuffer.Image, TileInput.x, TileInput.y, 100, 24, 0, 0, 100, 24, vbSrcPaint
    ElseIf blnL Then
        .picBackground.PaintPicture .picSpacer.Image, TileInput.x, TileInput.y, 50, 24, 0, 0, 50, 24, vbSrcCopy
        .picBuffer.PaintPicture .picSpacer.Image, 0, 0, 50, 24, 0, 0, 50, 24, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, TileInput.x, TileInput.y, 50, 24, 0, 0, 50, 24, vbSrcAnd
        .PaintPicture .picBuffer.Image, TileInput.x, TileInput.y, 50, 24, 0, 0, 50, 24, vbSrcPaint
    ElseIf blnR Then
        .picBackground.PaintPicture .picSpacer.Image, TileInput.x + 50, TileInput.y, 50, 24, 50, 0, 50, 24, vbSrcCopy
        .picBuffer.PaintPicture .picSpacer.Image, 50, 0, 50, 24, 50, 0, 50, 24, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, TileInput.x + 50, TileInput.y, 50, 24, 50, 0, 50, 24, vbSrcAnd
        .PaintPicture .picBuffer.Image, TileInput.x + 50, TileInput.y, 50, 24, 50, 0, 50, 24, vbSrcPaint
    End If
End With
End Sub

Public Function PaintObj(ByVal strObjType As String, ByVal strType As String, ByVal intFrame As Integer, ByVal intObjX As Integer, ByVal intObjY As Integer)
Dim intXOffset As Integer
Dim intYOffset As Integer
Dim intFrameOffset As Integer
If strObjType = "Coin" Then
    intXOffset = 41
    intYOffset = -1
    If intFrame > 13 Then
        intFrameOffset = -14
    Else
        intFrameOffset = 0
    End If
    Call clearTile(tile(intObjX, intObjY), False, -1)
    If intObjY > 0 Then
        If intObjX < mapWidth Then
            
        End If
        If oddRow(intObjY) Then
            If intObjX <= mapWidth - 1 Then
                Call clearTile(tile(intObjX, intObjY - 1), False, -1)
            End If
            If intObjX > 0 Then
                Call clearTile(tile(intObjX - 1, intObjY - 1), False, -1)
            End If
        Else
            If intObjX <= mapWidth Then
                Call clearTile(tile(intObjX, intObjY - 1), False, -1)
            End If
            If intObjX < mapWidth Then
                Call clearTile(tile(intObjX + 1, intObjY - 1), False, -1)
            End If
        End If
    End If
    'paint coin
    With frmMain
    .PaintPicture .picCoinMask(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
    If strType = "Y" Then
        .PaintPicture .picCoinY(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    ElseIf strType = "R" Then
        .PaintPicture .picCoinR(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    ElseIf strType = "B" Then
        .PaintPicture .picCoinB(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    End If
    'paint sparkle
    If intFrame > 12 And intFrame < 20 Then
        .PaintPicture .picSparkleMask(intFrame - 13).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcAnd
        .PaintPicture .picSparkle(intFrame - 13).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcPaint
    End If
    End With
ElseIf strObjType = "Pow" Then
    intXOffset = 35
    intYOffset = -1
    If intFrame > 4 Then
        intFrameOffset = (intFrame - (9 - intFrame)) * -1 'frames 5 - 9 will play in reverse
    Else
        intFrameOffset = 0
    End If
    Call clearTile(tile(intObjX, intObjY), False, -1)
    If intObjY > 0 Then
        Call clearTile(tile(intObjX, intObjY - 1), False, -1)
        If oddRow(intObjY) Then
            If intObjX > 0 Then
                Call clearTile(tile(intObjX - 1, intObjY - 1), False, -1)
            End If
        Else
            If intObjX < mapWidth - 1 Then
                Call clearTile(tile(intObjX + 1, intObjY - 1), False, -1)
            End If
        End If
    End If
    'paint power-up
    With frmMain
    If strType = "Scare" Then
        .PaintPicture .picPowScareMask(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
        .PaintPicture .picPowScare(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    End If
    End With
End If
End Function
Public Function PaintSelector(ByVal index As Integer, ByVal imgIndex As Integer) As Integer
'paint over last sel
If blnClearPrevTile(index) = True Then
    Call clearTile(tile(prevX(index), prevY(index)), True, index) 'clear (prevX, prevY)
    blnClearPrevTile(index) = False
End If
'paint over frame
Call clearTile(tile(curX(index), curY(index)), True, index)
'paint sel
With frmMain
    .PaintPicture .picSelMask(imgIndex + 5 * index).Image, tile(curX(index), curY(index)).x, tile(curX(index), curY(index)).y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    .PaintPicture .picSel(imgIndex + 5 * index).Image, tile(curX(index), curY(index)).x, tile(curX(index), curY(index)).y, 100, 100, 0, 0, 100, 100, vbSrcPaint
End With
Call PaintCharSprite(index, spriteX(index), spriteY(index))
End Function

Public Sub PaintCharSprite(ByVal index As Integer, ByVal charX As Integer, ByVal charY As Integer)
With frmMain
'clear tiles character may be touching
If curY(index) > 0 Then
    If (oddRow(curY(index)) And curX(index) <= mapWidth - 1) Or (Not oddRow(curY(index)) And curX(index) <= mapWidth) Then
        Call clearTile(tile(curX(index), curY(index) - 1), True, index) 'clear (curX, curY - 1)
    End If
    If oddRow((curY(index))) Then 'odd row
        If curX(index) < mapWidth - 1 Then  'if column is less than last column
            Call clearTile(tile(curX(index) + 1, curY(index) - 1), True, index) 'clear (curX - 1, curY - 1)
        End If
    Else 'even row
        If curX(index) > 0 Then 'if column is greater than first column
            Call clearTile(tile(curX(index) - 1, curY(index) - 1), True, index) 'clear (curX + 1, curY - 1)
        End If
    End If
End If
If (curX(index) <> nextX(index) Or curY(index) <> nextY(index)) Then
    Call clearTile(tile(nextX(index), nextY(index)), True, index) 'clear (nextX, nextY)
    If nextY(index) > 0 Then
        If strDir(index) = "U" Then
            If oddRow(nextY(index)) Then
                If nextX(index) <= mapWidth - 1 Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index) 'clear (nextX, nextY - 1)
                End If
                If nextX(index) < mapWidth - 1 Then
                    Call clearTile(tile(nextX(index) + 1, nextY(index) - 1), True, index) 'clear (nextX + 1, nextY - 1)
                End If
            Else
                If nextX(index) <= mapWidth Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index) 'clear (nextX, nextY - 1)
                End If
                If nextX(index) > 0 Then
                    Call clearTile(tile(nextX(index) - 1, nextY(index) - 1), True, index) 'clear (nextX - 1, nextY - 1)
                End If
            End If
        ElseIf strDir(index) = "R" Then
            If oddRow(nextY(index) - 1) Then
                If nextX(index) <= mapWidth - 1 Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index) 'clear (nextX, nextY - 1)
                End If
                If nextX(index) < mapWidth - 1 Then
                    Call clearTile(tile(nextX(index) + 1, nextY(index) - 1), True, index) 'clear (nextX + 1, nextY - 1)
                End If
            Else
                If nextX(index) <= mapWidth Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index) 'clear (nextX, nextY - 1)
                End If
                If nextX(index) > 0 Then
                    Call clearTile(tile(nextX(index) - 1, nextY(index) - 1), True, index) 'clear (nextX - 1, nextY - 1)
                End If
            End If
        End If
    End If
End If
Static counterC As Integer
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
        frameC = Int((counterC + 1) / 2)
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

Public Sub getTileAnim(ByVal intFrame As Integer, TileInput As terrain)
With frmMain
    .picBuffer.PaintPicture .picBackground.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
    .PaintPicture .picBuffer.Image, TileInput.x, (TileInput.y - 400) + (intFrame - 1) * 50, 100, 100, 0, 0, 100, 100, vbSrcCopy
    'paint tile mask with new y
    .PaintPicture .picMask.Picture, TileInput.x, (TileInput.y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcAnd
    'paint tile with new y
    .PaintPicture .picScene(0).Picture, TileInput.x, (TileInput.y - 400) + intFrame * 50, 100, 100, 0, 0, 100, 100, vbSrcPaint
    If intFrame >= 8 And TileInput.Xc = tile(0, 0).Xc And TileInput.Yc = tile(0, 0).Yc Then
        Call gameStart
        .tmrTileAnim.Enabled = False
    End If
End With
End Sub
