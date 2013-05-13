Attribute VB_Name = "modPaint"
Option Explicit

Public Sub clearTile(tileInput As terrain, ByVal bypassForObj As Boolean, Optional Index As Integer, Optional callID As String)
With frmMain
'if object not enabled on tile or bypassForObj is true (doesn't paint if bypassForObj is true and object is enabled on tile)
If Not tileInput.hasObj Or Not bypassForObj Then
    If Not bypassForObj Then 'if object is caller of function
        If tileInput.hasObj Then
            'paint over object
            If tileInput.objType(0) = "Coin" Then
                'paint terrain over coin
                intObjxOffset = 41
                .picBackground.PaintPicture tileInput.picTile.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 35, intObjxOffset, 0, 18, 35, vbSrcCopy
                .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 18, 35, intObjxOffset, 0, 18, 35, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 35, intObjxOffset, 0, 18, 35, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 35, 0, 0, 18, 35, vbSrcPaint
            ElseIf tileInput.objType(0) = "Pow" Then
                If tileInput.objType(1) = "Scare" Then
                    'paint terrain over scare power-up
                    intObjxOffset = 35
                    .picBackground.PaintPicture tileInput.picTile.Image, tileInput.x + intObjxOffset, tileInput.y, 30, 30, intObjxOffset, 0, 30, 30, vbSrcCopy
                    .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 30, 30, intObjxOffset, 0, 30, 30, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 30, 30, intObjxOffset, 0, 30, 30, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x + intObjxOffset, tileInput.y, 30, 30, 0, 0, 30, 30, vbSrcPaint
                ElseIf tileInput.objType(1) = "Speed" Then
                    'paint terrain over speed power-up
                    intObjxOffset = 41
                    .picBackground.PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 33, intObjxOffset, 0, 18, 33, vbSrcCopy
                    .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 18, 33, intObjxOffset, 0, 18, 33, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 33, intObjxOffset, 0, 18, 33, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 33, 0, 0, 18, 33, vbSrcPaint
                End If
            ElseIf tileInput.objType(0) = "Egg" Then
                'paint terrain over egg
                intObjxOffset = 39
                .picBackground.PaintPicture tileInput.picTile.Image, tileInput.x + intObjxOffset, tileInput.y, 22, 30, intObjxOffset, 0, 22, 30, vbSrcCopy
                .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 22, 30, intObjxOffset, 0, 22, 30, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 22, 30, intObjxOffset, 0, 22, 30, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x + intObjxOffset, tileInput.y, 22, 30, 0, 0, 22, 30, vbSrcPaint
            End If
        Else 'no object on tile
            If callID = "ObjOdd+X-Y" Or callID = "ObjEvenX-Y" Then
                '.picBackground.PaintPicture tileInput.picTile.Image, tileInput.x, tileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
                '.picBuffer.PaintPicture tileInput.picTile.Image, 0, 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
                '.PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcAnd
                '.PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcPaint
                'paint over bottom right half of tile
                .picBackground.PaintPicture tileInput.picTile.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcCopy
                .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 50, 50, 50, 50, 50, 50, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcPaint
            ElseIf callID = "ObjOddX-Y" Or callID = "ObjEven-X-Y" Then
                'paint over bottom left half of tile
                .picBackground.PaintPicture tileInput.picTile.Image, tileInput.x, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcCopy
                .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 50, 50, 50, 50, 50, 50, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcPaint
            End If
        End If
    Else 'called by character
        If tileInput.hasChar Then 'if character is on tile
            If (tileInput.Xc = curX(Index) And tileInput.Yc = curY(Index)) Then 'if character on tile is character that called clear
                If tileInput.Yc = 0 Then '(x, 0)
                    Call clearVoid(tileInput, True, True) 'paint spacer
                ElseIf oddRow(tileInput.Yc) Then '(x, odd)
                    If tileInput.Xc = 0 Then 'if first column
                        Call clearVoid(tileInput, True, False) 'paint right spacer
                    ElseIf tileInput.Xc = mapWidth Then 'if last column
                        Call clearVoid(tileInput, False, True) 'paint left spacer
                    End If
                End If
                If Not tileTouchingChar(tileInput) Then 'if not touching a char, paint full tile
                    'paint over tile
                    .picBackground.PaintPicture tileInput.picTile.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                Else 'if touching a char
                    'only paint over top half of tile
                    .picBackground.PaintPicture tileInput.picTile.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcPaint
                End If
            End If
        Else 'if not character on tile
            'if not touching character or coordinates match previous tile of character
            If Not tileTouchingChar(tileInput) Or (tileInput.Xc = prevX(Index) And tileInput.Yc = prevY(Index)) Then
                'paint over tile
                .picBackground.PaintPicture tileInput.picTile, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            Else 'if touching a char
                'only paint over top half of tile
                '.picBackground.PaintPicture tileInput.picTile.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                '.picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                '.PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                '.PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcPaint
                'paint over tile
                .picBackground.PaintPicture tileInput.picTile.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tileInput.picTile.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    End If
ElseIf bypassForObj And tileInput.hasObj Then
    'paint over bottom half of tile
    '.picBackground.PaintPicture tileInput.picTile.Image, tileInput.x, tileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
    '.picBuffer.PaintPicture tileInput.picTile.Image, 0, 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
    '.PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcAnd
    '.PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcPaint
End If
End With
End Sub

Public Sub clearVoid(tileInput As terrain, ByVal blnL As Boolean, ByVal blnR As Boolean) 'clear empty spots on map
With frmMain
    If blnL And blnR Then
        .picBackground.PaintPicture .picSpacer.Image, tileInput.x, tileInput.y, 100, 24, 0, 0, 100, 24, vbSrcCopy
        .picBuffer.PaintPicture .picSpacer.Image, 0, 0, 100, 24, 0, 0, 100, 24, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, tileInput.x, tileInput.y, 100, 24, 0, 0, 100, 24, vbSrcAnd
        .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 24, 0, 0, 100, 24, vbSrcPaint
    ElseIf blnL Then
        .picBackground.PaintPicture .picSpacer.Image, tileInput.x, tileInput.y, 50, 24, 0, 0, 50, 24, vbSrcCopy
        .picBuffer.PaintPicture .picSpacer.Image, 0, 0, 50, 24, 0, 0, 50, 24, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, tileInput.x, tileInput.y, 50, 24, 0, 0, 50, 24, vbSrcAnd
        .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 50, 24, 0, 0, 50, 24, vbSrcPaint
    ElseIf blnR Then
        .picBackground.PaintPicture .picSpacer.Image, tileInput.x + 50, tileInput.y, 50, 24, 50, 0, 50, 24, vbSrcCopy
        .picBuffer.PaintPicture .picSpacer.Image, 50, 0, 50, 24, 50, 0, 50, 24, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, tileInput.x + 50, tileInput.y, 50, 24, 50, 0, 50, 24, vbSrcAnd
        .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y, 50, 24, 50, 0, 50, 24, vbSrcPaint
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
    End If
ElseIf strObjType = "Egg" Then
    intXOffset = 39
    intYOffset = -1
    If intFrame = 6 Or intFrame = 7 Then
        intFrameOffset = -4 'frames 6 and 7 will both use 2nd sprite (egg sprites play at half speed)
    Else
        intFrameOffset = 0
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
End If

If strObjType = "Coin" Then
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
    'paint power-up
    With frmMain
    If strType = "Scare" Then
        .PaintPicture .picPowScareMask(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
        .PaintPicture .picPowScare(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    ElseIf strType = "Speed" Then
        .PaintPicture .picPowSpeedMask(intFrame).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
        .PaintPicture .picPowSpeed(intFrame).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    End If
    End With
ElseIf strObjType = "Egg" Then
    'paint egg
    With frmMain
    .PaintPicture .picEggMask((intFrame \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
    .PaintPicture .picEgg((intFrame \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
    End With
End If
End Function
Public Function PaintSelector(ByVal Index As Integer, ByVal imgIndex As Integer) As Integer
'paint over last sel
If blnClearPrevTile(Index) = True Then
    Call clearTile(tile(prevX(Index), prevY(Index)), True, Index) 'clear (prevX, prevY)
    blnClearPrevTile(Index) = False
End If
'paint over frame
Call clearTile(tile(curX(Index), curY(Index)), True, Index, "SelXY")
'paint sel
With frmMain
    .PaintPicture .picSelMask(imgIndex + 5 * Index).Image, tile(curX(Index), curY(Index)).x, tile(curX(Index), curY(Index)).y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    .PaintPicture .picSel(imgIndex + 5 * Index).Image, tile(curX(Index), curY(Index)).x, tile(curX(Index), curY(Index)).y, 100, 100, 0, 0, 100, 100, vbSrcPaint
End With
Call PaintCharSprite(Index, spriteX(Index), spriteY(Index))
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
    If strDir(Index) = "D" Then
        If oddRow(nextY(Index)) Then
        Else
        End If
    ElseIf nextY(Index) > 0 Then
        If strDir(Index) = "U" Then
            If oddRow(nextY(Index)) Then
                If nextX(Index) <= mapWidth - 1 Then
                    Call clearTile(tile(nextX(Index), nextY(Index) - 1), True, Index, "CharNOddNXN-Y") 'clear (nextX, nextY - 1)
                End If
                If nextX(Index) < mapWidth - 1 Then
                    Call clearTile(tile(nextX(Index) + 1, nextY(Index) - 1), True, Index, "CharNOddN+XN-Y") 'clear (nextX + 1, nextY - 1)
                End If
            Else
                If nextX(Index) <= mapWidth Then
                    Call clearTile(tile(nextX(Index), nextY(Index) - 1), True, Index, "CharNEvenNXN-Y") 'clear (nextX, nextY - 1)
                End If
                If nextX(Index) > 0 Then
                    Call clearTile(tile(nextX(Index) - 1, nextY(Index) - 1), True, Index, "CharNEvenN-XN-Y") 'clear (nextX - 1, nextY - 1)
                End If
            End If
        ElseIf strDir(Index) = "R" Then
            If oddRow(nextY(Index) - 1) Then
                If nextX(Index) > 0 Then
                    Call clearTile(tile(nextX(Index) - 1, nextY(Index) - 1), True, Index, "CharNOddN-XN-Y") 'clear (nextX - 1, nextY - 1)
                End If
            Else
                If nextX(Index) < mapWidth - 1 Then
                    Call clearTile(tile(nextX(Index) + 1, nextY(Index) - 1), True, Index, "CharNEvenN+XN-Y") 'clear (nextX + 1, nextY - 1)
                End If
            End If
        End If
    End If
End If
Static counterC As Integer
If strState(Index) = "I" Then
    If strDir(Index) = "L" Then
        .PaintPicture .picCharMaskIL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            .PaintPicture .picP1IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            .PaintPicture .picP1IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "U" Then
        .PaintPicture .picCharMaskIU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            .PaintPicture .picP1IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            .PaintPicture .picP1IU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "R" Then
        .PaintPicture .picCharMaskIR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            .PaintPicture .picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            .PaintPicture .picP1IR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(Index) = "D" Then
        .PaintPicture .picCharMaskID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
        If Index = 0 Then
            .PaintPicture .picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf Index > 0 Then
            .PaintPicture .picP1ID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
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
            .PaintPicture .picCharMaskCL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If Index = 0 Then
                .PaintPicture .picP1CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index > 0 Then
                .PaintPicture .picP1CL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "U" Then
            .PaintPicture .picCharMaskCU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If Index = 0 Then
                .PaintPicture .picP1CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index > 0 Then
                .PaintPicture .picP1CU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "R" Then
            .PaintPicture .picCharMaskCR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If Index = 0 Then
                .PaintPicture .picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index > 0 Then
                .PaintPicture .picP1CR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "D" Then
            .PaintPicture .picCharMaskCD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If Index = 0 Then
                .PaintPicture .picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index > 0 Then
                .PaintPicture .picP1CD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    ElseIf strState(Index) = "J" Then
        counterC = 0
        If strDir(Index) = "L" Then
            .PaintPicture .picCharMaskJL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If Index = 0 Then
                .PaintPicture .picP1JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index > 0 Then
                .PaintPicture .picP1JL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "U" Then
            .PaintPicture .picCharMaskJU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If Index = 0 Then
                .PaintPicture .picP1JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index > 0 Then
                .PaintPicture .picP1JU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "R" Then
            .PaintPicture .picCharMaskJR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If Index = 0 Then
                .PaintPicture .picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index > 0 Then
                .PaintPicture .picP1JR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        ElseIf strDir(Index) = "D" Then
            .PaintPicture .picCharMaskJD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If Index = 0 Then
                .PaintPicture .picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf Index > 0 Then
                .PaintPicture .picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    End If
End If
End With
End Sub

Public Sub getTileAnim(ByVal intFrame As Integer, tileInput As terrain)
With frmMain
    .picBuffer.PaintPicture .picBackground.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
    .PaintPicture .picBuffer.Image, tileInput.x, (tileInput.y - 400) + (intFrame - 1) * 50, 100, 100, 0, 0, 100, 100, vbSrcCopy
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
