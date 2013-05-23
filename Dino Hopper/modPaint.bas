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
            If tileInput.objType(0) = "Coin" Then
                'paint terrain over coin
                intObjxOffset = 41
                .picBackground.PaintPicture tilePic.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 35, intObjxOffset, 0, 18, 35, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 18, 35, intObjxOffset, 0, 18, 35, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 35, intObjxOffset, 0, 18, 35, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 35, 0, 0, 18, 35, vbSrcPaint
            ElseIf tileInput.objType(0) = "Pow" Then
                If tileInput.objType(1) = "Scare" Then
                    'paint terrain over scare power-up
                    intObjxOffset = 35
                    .picBackground.PaintPicture tilePic.Image, tileInput.x + intObjxOffset, tileInput.y, 30, 30, intObjxOffset, 0, 30, 30, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 30, 30, intObjxOffset, 0, 30, 30, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 30, 30, intObjxOffset, 0, 30, 30, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x + intObjxOffset, tileInput.y, 30, 30, 0, 0, 30, 30, vbSrcPaint
                ElseIf tileInput.objType(1) = "Speed" Then
                    'paint terrain over speed power-up
                    intObjxOffset = 41
                    .picBackground.PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 33, intObjxOffset, 0, 18, 33, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 18, 33, intObjxOffset, 0, 18, 33, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 33, intObjxOffset, 0, 18, 33, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x + intObjxOffset, tileInput.y, 18, 33, 0, 0, 18, 33, vbSrcPaint
                ElseIf tileInput.objType(1) = "Freeze" Then
                    'paint terrain over freeze power-up
                    intObjxOffset = 32
                    .picBackground.PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 36, 37, intObjxOffset, 0, 36, 37, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 36, 37, intObjxOffset, 0, 36, 37, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 36, 37, intObjxOffset, 0, 36, 37, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x + intObjxOffset, tileInput.y, 36, 37, 0, 0, 36, 37, vbSrcPaint
                End If
            ElseIf tileInput.objType(0) = "Egg" Then
                'paint terrain over egg
                intObjxOffset = 39
                .picBackground.PaintPicture tilePic.Image, tileInput.x + intObjxOffset, tileInput.y, 22, 30, intObjxOffset, 0, 22, 30, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 22, 30, intObjxOffset, 0, 22, 30, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x + intObjxOffset, tileInput.y, 22, 30, intObjxOffset, 0, 22, 30, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x + intObjxOffset, tileInput.y, 22, 30, 0, 0, 22, 30, vbSrcPaint
            ElseIf tileInput.objType(0) = "Terrain" Then
                'paint over full tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        Else 'no object on tile
            If callID = "ObjOddX-Y" Or callID = "ObjEven+X-Y" Then
                'paint over bottom left half of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y + 50, 50, 50, 0, 50, 50, 50, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 0, 50, 50, 50, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y + 50, 50, 50, 0, 50, 50, 50, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
            ElseIf callID = "ObjOdd-X-Y" Or callID = "ObjEvenX-Y" Then
                'paint over bottom right half of tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 50, 50, 50, 50, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
            End If
        End If
    Else 'called by character
        If tileInput.hasChar Then 'if character is on tile
            If (tileInput.Xc = curX(index) And tileInput.Yc = curY(index)) Then 'if character on tile is character that called clear
                If tileInput.Yc = 0 Then '(x, 0)
                    Call clearVoid(tileInput, True, True) 'paint spacer
                ElseIf oddRow(tileInput.Yc) Then '(x, odd)
                    If tileInput.Xc = 0 Then 'if first column
                        Call clearVoid(tileInput, True, False) 'paint right spacer
                    ElseIf tileInput.Xc = mapWidth Then 'if last column
                        Call clearVoid(tileInput, False, True) 'paint left spacer
                    End If
                End If
                If Not tileTouchingChar(index, tileInput) Then 'if not touching a char, paint full tile
                    'paint over tile
                    .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
                ElseIf tileInput.Yc = mapHeight Then 'char jumping off bottom
                    If callID = "CharBottom+X+Y" Then
                        'paint over bottom left half of tile
                        .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y + 50, 50, 50, 0, 50, 50, 50, vbSrcCopy
                        .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 0, 50, 50, 50, vbSrcCopy
                        .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y + 50, 50, 50, 0, 50, 50, 50, vbSrcAnd
                        .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
                    ElseIf callID = "CharBottom-X+Y" Then
                        'paint over bottom right half of tile
                        .picBackground.PaintPicture tilePic.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcCopy
                        .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 50, 50, 50, 50, vbSrcCopy
                        .PaintPicture tileInput.picMask.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcAnd
                        .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
                    End If
                ElseIf callID = "CharOdd+X-2Y" Or callID = "CharOdd-X-2Y" Then 'if jumping off side edge
                    If tileInput.Xc = 0 Then
                        'paint over bottom left half of tile
                        .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y + 50, 50, 50, 0, 50, 50, 50, vbSrcCopy
                        .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 0, 50, 50, 50, vbSrcCopy
                        .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y + 50, 50, 50, 0, 50, 50, 50, vbSrcAnd
                        .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
                    ElseIf tileInput.Xc = mapWidth Then
                        'paint over bottom right half of tile
                        .picBackground.PaintPicture tilePic.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcCopy
                        .picBuffer.PaintPicture tilePic.Image, 0, 0, 50, 50, 50, 50, 50, 50, vbSrcCopy
                        .PaintPicture tileInput.picMask.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 50, 50, 50, 50, vbSrcAnd
                        .PaintPicture .picBuffer.Image, tileInput.x + 50, tileInput.y + 50, 50, 50, 0, 0, 50, 50, vbSrcPaint
                    End If
                Else 'if touching a char
                    'only paint over top half of tile
                    .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                    .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                    .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcPaint
                End If
            End If
        Else 'if not character on tile
            'if not touching character or coordinates match previous tile of character
            If Not tileTouchingChar(index, tileInput) Or (tileInput.Xc = prevX(index) And tileInput.Yc = prevY(index)) Then
                'paint over tile
                .picBackground.PaintPicture tilePic, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            Else 'if touching a char
                'only paint over top half of tile
                '.picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcCopy
                '.picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 50, 0, 0, 100, 50, vbSrcCopy
                '.PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcAnd
                '.PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 50, 0, 0, 100, 50, vbSrcPaint
                'paint over tile
                .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .picBuffer.PaintPicture tilePic.Image, 0, 0, 100, 100, 0, 0, 100, 100, vbSrcCopy
                .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcAnd
                .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
            End If
        End If
    End If
ElseIf bypassForObj And tileInput.hasObj Then
    'paint over bottom half of tile
    .picBackground.PaintPicture tilePic.Image, tileInput.x, tileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
    .picBuffer.PaintPicture tilePic.Image, 0, 50, 100, 50, 0, 50, 100, 50, vbSrcCopy
    .PaintPicture tileInput.picMask.Image, tileInput.x, tileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcAnd
    .PaintPicture .picBuffer.Image, tileInput.x, tileInput.y + 50, 100, 50, 0, 50, 100, 50, vbSrcPaint
End If
End With
End Sub

Public Sub clearVoid(tileInput As terrain, ByVal blnL As Boolean, ByVal blnR As Boolean) 'clear empty spots on map
With frmMain
    If blnL And blnR Then
        '.picBackground.PaintPicture .picSpacer.Image, tileInput.x, tileInput.y - 50, 100, 98, 0, 0, 100, 98, vbSrcCopy
        '.PaintPicture .picSpacerMask.Image, tileInput.x, tileInput.y - 50, 100, 98, 0, 0, 100, 98, vbSrcAnd
        '.PaintPicture .picBackground.Image, tileInput.x, tileInput.y - 50, 100, 98, 0, 0, 100, 98, vbSrcCopy
        .picBackground.PaintPicture .picSpacer.Image, tileInput.x, tileInput.y - 50, 50, 98, 0, 0, 50, 98, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, tileInput.x, tileInput.y - 50, 50, 98, 0, 0, 50, 98, vbSrcAnd
        .PaintPicture .picBackground.Image, tileInput.x, tileInput.y - 50, 50, 98, 0, 0, 50, 98, vbSrcCopy
    ElseIf blnL Then
        .picBackground.PaintPicture .picSpacer.Image, tileInput.x, tileInput.y - 50, 50, 98, 0, 0, 50, 98, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, tileInput.x, tileInput.y - 50, 50, 98, 0, 0, 50, 98, vbSrcAnd
        .PaintPicture .picBackground.Image, tileInput.x, tileInput.y - 50, 50, 98, 0, 0, 50, 98, vbSrcCopy
    ElseIf blnR Then
        .picBackground.PaintPicture .picSpacer.Image, tileInput.x + 50, tileInput.y - 50, 50, 98, 0, 0, 50, 98, vbSrcCopy
        .PaintPicture .picSpacerMask.Image, tileInput.x + 50, tileInput.y - 50, 50, 98, 0, 0, 50, 98, vbSrcAnd
        .PaintPicture .picBackground.Image, tileInput.x + 50, tileInput.y - 50, 50, 98, tileInput.x + 50, 0, 50, 98, vbSrcCopy
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
End If
If Not killObj Then 'if object has not expired
    If strObjType = "Coin" Then
        'paint coin
        With frmMain
        If paintMask(tile(intObjX, intObjY)) Then
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
            If paintMask(tile(intObjX, intObjY)) Then
                .PaintPicture .picSparkleMask(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picSparkle(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
        End With
    ElseIf strObjType = "Pow" Then
        'paint power-up
        With frmMain
        If strType = "Scare" Then
            If paintMask(tile(intObjX, intObjY)) Then
                .PaintPicture .picPowScareMask(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picPowScare(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "Speed" Then
            If paintMask(tile(intObjX, intObjY)) Then
                .PaintPicture .picPowSpeedMask(intFrame).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picPowSpeed(intFrame).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "Freeze" Then
            If paintMask(tile(intObjX, intObjY)) Then
                .PaintPicture .picPowFreezeMask.Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picPowFreeze(intFrame + intFrameOffset).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
        End With
    ElseIf strObjType = "Egg" Then
        'paint egg
        With frmMain
        If strType = "M" Then
            If paintMask(tile(intObjX, intObjY)) Then
                .PaintPicture .picEggMask((intFrame \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picEgg((intFrame \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf strType = "G" Then
            If paintMask(tile(intObjX, intObjY)) Then
                .PaintPicture .picEggMask(((intFrame - (8 * Int(intFrame / 8))) \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcAnd
            End If
            .PaintPicture .picEggG(((intFrame - (8 * Int(intFrame / 8))) \ 2) + (intFrameOffset \ 2)).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + intYOffset, 100, 100, 0, 0, 100, 100, vbSrcPaint
            'paint sparkle
            If intFrame > 11 And intFrame < 19 Then
                If paintMask(tile(intObjX, intObjY)) Then
                    .PaintPicture .picSparkleMask(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcAnd
                End If
                .PaintPicture .picSparkle(intFrame - 12).Image, tile(intObjX, intObjY).x + intXOffset, tile(intObjX, intObjY).y + (intYOffset + 2), 100, 100, 0, 0, 100, 100, vbSrcPaint
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
    .PaintPicture .picSelMask(imgIndex + 5 * index).Image, tile(curX(index), curY(index)).x, tile(curX(index), curY(index)).y, 100, 100, 0, 0, 100, 100, vbSrcAnd
    .PaintPicture .picSel(imgIndex + 5 * index).Image, tile(curX(index), curY(index)).x, tile(curX(index), curY(index)).y, 100, 100, 0, 0, 100, 100, vbSrcPaint
End With
Call PaintCharSprite(index, spriteX(index), spriteY(index))
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
    If strDir(index) = "D" Then
        If oddRow(nextY(index)) Then
        Else
        End If
    ElseIf nextY(index) > 0 Then
        If strDir(index) = "U" Then
            If oddRow(nextY(index)) Then
                If nextX(index) <= mapWidth - 1 Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index, "CharNOddNXN-Y") 'clear (nextX, nextY - 1)
                End If
                If nextX(index) < mapWidth - 1 Then
                    Call clearTile(tile(nextX(index) + 1, nextY(index) - 1), True, index, "CharNOddN+XN-Y") 'clear (nextX + 1, nextY - 1)
                End If
            Else
                If nextX(index) <= mapWidth Then
                    Call clearTile(tile(nextX(index), nextY(index) - 1), True, index, "CharNEvenNXN-Y") 'clear (nextX, nextY - 1)
                End If
                If nextX(index) > 0 Then
                    Call clearTile(tile(nextX(index) - 1, nextY(index) - 1), True, index, "CharNEvenN-XN-Y") 'clear (nextX - 1, nextY - 1)
                End If
            End If
        ElseIf strDir(index) = "R" Then
            If oddRow(nextY(index) - 1) Then
                If nextX(index) > 0 Then
                    Call clearTile(tile(nextX(index) - 1, nextY(index) - 1), True, index, "CharNOddN-XN-Y") 'clear (nextX - 1, nextY - 1)
                End If
            Else
                If nextX(index) < mapWidth - 1 Then
                    Call clearTile(tile(nextX(index) + 1, nextY(index) - 1), True, index, "CharNEvenN+XN-Y") 'clear (nextX + 1, nextY - 1)
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
        ElseIf index = 1 Then
            .PaintPicture .picP2IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 2 Then
            .PaintPicture .picP3IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        ElseIf index = 3 Then
            .PaintPicture .picP4IL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf strDir(index) = "U" Then
        .PaintPicture .picCharMaskIU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
        .PaintPicture .picCharMaskIR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
        .PaintPicture .picCharMaskID.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
            .PaintPicture .picCharMaskCL(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
            .PaintPicture .picCharMaskCU(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
            .PaintPicture .picCharMaskCR(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
            .PaintPicture .picCharMaskCD(frameC).Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
            .PaintPicture .picCharMaskJL.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
            .PaintPicture .picCharMaskJU.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
            .PaintPicture .picCharMaskJR.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
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
            .PaintPicture .picCharMaskJD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcAnd
            If index = 0 Then
                .PaintPicture .picP1JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 1 Then
                .PaintPicture .picP2JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 2 Then
                .PaintPicture .picP2JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
            ElseIf index = 3 Then
                .PaintPicture .picP2JD.Image, charX, charY, 100, 100, 0, 0, 100, 100, vbSrcPaint
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
