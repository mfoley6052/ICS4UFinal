Attribute VB_Name = "modMaths"
Public Sub addScore(ByVal Index As Integer, ByVal intAdd As Long)
intScore(Index) = intScore(Index) + (intMulti(Index) * intAdd)
End Sub

Public Sub refreshLabels(ByVal Index As Integer, ByVal blnScore As Boolean, ByVal blnLives As Boolean, ByVal blnMulti As Boolean)
With frmGUI
If blnScore Then
    .lblScore(Index) = "Score: " & Format(intScore(Index), "0000000")
End If
If blnLives Then
    .lblLives(Index) = "x" & intLives(Index)
End If
If blnMulti Then
    .lblMulti(Index) = "Multiplier: " & intMulti(Index) & "x"
End If
End With
End Sub

Public Function charTouchingTile(ByVal Index As Integer, tileInput As terrain) As Boolean
charTouchingTile = False
Dim charCheck As Integer
For charCheck = 0 To 3 Step 1
    With frmMain
    If charCheck <> Index And .tmrChar(charCheck).Enabled Then
        If curX(charCheck) = tileInput.Xc And curY(charCheck) = tileInput.Yc Then
            charTouchingTile = True
        ElseIf oddRow(curY(charCheck)) Then
            If curY(charCheck) = tileInput.Yc + 1 Then
                If curX(charCheck) = tileInput.Xc Or curX(charCheck) - 1 = tileInput.Xc Then
                    charTouchingTile = True
                End If
            End If
        Else
            If curY(charCheck) = tileInput.Yc + 1 Then
                If curX(charCheck) = tileInput.Xc Or curX(charCheck) + 1 = tileInput.Xc Then
                    charTouchingTile = True
                End If
            End If
        End If
    End If
    End With
Next charCheck
End Function

Public Function checkClearVoid(tileInput As terrain, ByVal blnL As Boolean, ByVal blnR As Boolean) As Boolean
checkClearVoid = False
If blnL Then
    If tileInput.Yc = 0 Then
        checkClearVoid = True
    ElseIf tileInput.Xc = 0 And oddRow(tileInput.Yc) Then
        checkClearVoid = True
    End If
ElseIf blnR Then
    If tileInput.Yc = 0 Then
        checkClearVoid = True
    ElseIf tileInput.Xc = mapWidth Then
        checkClearVoid = True
    End If
End If
End Function

Public Sub killObj(inputTile As terrain)
inputTile.objTimer = 0
inputTile.objFrame = 0
If inputTile.objType(0) <> "Terrain" Then
    Call PaintObj(inputTile.objType(0), inputTile.objType(1), 0, inputTile.Xc, inputTile.Yc, True)
Else
    Call getChangeTerrain(inputTile, Mid(inputTile.terType, 2, 1), True)
    Call clearTile(inputTile, False, -1, "ObjXYF")
End If
inputTile.hasObj = False
inputTile.objType(0) = ""
inputTile.objType(1) = ""
tile(inputTile.Xc, inputTile.Yc) = inputTile
objTileCount = objTileCount - 1
End Sub

Public Function paintMask(tileInput As terrain, ByVal Index As Integer) As Boolean
If Index < 0 Then
    If gameMode = 0 Then
        Dim ratExpire As Single
        If tileInput.objTimer < 0.7 * objExpire Then
            paintMask = True
        Else
            ratExpire = tileInput.objTimer / objExpire
            If (ratExpire >= 0.75 And ratExpire < 0.8) Or (ratExpire >= 0.85 And ratExpire < 0.88) Or (ratExpire >= 0.91 And ratExpire < 0.94) Or (ratExpire >= 0.96 And ratExpire < 0.97) Or (ratExpire >= 0.98 And ratExpire < 0.99) Then
                paintMask = True
            Else
                paintMask = False
            End If
        End If
    Else
        If tileInput.objTimer = objExpire - 1 Then
            With frmMain
            paintMask = .tmrAlternate.Tag
            End With
        Else
            paintMask = True
        End If
    End If
ElseIf blnRecover(Index) Then
    With frmMain
        paintMask = .tmrAlternate.Tag
    End With
End If
End Function

Public Function isPlayer(ByVal Index As Integer) As Boolean
If Index <= numPlayers - 1 Then
    isPlayer = True
Else
    isPlayer = False
End If
End Function

Public Function getCharFromTile(tileInput As terrain) As Integer
For x = 0 To 3
    If frmMain.tmrChar(x).Enabled Then
        If curX(x) = tileInput.Xc And curY(x) = tileInput.Yc Then
            getCharFromTile = Index
            Exit Function
        End If
    End If
Next x
End Function

Public Function getCharJumpAnim(ByVal Index As Integer, ByVal curFrame As Integer, curTile As terrain, ByVal nextX As Integer, ByVal nextY As Integer)
'if frame 5 to 10
If curFrame >= 5 And curFrame <= 10 Then
    spriteX(Index) = curTile.x + ((curFrame - 5) * Int((nextX - curTile.x) / 5)) + 25
    '5 to 7 is jump up
    If curFrame < 8 Then
        spriteY(Index) = (curTile.y + ((curFrame - 5) * Int((nextY - curTile.y) / 5))) - (10 * (curFrame - 5)) - 15
    '8 to 10 is fall to ground
    ElseIf curFrame <= 10 Then
        spriteY(Index) = (curTile.y + ((curFrame - 5) * Int((nextY - curTile.y) / 5))) - (10 * (10 - curFrame)) - 15
    End If
    With frmMain
    If curFrame = 5 Then
        If .tmrChar(Index).Tag = "Freeze" Then 'if freeze, change terrain to ice and paint half-transparency ice tile
            Call getChangeTerrain(curTile, "I", False)
            .PaintPicture curTile.picTile.Image, curTile.x, curTile.y, 100, 100, 0, 0, 100, 100, vbSrcPaint
        End If
    ElseIf curFrame = 8 Then
        If .tmrChar(Index).Tag = "Freeze" Then 'if freeze power-up
            'if terrain has no object (if it has an object yet the character is on it, the object is a terrain)
            If Not tile(curX(Index), curY(Index)).hasObj Then
                Call clearTile(curTile, False, -1, "ObjXYF") 'paint full-transparency ice tile
            End If
        End If
    End If
    End With
ElseIf curFrame > 10 And curFrame <= 16 Then
    spriteX(Index) = curTile.x + ((curFrame - 5) * Int((nextX - curTile.x) / 5)) + 25
    spriteY(Index) = (nextY - 10) + ((curFrame - 10) * (curFrame * 2))
End If
End Function


Public Function getTileFromInt(ByVal blnX As Boolean, ByVal intInput As Integer) As Integer
If blnX = True Then
    If intInput >= mapWidth Then
        If intInput >= (mapWidth * 2) + 1 Then
            Dim a As Integer
            a = 2
            Do Until intInput < (mapWidth * a) + Int(a / 2)
                a = a + 1
            Loop
            getTileFromInt = intInput - ((mapWidth * (a - 1)) + Int((a - 1) / 2))
        Else
            getTileFromInt = intInput - mapWidth
        End If
    Else
        getTileFromInt = intInput
    End If
ElseIf blnX = False Then
    If intInput >= mapWidth Then
        If intInput >= (mapWidth * 2) + 1 Then
            Dim b As Integer
            'a = 2
            Do Until intInput < (mapWidth * b) + Int(b / 2)
                b = b + 1
            Loop
            getTileFromInt = b - 1
        Else
            getTileFromInt = 1
        End If
    Else
        getTileFromInt = 0
    End If
End If
End Function

Public Sub setCharRespawn(ByVal Index As Integer, tileInput As terrain)
Dim newX As Integer
Dim newY As Integer
If curY(Index) = mapHeight - 1 Then
    newX = curX(Index)
    newY = 0
ElseIf curY(Index) = 0 Then
    newX = curX(Index)
    newY = mapHeight - 1
ElseIf curX(Index) = 0 Then
    If oddRow(curY(Index)) Then
        newX = mapWidth
    Else
        newX = mapWidth - 1
    End If
    newY = curY(Index)
Else
    newX = 0
    newY = curY(Index)
End If
If tile(newX, newY).hasChar Then
    newY = 3
    newX = 3
    If tile(3, newY).hasChar Then
        newX = 4
    End If
End If
nextX(Index) = newX
nextY(Index) = newY
End Sub

Public Function getAbs(ByVal valInput As Single) As Integer
If valInput < 0 Then
    valInput = 0
End If
getAbs = valInput
End Function

Public Function randInt(ByVal min As Integer, ByVal max As Integer) As Integer
randInt = Int(Rnd() * (max + 1)) + min
End Function

Public Function oddRow(ByVal intY As Integer) As Boolean
If (intY + 1) Mod 2 = 0 Then
    oddRow = True
ElseIf (intY + 1) Mod 2 = 1 Then
    oddRow = False
End If
End Function

