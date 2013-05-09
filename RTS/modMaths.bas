Attribute VB_Name = "modMaths"
Option Explicit

Public Sub addScore(ByVal intAdd As Integer)
intScore = intScore + (intMulti * intAdd)
End Sub

Public Sub refreshLabels(ByVal blnScore As Boolean, ByVal blnLives As Boolean, ByVal blnMulti As Boolean)
With frmMain
If blnScore Then
    .lblScore = "Score " & Format(intScore, "0000000")
End If
If blnLives Then
End If
If blnMulti Then
    .lblMulti = "Multiplier: " & intMulti & "x"
End If
End With
End Sub

Public Function tileTouchingChar(TileInput As terrain) As Boolean
tileTouchingChar = False
If TileInput.Yc < mapHeight - 1 Then 'if above bottom row of tiles
    If oddRow(TileInput.Yc + 1) Then 'if tile row is odd
        If tile(TileInput.Xc, TileInput.Yc + 1).hasChar Then 'if other below tile is taken by a character
            tileTouchingChar = True
        Else 'if tile (TileInput.xc, TileInput.yc + 1) is not taken by a character
            If TileInput.Xc > 0 Then 'if column is less than last column
                If tile(TileInput.Xc - 1, TileInput.Yc + 1).hasChar Then 'if tile (TileInput.xc - 1, TileInput.yc + 1) is taken by a character
                    tileTouchingChar = True
                End If
            End If
        End If
    Else 'if tile row is even
        If tile(TileInput.Xc, TileInput.Yc + 1).hasChar Then 'if tile (TileInput.xc, TileInput.yc + 1) is taken by a character
            tileTouchingChar = True
        ElseIf TileInput.Xc < mapWidth - 1 Then 'if column is greater than first
            If tile(TileInput.Xc + 1, TileInput.Yc + 1).hasChar Then 'if other below tile is taken by a character
                tileTouchingChar = True
            End If
        End If
    End If
End If
End Function

Public Sub killObj(inputTile As terrain)
Call clearTile(inputTile, False, -1)
tile(inputTile.Xc, inputTile.Yc).objTimer = 0
tile(inputTile.Xc, inputTile.Yc).hasObj = False
objTileCount = objTileCount - 1
End Sub

'glitch: frames are often skipped (eg. 5, 6, 8, 9, 10) even though tmrFrame and tmrChar are same interval
Public Function getCharJumpAnim(ByVal Index As Integer, curTile As terrain, nextTile As terrain)
'if frame 5 to 10
If frameCounter(Index) >= 5 And frameCounter(Index) <= 10 Then
    spriteX(Index) = curTile.x + ((frameCounter(Index) - 5) * Int((nextTile.x - curTile.x) / 5)) + 25
    '5 to 7 is jump up
    If frameCounter(Index) < 8 Then
        spriteY(Index) = (curTile.y + ((frameCounter(Index) - 5) * Int((nextTile.y - curTile.y) / 5))) - (10 * (frameCounter(Index) - 5)) - 15
    '8 to 10 is fall to ground
    ElseIf frameCounter(Index) <= 10 Then
        spriteY(Index) = (curTile.y + ((frameCounter(Index) - 5) * Int((nextTile.y - curTile.y) / 5))) - (10 * (10 - frameCounter(Index))) - 15
    End If
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
            a = 2
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

Public Function getAbs(ByVal valInput As Single) As Integer
If valInput < 0 Then
    valInput = 0
End If
getAbs = valInput
End Function

Public Function randInt(ByVal min As Integer, ByVal max As Integer) As Integer
randInt = Int(Rnd() * max) + min
End Function

Public Function oddRow(ByVal intY As Integer) As Boolean
If (intY + 1) Mod 2 = 0 Then
    oddRow = True
ElseIf (intY + 1) Mod 2 = 1 Then
    oddRow = False
End If
End Function

