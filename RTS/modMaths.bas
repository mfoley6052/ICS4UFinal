Attribute VB_Name = "modMaths"
Option Explicit

Public Sub addScore(ByVal intAdd As Integer)
intScore = intScore + intAdd
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

Public Function getCharJumpAnim(ByVal index As Integer, ByVal IntNewX As Integer, ByVal intNewY As Integer, ByVal intOldX As Integer, ByVal intOldY As Integer)
'if frame 5 to 10
If frameCounter(index) >= 5 And frameCounter(index) <= 10 Then
    spriteX(index) = intOldX + ((frameCounter(index) - 5) * Int((IntNewX - intOldX) / 5))
    '5 to 7 is jump up
    If frameCounter(index) < 8 Then
        spriteY(index) = (intOldY + ((frameCounter(index) - 5) * Int((intNewY - intOldY) / 5))) - (4 * (frameCounter(index) - 5))
    '8 to 10 is fall to ground
    ElseIf frameCounter(index) <= 10 Then
        spriteY(index) = (intOldY + ((frameCounter(index) - 5) * Int((intNewY - intOldY) / 5))) - (4 * (10 - frameCounter(index)))
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

