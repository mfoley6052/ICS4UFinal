Attribute VB_Name = "modCpuAi"

Public Function cpuAI(ByVal Index As Integer)
Dim openList() As terrain
Dim closedList() As terrain
Dim ext As Integer
Dim extraX As Integer
Dim extraY As Integer
'if the row is odd then there is one less tile on the x axis
If oddRow(curY(Index)) Then
    ext = -1
Else
    ext = 0
End If
'check for outer edges, and then set the array size and locations of the adjacent tiles to the current tile
If curX(Index) > 0 And curX(Index) < mapWidth + ext Then  'Not on map edge horizontally
    If curY(Index) > 0 And curY(Index) < mapHeight Then 'Not on map edge vertically
        'All adjacent tiles are available
        ReDim Preserve openList(3) As terrain
        openList(0) = tile(curX(Index) - 1, curY(Index) - 1)
        openList(1) = tile(curX(Index) - 1, curY(Index) + 1)
        openList(2) = tile(curX(Index) + 1, curY(Index) - 1)
        openList(3) = tile(curX(Index) + 1, curY(Index) + 1)
    Else 'Vertical Edge
        ReDim Preserve openList(1) As terrain
        If curY(Index) = 0 Then
            openList(1) = tile(curX(Index) - 1, curY(Index) + 1)
            openList(3) = tile(curX(Index) + 1, curY(Index) + 1)
        Else
            openList(0) = tile(curX(Index) - 1, curY(Index) - 1)
            openList(2) = tile(curX(Index) + 1, curY(Index) - 1)
        End If
    End If
Else ' Horizontal Edge
    ReDim Preserve openList(1) As terrain
    openList(2) = tile(curX(Index) + 1, curY(Index) - 1)
    openList(3) = tile(curX(Index) + 1, curY(Index) + 1)
End If
'use quadrant-style math to change look ahead into the path and calculate the steps needed
For q = LBound(openList) To UBound(openList)
    openList(q).pathCount = 0
    'extrax and extray are used to simulate the direction of the steps
    'loop until the number of steps in the direction specified from the current location = the destination
    Do Until openList(q).Xc + extraX * openList(q).pathCount = curX(0) And openList(q).Yc + extraY * openList(q).pathCount = curY(0)
        openList(q).pathCount = openList(q).pathCount + 1
        If openList(q).Xc <> curX(0) Then ' If the x coord of the tile isnt the player tile
            If openList(q).Yc <> curY(0) Then ' If the y coord of the tile isnt the player tile
                If openList(q).Xc > curX(0) And openList(q).Yc > curY(0) Then
                    extraX = -1
                    extraY = -1
                ElseIf openList(q).Xc < curX(0) And openList(q).Yc > curY(0) Then
                    extraX = 1
                    extraY = -1
                ElseIf openList(q).Xc > curX(0) And openList(q).Yc < curY(0) Then
                    extraX = -1
                    extraY = 1
                ElseIf openList(q).Xc < curX(0) And openList(q).Yc < curY(0) Then
                    extraX = 1
                    extraY = 1
               End If
            End If
        End If
    Loop
Next q
'Sort the choices by number of steps to player, holding on to the lowest amount
Dim bestTile As terrain
For q = LBound(openList) + 1 To UBound(openList)
    If openList(q).pathCount < openList(q - 1).pathCount Then
        If bestTile.pathCount > openList(q).pathCount Then
            bestTile = openList(q)
        End If
    End If
Next q
'Send choice to movement part of engine
nextX(Index) = bestTile.Xc
nextY(Index) = bestTile.Yc
End Function
