Attribute VB_Name = "modVars"
Public selType(0 To 3) As String
Public picCount(0 To 3) As Integer
Public strDir(0 To 3) As String
Public blnPlayerMoveable As Boolean
Public counterLimit(1 To 3) As Integer
Public frameCounter(0 To 3) As Integer
Public frameLimit(0 To 3) As Integer
Public strState(0 To 3) As String
Public intScore As Integer
Public spriteX(0 To 3) As Integer
Public spriteY(0 To 3) As Integer
Public curX(0 To 3) As Integer
Public curY(0 To 3) As Integer
Public prevX(0 To 3) As Integer
Public prevY(0 To 3) As Integer
Public nextX(0 To 3) As Integer
Public nextY(0 To 3) As Integer
Public coinTileCount As Integer
Public blnClearPrevTile(0 To 3) As Boolean
Public tileSwitch(0 To 100) As Boolean
Public limswitch As Long
Public smallestX As Integer
Public smallestY As Integer
Public pathStep() As BreadCrumb
