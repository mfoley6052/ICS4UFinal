Attribute VB_Name = "modVars"
Public gameMode As Integer
Public blnDebug As Boolean
Public blnGame As Boolean
Public selType(0 To 3) As String
Public picCount(0 To 3) As Integer
Public strDir(0 To 3) As String
Public blnPlayerMoveable(0 To 3) As Boolean
Public blnEdgeJump(0 To 3) As Boolean
Public blnBounceJump(0 To 3) As Boolean
Public blnMoveOnTick(0 To 3) As Boolean
Public counterLimit(1 To 3) As Integer
Public frameCounter(0 To 3) As Integer
Public frameProg(0 To 3) As Integer
Public frameLimit(0 To 3) As Integer
Public intMoveCount As Integer
Public strState(0 To 3) As String
Public intScore(0 To 3) As Long
Public intLives(0 To 3) As Integer
Public intMulti(0 To 3) As Long
Public spriteX(0 To 3) As Integer
Public spriteY(0 To 3) As Integer
Public curX(0 To 3) As Integer
Public curY(0 To 3) As Integer
Public prevX(0 To 3) As Integer
Public prevY(0 To 3) As Integer
Public nextX(0 To 3) As Integer
Public nextY(0 To 3) As Integer
Public targIndex(0 To 3) As Integer
Public defaultTile(0 To 3) As terrain
Public intMoves(0 To 3) As Integer
Public blnRecover(0 To 3) As Boolean
Public objExpire As Integer
Public objTileCount As Integer
Public tmrPowCounter(0 To 3) As Integer
Public tmrPowLimit As Integer
Public blnClearPrevTile(0 To 3) As Boolean
Public tileSwitch(0 To 100) As Boolean
Public limswitch As Long
Public smallestX As Integer
Public smallestY As Integer
Public intObjxOffset As Integer
Public key() As Integer
Public DefaultKey() As Integer
Public cpuInterval As Integer
Public picBG As Object
Public numPlayers As Integer
Public numCPU As Integer
Public isScared As Boolean
Public gameStarted As Boolean
Public attackMode(1 To 3) As Integer
Public CPUWait(1 To 3) As Integer

