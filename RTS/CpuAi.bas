Attribute VB_Name = "modCpuAi"
Public Type BreadCrumb
    x As Integer
    y As Integer
    tile As terrain
End Type
Public Function cpuAI(ByVal index As Integer)
Dim tempRand As Integer
Dim hasBC As Boolean
Static pathStep() As BreadCrumb
If index = 0 Then
'Lead the way
    For q = LBound(pathStep) To UBound(pathStep)
        if pathStep(q).x = 'Current position of ai
    Next q
    Randomize Timer
    tempRand = Int(rand() * 3) + 1
    
Else
'Follow Leader with breadcrumbs
End If
End Function

Public Function nearestCoin(ByVal index As Integer) As Long
smallestX = 0
smallestY = 0
For x = 0 To mapWidth
    For y = 0 To mapHeight
    'if the tile has a coin then check if it is closer than the closest one, if it is make it the closest one
        If tile(x, y).coinEnabled = True Then
            If Abs(curX(index) - tile(x, y).x) < smallestX And Abs(curY(index) - tile(x, y).y) < smallestY Then
                smallestX = Abs(curX(index) - tile(x, y).x)
                smallestY = Abs(curY(index) - tile(x, y).y)
            End If
        End If
    Next y
Next x
End Function
