Attribute VB_Name = "modDataTypes"
 Public Type npc 'Data type for all characters in game
    health As Integer
    speed As Integer
    strength As Integer
    movement As Integer
    level As Integer
    x As Single
    y As Single
    alive As Boolean
    turn As Boolean
End Type

Public Type terrain ' data type for different terrains
    pic As String
    x As Single
    y As Single
    Xc As Integer
    Yc As Integer
    selectable As Boolean
    oldX As Single
    oldY As Single
    hasChar As Boolean
    hasObj As Boolean
    objType(1) As String
    objTimer As Long
    picTile As Object
    picMask As Object
End Type

Public Type BreadCrumb 'path for AI
    x As Integer
    y As Integer
    tile As terrain
End Type

