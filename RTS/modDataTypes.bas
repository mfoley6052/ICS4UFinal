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
    selectable As Boolean
    oldX As Single
    oldY As Single
    hasChar As Boolean
    coinEnabled As Boolean
    coinType As String
    coinTimer As Long
End Type

Public Type BreadCrumb 'path for AI
    x As Integer
    y As Integer
    tile As terrain
End Type


