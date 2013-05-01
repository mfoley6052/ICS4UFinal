Attribute VB_Name = "modDataTypes"
 Public Type npc 'Data type for all characters in game
    health As Integer
    speed As Integer
    strength As Integer
    movement As Integer
    level As Integer
    X As Single
    Y As Single
    alive As Boolean
    turn As Boolean
End Type

Public Type terrain ' data type for different terrains
    pic As String
    X As Single
    Y As Single
    selectable As Boolean
    oldX As Single
    oldY As Single
    hasChar As Boolean
    coinEnabled As Boolean
    coinType As String
    coinTimer As Long
End Type



