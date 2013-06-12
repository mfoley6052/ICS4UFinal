Attribute VB_Name = "modDataTypes"
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

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
    picObj As Object
    objFrame As Integer
    objMask As Object
    objSparkle As Boolean
    objXOffset As Single
    objYOffset As Single
    terType As String
    cutoff As String
    tempPic As Object
    picTile As Object
    picMask As Object
    pathCount As Integer
    dir As String
End Type
