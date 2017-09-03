Attribute VB_Name = "modPlayerCustomValues"
Option Explicit

Public PlayerValues(1 To MAX_PLAYERS) As PlayerValueRec

Private Type PlayerValueRec
    PestControlPoints As Long
    DidFightCaves As Byte
    DidKiln As Byte
    DungeoneeringTokens As Long
End Type

Public Sub LoadPlayerValue(ByVal Index As Long)
    
    Dim file As String
    file = App.Path & "\data\accounts\" & Player(Index).Name & ".ini"
    
    With PlayerValues(Index)
        If Not FileExist(file, True) Then
            .PestControlPoints = 0
            .DidFightCaves = 0
            .DidKiln = 0
            .DungeoneeringTokens = 0
        Else
            .PestControlPoints = GetVar(file, "Minigames", "PCPoints")
            .DidKiln = GetVar(file, "Minigames", "Kiln")
            .DidFightCaves = GetVar(file, "Minigames", "FightCaves")
            .DungeoneeringTokens = GetVar(file, "Skills", "DungTokens")
        End If
    End With

End Sub

Public Sub SavePlayerValue(ByVal Index As Long)

    Dim file As String
    file = App.Path & "\data\accounts\" & Player(Index).Name & ".ini"
    
    With PlayerValues(Index)
        Call PutVar(file, "Minigames", "PCPoints", STR(.PestControlPoints))
        Call PutVar(file, "Minigames", "Kiln", STR(.DidKiln))
        Call PutVar(file, "Minigames", "FightCaves", STR(.DidFightCaves))
        Call PutVar(file, "Skills", "DungTokens", STR(.DungeoneeringTokens))
    End With

End Sub

