Attribute VB_Name = "modInstances"
Option Explicit

Public Sub CopyMap(ByVal MapNum As Long, ByVal BaseMap As Long)
Dim x As Long, y As Long, l As Long

    Call ClearMap(MapNum)

    ' Copy the data
    With Map(MapNum)
        .BootMap = Map(BaseMap).BootMap
        .BootX = Map(BaseMap).BootX
        .BootY = Map(BaseMap).BootY
        .Down = Map(BaseMap).Down
        .Instance = Map(BaseMap).Instance
        .Left = Map(BaseMap).Left
        .Right = Map(BaseMap).Right
        .LoseItemsOnDeath = Map(BaseMap).LoseItemsOnDeath
        .MaxX = Map(BaseMap).MaxX
        .MaxY = Map(BaseMap).MaxY
        .Moral = Map(BaseMap).Moral
        .Music = Map(BaseMap).Music
        .Name = Map(BaseMap).Name
        .Right = Map(BaseMap).Right
        .Instance = Map(BaseMap).Revision + 1
        .Up = Map(BaseMap).Up
        
        For x = 1 To MAX_MAP_NPCS
            .Npc(x) = Map(BaseMap).Npc(x)
        Next
        
        ReDim Map(MapNum).Tile(0 To .MaxX, 0 To .MaxY)
        
        For x = 0 To .MaxX
            For y = 0 To .MaxY
                With .Tile(x, y)
                    .Data1 = Map(BaseMap).Tile(x, y).Data1
                    .Data2 = Map(BaseMap).Tile(x, y).Data2
                    .Data3 = Map(BaseMap).Tile(x, y).Data3
                    .DirBlock = Map(BaseMap).Tile(x, y).DirBlock
                    .Type = Map(BaseMap).Tile(x, y).Type
                    
                    For l = 1 To MapLayer.Layer_Count - 1
                        .Layer(l).Tileset = Map(BaseMap).Tile(x, y).Layer(l).Tileset
                        .Layer(l).x = Map(BaseMap).Tile(x, y).Layer(l).x
                        .Layer(l).y = Map(BaseMap).Tile(x, y).Layer(l).y
                    Next
                End With
            Next
        Next
    End With
    
    For x = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(x, MapNum)
    Next

    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)

    ' Clear out it all
    For x = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(x, 0, 0, MapNum, MapItem(MapNum, x).x, MapItem(MapNum, x).y)
        Call ClearMapItem(x, MapNum)
    Next

    ' Respawn
    Call SpawnMapItems(MapNum)
    ' Save the map
    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    Call ClearTempTile(MapNum)
    Call CacheResources(MapNum)

    ' Refresh map for everyone online
    For x = 1 To Player_HighIndex
        If IsPlaying(x) And GetPlayerMap(x) = MapNum Then
            Call PlayerWarp(x, MapNum, GetPlayerX(x), GetPlayerY(x), True)
        End If
    Next x
End Sub
