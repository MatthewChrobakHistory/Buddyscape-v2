Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, MapNum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim packet As String
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
            With MapItem(MapNum, i)
                .Owner = playerName
                .tmr = ITEM_SPAWN_TIME
                .state = ITEM_DESPAWN_PLAYER
                MapItem(MapNum, i).Num = ItemNum
                MapItem(MapNum, i).value = ItemVal
                MapItem(MapNum, i).x = x
                MapItem(MapNum, i).y = y
            End With
            ' send to map
            SendSpawnItemToMap MapNum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If item(Map(MapNum).Tile(x, y).Data1).Stackable And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Dim NpcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    NpcNum = Map(MapNum).Npc(MapNpcNum)

    If NpcNum > 0 Then
    
        MapNpc(MapNum).Npc(MapNpcNum).Num = NpcNum
        MapNpc(MapNum).Npc(MapNpcNum).Target = 0
        MapNpc(MapNum).Npc(MapNpcNum).TargetType = 0 ' clear
        
        For i = 1 To MAX_PLAYERS
            MapNpc(MapNum).Npc(MapNpcNum).PlayerDamage(i) = 0
        Next
        
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HitPoints) = GetNpcMaxVital(NpcNum, Vitals.HitPoints)
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.Prayer) = GetNpcMaxVital(NpcNum, Vitals.Prayer)
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.Summoning) = GetNpcMaxVital(NpcNum, Vitals.Summoning)
        
        MapNpc(MapNum).Npc(MapNpcNum).dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(MapNum).MaxX
            For y = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(MapNum).Tile(x, y).Data1 = MapNpcNum Then
                        MapNpc(MapNum).Npc(MapNpcNum).x = x
                        MapNpc(MapNum).Npc(MapNpcNum).y = y
                        MapNpc(MapNum).Npc(MapNpcNum).dir = Map(MapNum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, Map(MapNum).MaxX)
                y = Random(0, Map(MapNum).MaxY)
    
                If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
                If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, x, y) Then
                    MapNpc(MapNum).Npc(MapNpcNum).x = x
                    MapNpc(MapNum).Npc(MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(MapNum).MaxX
                For y = 0 To Map(MapNum).MaxY

                    If NpcTileIsOpen(MapNum, x, y) Then
                        MapNpc(MapNum).Npc(MapNpcNum).x = x
                        MapNpc(MapNum).Npc(MapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong SSpawnNpc
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Num
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).dir
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
        
        SendMapNpcVitals MapNum, MapNpcNum
    End If

End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(MapNum).Npc(LoopI).x = x Then
                If MapNpc(MapNum).Npc(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next

End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(MapNum).Npc(MapNpcNum).x
    y = MapNpc(MapNum).Npc(MapNpcNum).y
    CanNpcMove = True

    Select Case dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(MapNum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                If n = TILE_TYPE_OBJECT Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                If n = TILE_TYPE_OBJECT Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(MapNum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                If n = TILE_TYPE_OBJECT Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If
                
                If n = TILE_TYPE_OBJECT Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(MapNpcNum).dir = dir

    Select Case dir
        Case DIR_UP
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_DOWN
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_LEFT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        Case DIR_RIGHT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
            Set Buffer = New clsBuffer
            Buffer.WriteLong SNpcMove
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).dir
            Buffer.WriteLong movement
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal dir As Long)
    Dim packet As String
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(MapNum).Npc(MapNpcNum).dir = dir
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcDir
    Buffer.WriteLong MapNpcNum
    Buffer.WriteLong dir
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Sub ClearTempTiles()
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal MapNum As Long)
    Dim y As Long
    Dim x As Long
    TempTile(MapNum).DoorTimer = 0
    ReDim TempTile(MapNum).DoorOpen(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            TempTile(MapNum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Function BuyItem(ByVal Index As Long, ByVal ShopNum As Long, ByVal ShopSlot As Long, ByVal Amount As Long) As Boolean
    Dim ItemAmount(1 To MAX_SHOP_ITEM_COSTS) As Long
    Dim i As Long
    
    With Shop(ShopNum).TradeItem(ShopSlot)
        ' check has the cost item
        For i = 1 To MAX_SHOP_ITEM_COSTS
            ' Check if the cost item is bigger than 0
            If .CostItem(i) > 0 Then
                ItemAmount(i) = HasItem(Index, .CostItem(i))
                If ItemAmount(i) = 0 Or ItemAmount(i) < .CostValue(i) Then
                    PlayerMsg Index, "You do not have enough to buy this item.", BrightRed
                    BuyItem = False
                    Exit Function
                End If
            End If
        Next
        
        For i = 1 To Skills.Skill_Count - 1
            If Player(Index).Skill(i).Level < item(.item).SkillMakeReq(i) Then
                PlayerMsg Index, "You need to have a " & GetSkillName(i) & " level of " & item(.item).SkillMakeReq(i) & " to get this item.", BrightRed
            End If
        Next
        
        If FindOpenInvSlot(Index, .item) = 0 Then
            PlayerMsg Index, "Your inventory is full.", BrightRed
            BuyItem = False
            Exit Function
        End If
        
        If CustomShop_Buy(Index, ShopNum, ShopSlot, Amount) Then
            BuyItem = True
            Exit Function
        End If
        
        ' it's fine, let's go ahead
        ' Take all the items!
        For i = 1 To MAX_SHOP_ITEM_COSTS
            ' Check if the cost item is bigger than 0
            If .CostItem(i) > 0 Then
                TakeInvItem Index, .CostItem(i), .CostValue(i)
            End If
        Next
        
        ' Give the item
        GiveInvItem Index, .item, Amount
        
        For i = 1 To Skills.Skill_Count - 1
            Call GivePlayerSkillXP(Index, i, item(.item).SkillMakeRew(i))
        Next
    End With
    
    BuyItem = True
End Function

Public Sub SellItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal Amount As Long)
Dim i As Long
Dim multiplier As Double
    
    multiplier = Shop(TempPlayer(Index).InShop).BuyRate / 100
    
    If CustomShop_Sell(Index, TempPlayer(Index).InShop, HasItem(Index, ItemNum), Amount) Then
        Exit Sub
    End If

    If item(ItemNum).Stackable = 1 Then
        If GetPlayerInvItemValue(Index, HasItem(Index, ItemNum)) > Amount Then Amount = GetPlayerInvItemValue(Index, HasItem(Index, ItemNum))
    
        TakeInvItem Index, ItemNum, Amount
        GiveInvItem Index, Shop(TempPlayer(Index).InShop).ShopCurrency, multiplier * item(ItemNum).MonetaryValue * Amount
    
        ' send confirmation message & reset their shop action
        PlayerMsg Index, "Trade successful.", BrightGreen
    Else
        For i = 1 To Amount
            If HasItem(Index, ItemNum) Then
                TakeInvItem Index, ItemNum, 1
                GiveInvItem Index, Shop(TempPlayer(Index).InShop).ShopCurrency, multiplier * item(ItemNum).MonetaryValue
            Else
                Amount = i
                Exit For
            End If
        Next
        ' send confirmation message & reset their shop action
        PlayerMsg Index, "Trade successful.", BrightGreen
    End If
End Sub

Public Sub CacheResources(ByVal MapNum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).x = x
                ResourceCache(MapNum).ResourceData(Resource_Count).y = y
                ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = RAND(Resource(Map(MapNum).Tile(x, y).Data1).minHealth, Resource(Map(MapNum).Tile(x, y).Data1).maxHealth)
            End If

        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal Index As Long, ByVal BankTab As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(Index, BankTab, oldSlot)
    OldValue = GetPlayerBankItemValue(Index, BankTab, oldSlot)
    NewNum = GetPlayerBankItemNum(Index, BankTab, newSlot)
    NewValue = GetPlayerBankItemValue(Index, BankTab, newSlot)
    
    SetPlayerBankItemNum Index, BankTab, newSlot, OldNum
    SetPlayerBankItemValue Index, BankTab, newSlot, OldValue
    
    SetPlayerBankItemNum Index, BankTab, oldSlot, NewNum
    SetPlayerBankItemValue Index, BankTab, oldSlot, NewValue
        
    SendBank Index
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(Index, oldSlot)
    OldValue = GetPlayerInvItemValue(Index, oldSlot)
    NewNum = GetPlayerInvItemNum(Index, newSlot)
    NewValue = GetPlayerInvItemValue(Index, newSlot)
    SetPlayerInvItemNum Index, newSlot, OldNum
    SetPlayerInvItemValue Index, newSlot, OldValue
    SetPlayerInvItemNum Index, oldSlot, NewNum
    SetPlayerInvItemValue Index, oldSlot, NewValue
    SendInventory Index
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(Index, oldSlot)
    NewNum = GetPlayerSpell(Index, newSlot)
    SetPlayerSpell Index, oldSlot, NewNum
    SetPlayerSpell Index, newSlot, OldNum
    SendPlayerSpells Index
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long)

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(Index, GetPlayerEquipment(Index, EqSlot)) > 0 Then
        GiveInvItem Index, GetPlayerEquipment(Index, EqSlot), Player(Index).Equipment(EqSlot).value
        ' send the sound
        SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, EqSlot)
        ' remove equipment
        SetPlayerEquipment Index, 0, EqSlot
        
        If GetPlayerEquipment(Index, Weapon) > 0 Then
            If item(GetPlayerEquipment(Index, Weapon)).isTwoHander Then
                If Player(Index).Gender = Male Then
                    ' Player(Index).Sprite = 3
                Else
                    ' Player(Index).Sprite = 6
                End If
            End If
        End If
        
        If GetPlayerEquipment(Index, Shield) > 0 Then
            If Player(Index).Gender = Male Then
                ' Player(Index).Sprite = 2
            Else
                ' Player(Index).Sprite = 5
            End If
        Else
            If Player(Index).Gender = Male Then
                ' Player(Index).Sprite = 1
            Else
                ' Player(Index).Sprite = 4
            End If
        End If
        
        SendUpdatePlayerSprite (Index)
        SendWornEquipment Index
        SendMapEquipment Index
    Else
        PlayerMsg Index, "Your inventory is full.", BrightRed
    End If

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef dir As Byte) As Boolean
    If Not blockvar And (2 ^ dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function
