Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal Index As Long)
    If Not IsPlaying(Index) Then
        Call JoinGame(Index)
        Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal Index As Long)
    Dim i As Long
    
    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True
    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    
    ' send the login ok
    SendLoginOk Index
    
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendItems(Index)
    Call SendAnimations(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendSpells(Index)
    Call SendResources(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendPlayerSpells(Index)
    Call SendHotbar(Index)
    Call SendClanListTo(Index, 0)
    
    ' send vitals, exp + stats
    SendSkills Index
    
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(Index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Send welcome messages
    Call SendWelcome(Index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next
    
    ' Send the flag so they know they can start doing stuff
    SendInGame Index
End Sub

Sub LeftGame(ByVal Index As Long)
    Dim n As Long, i As Long
    Dim tradeTarget As Long
    
    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
        ' cancel any trade they're in
        If TempPlayer(Index).InTrade > 0 Then
            tradeTarget = TempPlayer(Index).InTrade
            PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has declined the trade.", BrightRed
            ' clear out trade
            For i = 1 To MAX_INV
                TempPlayer(tradeTarget).TradeOffer(i).Num = 0
                TempPlayer(tradeTarget).TradeOffer(i).value = 0
            Next
            TempPlayer(tradeTarget).InTrade = 0
            SendCloseTrade tradeTarget
        End If
        
        ' Leave the clan if the player is in a clan.
        If TempPlayer(Index).inClan Then
            Dim x() As Byte
            Call HandleLeaveClan(Index, x, 0, 0)
        End If

        ' save and clear data.
        Call SavePlayer(Index)
        Call SaveBank(Index)
        Call ClearBank(Index)

        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(Index)
End Sub

Function GetPlayerProtection(ByVal Index As Long) As Long
    GetPlayerProtection = 0
End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    CanPlayerCriticalHit = False
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    CanPlayerBlockHit = False
End Function

Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal Ignore As Boolean = False)
    Dim ShopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(Index) Then
        SendPlayerXYToMap Index
    End If
    
    ' clear target
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = TARGET_TYPE_NONE
    SendTarget Index

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If Map(MapNum).Instance = 1 And Ignore = False Then
        Call CopyMap(MAX_MAPS - MAX_PLAYERS + Index, MapNum)
        MapNum = MAX_MAPS - MAX_PLAYERS + Index
    End If

    If OldMap <> MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    
    ' send player's equipment to new map
    SendMapEquipment Index
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(MapNum) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    SendMapEquipmentTo i, Index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If MapNpc(OldMap).Npc(i).Num > 0 Then
                MapNpc(OldMap).Npc(i).Vital(Vitals.HitPoints) = GetNpcMaxVital(MapNpc(OldMap).Npc(i).Num, Vitals.HitPoints)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    TempPlayer(Index).GettingMap = YES
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    Buffer.WriteLong MapNum
    Buffer.WriteLong Map(MapNum).Revision
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal dir As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer, MapNum As Long
    Dim x As Long, y As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
    Dim TileType As Long, VitalType As Long, Colour As Long, Amount As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or dir < DIR_UP Or dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, dir)
    Moved = NO
    MapNum = GetPlayerMap(Index)
    
    y = GetPlayerY(Index)
    x = GetPlayerX(Index)
    
    Select Case dir
        Case DIR_UP
            y = y - 1
        Case DIR_DOWN
            y = y + 1
        Case DIR_RIGHT
            x = x + 1
        Case DIR_LEFT
            x = x - 1
    End Select
    
    If y < 0 Then
        If Map(GetPlayerMap(Index)).Up > 0 Then
            NewMapY = Map(Map(GetPlayerMap(Index)).Up).MaxY
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, x, NewMapY)
            Moved = YES
        End If
    ElseIf y > Map(GetPlayerMap(Index)).MaxY Then
        If Map(GetPlayerMap(Index)).Down > 0 Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, x, 0)
            Moved = YES
        End If
    End If
    
    If x < 0 Then
        If Map(GetPlayerMap(Index)).Left > 0 Then
            NewMapX = Map(Map(GetPlayerMap(Index)).Left).MaxX
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, NewMapX, y)
            Moved = YES
        End If
    ElseIf x > Map(GetPlayerMap(Index)).MaxX Then
        If Map(GetPlayerMap(Index)).Right > 0 Then
            Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, y)
            Moved = YES
        End If
    End If
    
    If Moved Then
        TempPlayer(Index).Target = 0
        TempPlayer(Index).TargetType = TARGET_TYPE_NONE
        SendTarget Index
        y = GetPlayerY(Index)
        x = GetPlayerX(Index)
    End If
    
    If Not isDirBlocked(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).DirBlock, dir + 1) Then
        If Map(GetPlayerMap(Index)).Tile(x, y).Type <> TILE_TYPE_BLOCKED Then
            If Map(GetPlayerMap(Index)).Tile(x, y).Type <> TILE_TYPE_RESOURCE Then
                If Map(GetPlayerMap(Index)).Tile(x, y).Type <> TILE_TYPE_OBJECT Then
                
                    If CustomMapTile(Index, x, y) Then
                        If Map(GetPlayerMap(Index)).Tile(x, y).Type <> TILE_TYPE_KEY Or (Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES) Then
                            Call SetPlayerY(Index, y)
                            Call SetPlayerX(Index, x)
                            SendPlayerMove Index, movement, sendToSelf
                            Moved = YES
                        End If
                        
                    End If
                    
                End If
            End If
        End If
    End If
    
With Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            MapNum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(Index, MapNum, x, y)
            Moved = YES
        End If
    
        ' Check to see if the tile is a door tile, and if so warp them
        If .Type = TILE_TYPE_DOOR Then
            MapNum = .Data1
            x = .Data2
            y = .Data3
            ' send the animation to the map
            SendDoorAnimation GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
            Call PlayerWarp(Index, MapNum, x, y)
            Moved = YES
        End If
    
        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            x = .Data1
            y = .Data2
    
            If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                SendMapKey Index, x, y, 1
                Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
            End If
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop Index, x
                    TempPlayer(Index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank Index
            TempPlayer(Index).InBank = True
            Moved = YES
        End If
        
        
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove Index, MOVING_WALKING, GetPlayerDir(Index)
            Moved = YES
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
    End If

End Sub

Sub ForcePlayerMove(ByVal Index As Long, ByVal movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If movement < 1 Or movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(Index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(Index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove Index, Direction, movement, True
End Sub

Sub CheckEquippedItems(ByVal Index As Long)
    Dim Slot As Long
    Dim ItemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(Index, i)

        If ItemNum > 0 Then
            Select Case i
                Case Equipment.Weapon
                    If item(ItemNum).ItemType <> ITEM_TYPE_WEAPON Then SetPlayerEquipment Index, 0, i
                Case Equipment.Torso
                    If item(ItemNum).ItemType <> ITEM_TYPE_TORSO Then SetPlayerEquipment Index, 0, i
                Case Equipment.Helmet
                    If item(ItemNum).ItemType <> ITEM_TYPE_HELMET Then SetPlayerEquipment Index, 0, i
                Case Equipment.Shield
                    If item(ItemNum).ItemType <> ITEM_TYPE_SHIELD Then SetPlayerEquipment Index, 0, i
                Case Equipment.Cape
                    If item(ItemNum).ItemType <> ITEM_TYPE_CAPE Then SetPlayerEquipment Index, 0, i
                Case Equipment.Amulet
                    If item(ItemNum).ItemType <> ITEM_TYPE_AMULET Then SetPlayerEquipment Index, 0, i
                Case Equipment.Arrows
                    If item(ItemNum).ItemType <> ITEM_TYPE_ARROWS Then SetPlayerEquipment Index, 0, i
                Case Equipment.Legs
                    If item(ItemNum).ItemType <> ITEM_TYPE_LEGS Then SetPlayerEquipment Index, 0, i
                Case Equipment.Gloves
                    If item(ItemNum).ItemType <> ITEM_TYPE_GLOVES Then SetPlayerEquipment Index, 0, i
                Case Equipment.Boots
                    If item(ItemNum).ItemType <> ITEM_TYPE_BOOTS Then SetPlayerEquipment Index, 0, i
                Case Equipment.Ring
                    If item(ItemNum).ItemType <> ITEM_TYPE_RING Then SetPlayerEquipment Index, 0, i
            End Select
        Else
            SetPlayerEquipment Index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    If item(ItemNum).Stackable = 1 Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next
    End If

    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next

End Function

Function FindHowManyOpenInvSlots(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long, amt As Long
    
        ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, i) = 0 Then amt = amt + 1
        If GetPlayerInvItemNum(Index, i) = ItemNum And item(ItemNum).Stackable = 1 Then amt = amt + 1
    Next
    
    FindHowManyOpenInvSlots = amt
End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal BankTab As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    If Not IsPlaying(Index) Then Exit Function
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, BankTab, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, BankTab, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If item(ItemNum).Stackable = 1 Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function

Function TakeInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    
    TakeInvItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = ItemNum Then
            If item(ItemNum).Stackable = 1 Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    Call SendInventoryUpdate(Index, i)
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                Exit Function
            End If
        End If

    Next

End Function

Function TakeInvSlot(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ItemNum
    
    TakeInvSlot = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or invSlot <= 0 Or invSlot > MAX_ITEMS Then
        Exit Function
    End If
    
    ItemNum = GetPlayerInvItemNum(Index, invSlot)

    If item(ItemNum).Stackable = 1 Then

        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(Index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(Index, invSlot, GetPlayerInvItemValue(Index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(Index, invSlot, 0)
        Call SetPlayerInvItemValue(Index, invSlot, 0)
        Exit Function
    End If

End Function

Function GiveInvItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        GiveInvItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(Index, ItemNum)

    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, ItemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        If sendUpdate Then Call SendInventoryUpdate(Index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
        GiveInvItem = False
    End If

End Function

Function HasSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim i As Long
    Dim n As Long
    Dim MapNum As Long
    Dim Msg As String

    If Not IsPlaying(Index) Then Exit Sub
    MapNum = GetPlayerMap(Index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' our drop?
            If CanPlayerPickupItem(Index, i) Then
                ' Check if item is at the same location as the player
                If (MapItem(MapNum, i).x = GetPlayerX(Index)) Then
                    If (MapItem(MapNum, i).y = GetPlayerY(Index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(Index, MapItem(MapNum, i).Num)
    
                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(Index, n, MapItem(MapNum, i).Num)
    
                            If item(GetPlayerInvItemNum(Index, n)).Stackable = 1 Then
                                Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(MapNum, i).value)
                                Msg = MapItem(MapNum, i).value & " " & Trim$(item(GetPlayerInvItemNum(Index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(Index, n, 0)
                                Msg = Trim$(item(GetPlayerInvItemNum(Index, n)).Name)
                            End If
    
                            ' Erase item from the map
                            ClearMapItem i, MapNum
                            
                            Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0)
                            SendActionMsg GetPlayerMap(Index), Msg, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                            Exit For
                        Else
                            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Function CanPlayerPickupItem(ByVal Index As Long, ByVal mapItemNum As Long)
Dim MapNum As Long

    MapNum = GetPlayerMap(Index)
    
    ' no lock or locked to player?
    If MapItem(MapNum, mapItemNum).Owner = vbNullString Or MapItem(MapNum, mapItemNum).Owner = Trim$(GetPlayerName(Index)) Then
        CanPlayerPickupItem = True
        Exit Function
    End If
    
    CanPlayerPickupItem = False
End Function

Sub PlayerMapDropItem(ByVal Index As Long, ByVal invNum As Long, ByVal Amount As Long)
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or invNum <= 0 Or invNum > MAX_INV Then
        Exit Sub
    End If
    
    ' check the player isn't doing something
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Or TempPlayer(Index).InTrade > 0 Then Exit Sub

    If (GetPlayerInvItemNum(Index, invNum) > 0) Then
        If (GetPlayerInvItemNum(Index, invNum) <= MAX_ITEMS) Then
            i = FindOpenMapItemSlot(GetPlayerMap(Index))

            If i <> 0 Then
                MapItem(GetPlayerMap(Index), i).Num = GetPlayerInvItemNum(Index, invNum)
                MapItem(GetPlayerMap(Index), i).x = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), i).y = GetPlayerY(Index)
                MapItem(GetPlayerMap(Index), i).Owner = Trim$(GetPlayerName(Index))
                MapItem(GetPlayerMap(Index), i).tmr = 5000

                If item(GetPlayerInvItemNum(Index, invNum)).Stackable = 1 Then

                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, invNum) Then
                        MapItem(GetPlayerMap(Index), i).value = GetPlayerInvItemValue(Index, invNum)
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & GetPlayerInvItemValue(Index, invNum) & " " & Trim$(item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, invNum, 0)
                        Call SetPlayerInvItemValue(Index, invNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), i).value = Amount
                        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & Amount & " " & Trim$(item(GetPlayerInvItemNum(Index, invNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, invNum, GetPlayerInvItemValue(Index, invNum) - Amount)
                    End If

                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), i).value = 0
                    ' send message
                    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " drops " & CheckGrammar(Trim$(item(GetPlayerInvItemNum(Index, invNum)).Name)) & ".", Yellow)
                    Call SetPlayerInvItemNum(Index, invNum, 0)
                    Call SetPlayerInvItemValue(Index, invNum, 0)
                End If

                ' Send inventory update
                Call SendInventoryUpdate(Index, invNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(Index), i).Num, Amount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), Trim$(GetPlayerName(Index)))
            Else
                Call PlayerMsg(Index, "Too many items already on the ground.", BrightRed)
            End If
        End If
    End If

End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    ' Player(Index).Gender = 0
    ' If Sprite > 0 And Sprite < 4 Then Player(Index).Gender = Male
    ' If Sprite > 3 And Sprite < 7 Then Player(Index).Gender = Female
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerSkillLevel(ByVal Index As Long, ByVal Skill As Long) As Long
    GetPlayerSkillLevel = Player(Index).Skill(Skill).Level
End Function

Function SetPlayerSkillLevel(ByVal Index As Long, ByVal Skill As Long, ByVal Level As Long)
    Player(Index).Skill(Skill).Level = Level
End Function

Function GetPlayerSkillXP(ByVal Index As Long, ByVal Skill As Long) As Long
    GetPlayerSkillXP = Player(Index).Skill(Skill).xp
End Function

Sub SetPlayerSkillXP(ByVal Index As Long, ByVal Skill As Long, ByVal exp As Long)
    Player(Index).Skill(Skill).xp = Player(Index).Skill(Skill).xp + exp
End Sub

Sub GivePlayerSkillXP(ByVal Index As Long, ByVal Skill As Long, ByVal xp As Long)
Dim newXP As Long

    With Player(Index).Skill(Skill)
    
        If .xp = 200000000 Then Exit Sub
    
        If .xp + xp < 200000000 Then
            .xp = .xp + xp
        ElseIf .xp + xp > 200000000 Then
            .xp = 200000000
        End If
        
        If (.Level < 99 And Skill <> Skills.Dungeoneering) Or (.Level < 120 And Skill = Skills.Dungeoneering) Then
            Do While GetPlayerNextLevelXP(Index, Skill) <= .xp
                newXP = .xp - GetPlayerNextLevelXP(Index, Skill)
                .xp = newXP
                .Level = .Level + 1
                
                PlayerMsg Index, "Congratulations! You leveled up " & GetSkillName(Skill) & ".", Yellow
            Loop
            
            Call SendSkill(Index, Skill)
        End If
        
    End With
    
End Sub

Function GetPlayerNextLevelXP(ByVal Index As Long, ByVal Skill As Long)

    With Player(Index).Skill(Skill)
        GetPlayerNextLevelXP = ((1 / 4.25) * (.Level + (300 * 2 ^ (.Level / 7))))
    End With

End Function

Function WeaponDamageToResourceDamage(ByVal Index As Long, ByVal Skill As Long, ByVal ItemNum As Long)
    Dim Damage As Long
    
    If ItemNum = 0 Then Exit Function
    Damage = item(ItemNum).Damage
    
    WeaponDamageToResourceDamage = Damage
End Function

Function GetPlayerAccess(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If Player(Index).Map > MAX_MAPS Then Player(Index).Map = 1
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Map = MapNum
    End If

End Sub

Function GetPlayerX(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)
    Player(Index).dir = dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long) As Long
    If Index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(Index).Inv(invSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(invSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    GetPlayerInvItemValue = Player(Index).Inv(invSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(invSlot).value = ItemValue
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Player(Index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal SpellNum As Long)
    Player(Index).Spell(spellslot) = SpellNum
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal equipmentslot As Equipment) As Long

    If Index > MAX_PLAYERS Then Exit Function
    If equipmentslot = 0 Or equipmentslot > Equipment_Count - 1 Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(equipmentslot).Num
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal equipmentslot As Equipment, Optional ByVal value As Long = 0)
    Player(Index).Equipment(equipmentslot).Num = invNum
    Player(Index).Equipment(equipmentslot).value = value
End Sub

' ToDo
Sub OnDeath(ByVal Index As Long)
    Dim i As Long

    ' Drop all worn items
    If Map(Player(Index).Map).LoseItemsOnDeath Then
        For i = 1 To Equipment.Equipment_Count - 1
            If Player(Index).Equipment(i).Num > 0 Then
                PlayerMapDropItem Index, Player(Index).Equipment(i).Num, Player(Index).Equipment(i).value
                Player(Index).Equipment(i).Num = 0
                Player(Index).Equipment(i).value = 0
            End If
        Next
        
        For i = 1 To MAX_INV
            If Player(Index).Inv(i).Num > 0 Then
                PlayerMapDropItem Index, Player(Index).Inv(i).Num, Player(Index).Inv(i).value
                Player(Index).Inv(i).Num = 0
                Player(Index).Inv(i).value = 0
            End If
        Next
    End If

    ' Warp player away
    Call SetPlayerDir(Index, DIR_DOWN)
    
    With Map(GetPlayerMap(Index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp Index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(Index, START_MAP, START_X, START_Y)
        End If
    End With
    
    ' Clear spell casting
    TempPlayer(Index).spellBuffer.Spell = 0
    TempPlayer(Index).spellBuffer.Timer = 0
    TempPlayer(Index).spellBuffer.Target = 0
    TempPlayer(Index).spellBuffer.tType = 0
    Call SendClearSpellBuffer(Index)
    
    ' Restore vitals
    For i = 1 To Skills.Skill_Count - 1
        Player(Index).Skill(i).Level = Player(Index).Skill(i).MaxLevel
    Next
    Call SendSkills(Index)

End Sub

Sub CheckResource(ByVal Index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    
    If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = Map(GetPlayerMap(Index)).Tile(x, y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count

            If ResourceCache(GetPlayerMap(Index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If

        Next
        
        If Resource_num > 0 Then
            With Resource(Resource_num)
            
                ' Can they collect it?
                For i = 1 To Skills.Skill_Count - 1
                    If .SkillReq(i) > Player(Index).Skill(i).Level Then
                        PlayerMsg Index, "You need to have a " & GetSkillName(i) & " level of " & .SkillReq(i) & " to collect this resource.", BrightRed
                        Exit Sub
                    End If
                Next
            
                ' Is it already collected?
                If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState <> 0 Then
                    ' send message if it exists
                    If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then
                        SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                    End If
                End If
                
                ' Weapon requirements
                If .RequireWeapon > 0 Then
                    If GetPlayerEquipment(Index, Weapon) > 0 Then
                        If item(GetPlayerEquipment(Index, Weapon)).EquipmentType <> .ToolRequired Then
                            PlayerMsg Index, "You need the proper tool to collect this resource.", BrightRed
                            Exit Sub
                        End If
                    Else
                        PlayerMsg Index, "You need a tool to collect this resource.", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' Do they have room to collect the reward?
                For i = 1 To 10
                    If .RewardItem(i) > 0 Then
                        If FindOpenInvSlot(Index, .RewardItem(i)) = 0 Then
                            PlayerMsg Index, "You have no free inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If
                Next
                
                rX = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).x
                rY = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).y
                
                Damage = WeaponDamageToResourceDamage(Index, Skills.Agility, GetPlayerEquipment(Index, Weapon))
                
                If Damage > 0 Then
                    ' Did we cut it down?
                    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
                        SendActionMsg GetPlayerMap(Index), "-" & ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                        ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                        ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                        SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                        ' send message if it exists
                        If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then
                            SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                        End If
                        
                        ' Give the rewarditem.
                        For i = 1 To 10
                            If .RewardItem(i) > 0 Then
                                If Random(0, 100) >= .RewardChance(i) Then
                                    If .RewardValue(i) = 0 Then .RewardValue(i) = 1
                                    If FindOpenInvSlot(Index, .RewardItem(i)) > 0 Then
                                        GiveInvItem Index, .RewardItem(i), .RewardValue(i)
                                    Else
                                        Call SpawnItem(.RewardItem(i), .RewardValue(i), GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index))
                                        Call PlayerMsg(Index, "You notice something falls to the floor.", Magenta)
                                    End If
                                End If
                            End If
                        Next
                        
                        ' Give the XP
                        For i = 1 To Skills.Skill_Count - 1
                            If .RewardXP(i) > 0 Then
                                Call GivePlayerSkillXP(Index, i, .RewardXP(i))
                                Call SendSkill(Index, i)
                            End If
                        Next
                        
                        SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                    Else
                        ' Nope. Just display the damage.
                        ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - Damage
                        SendActionMsg GetPlayerMap(Index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
                        SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                    End If
                Else
                    SendActionMsg GetPlayerMap(Index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                End If
            End With
        End If
    End If

End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankTab As Long, ByVal bankslot As Long) As Long
    GetPlayerBankItemNum = Bank(Index).BankTab(BankTab).item(bankslot).Num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankTab As Long, ByVal bankslot As Long, ByVal ItemNum As Long)
    Bank(Index).BankTab(BankTab).item(bankslot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankTab As Long, ByVal bankslot As Long) As Long
    GetPlayerBankItemValue = Bank(Index).BankTab(BankTab).item(bankslot).value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankTab As Long, ByVal bankslot As Long, ByVal ItemValue As Long)
    Bank(Index).BankTab(BankTab).item(bankslot).value = ItemValue
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal BankTab As Long, ByVal invSlot As Long, ByVal Amount As Long)
Dim bankslot
Dim i As Long

    If invSlot < 0 Or invSlot > MAX_INV Then
        Exit Sub
    End If
    
    If Amount < 0 Then
        Exit Sub
    End If
    
    bankslot = FindOpenBankSlot(Index, BankTab, GetPlayerInvItemNum(Index, invSlot))
        
    If bankslot > 0 Then
        If item(GetPlayerInvItemNum(Index, invSlot)).Stackable = 1 Then
            
            If Amount > GetPlayerInvItemValue(Index, invSlot) Or Amount = 0 Then Amount = GetPlayerInvItemValue(Index, invSlot)
            
            If GetPlayerBankItemNum(Index, BankTab, bankslot) = GetPlayerInvItemNum(Index, invSlot) Then
                Call SetPlayerBankItemValue(Index, BankTab, bankslot, GetPlayerBankItemValue(Index, BankTab, bankslot) + Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), Amount)
            Else
                Call SetPlayerBankItemNum(Index, BankTab, bankslot, GetPlayerInvItemNum(Index, invSlot))
                Call SetPlayerBankItemValue(Index, BankTab, bankslot, Amount)
                Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), Amount)
            End If
        Else
            Dim ItemNum As Long
            ItemNum = GetPlayerInvItemNum(Index, invSlot)
            If Amount = 0 Then Amount = MAX_INV
        
            ' Not a currency.
            For i = 1 To Amount
                If HasItem(Index, ItemNum) > 0 Then
                    If GetPlayerBankItemNum(Index, BankTab, bankslot) = ItemNum Then
                        Call SetPlayerBankItemValue(Index, BankTab, bankslot, GetPlayerBankItemValue(Index, BankTab, bankslot) + 1)
                        Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), 0)
                    Else
                        Call SetPlayerBankItemNum(Index, BankTab, bankslot, GetPlayerInvItemNum(Index, invSlot))
                        Call SetPlayerBankItemValue(Index, BankTab, bankslot, 1)
                        Call TakeInvItem(Index, GetPlayerInvItemNum(Index, invSlot), 0)
                    End If
                Else
                    Exit For
                End If
                
                invSlot = HasItemInInv(Index, ItemNum)
                If invSlot = 0 Then Exit For
            Next
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal BankTab As Long, ByVal bankslot As Long, ByVal Amount As Long)
Dim invSlot
Dim i As Long

    If bankslot < 0 Or bankslot > MAX_BANK Then
        Exit Sub
    End If
    
    If Amount < 0 Then
        Exit Sub
    End If
    
    invSlot = FindOpenInvSlot(Index, GetPlayerBankItemNum(Index, BankTab, bankslot))
    
    If invSlot > 0 Then
        If item(GetPlayerBankItemNum(Index, BankTab, bankslot)).Stackable = 1 Then
            If GetPlayerBankItemValue(Index, BankTab, bankslot) < Amount Or Amount = 0 Then Amount = GetPlayerBankItemValue(Index, BankTab, bankslot)
            Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankTab, bankslot), Amount)
            Call SetPlayerBankItemValue(Index, BankTab, bankslot, GetPlayerBankItemValue(Index, BankTab, bankslot) - Amount)
            If GetPlayerBankItemValue(Index, BankTab, bankslot) <= 0 Then
                Call SetPlayerBankItemNum(Index, BankTab, bankslot, 0)
                Call SetPlayerBankItemValue(Index, BankTab, bankslot, 0)
            End If
        Else
            If Amount = 0 Then Amount = GetPlayerBankItemValue(Index, BankTab, bankslot)
            For i = 1 To Amount
                If GetPlayerBankItemValue(Index, BankTab, bankslot) > 1 Then
                    Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankTab, bankslot), 0)
                    Call SetPlayerBankItemValue(Index, BankTab, bankslot, GetPlayerBankItemValue(Index, BankTab, bankslot) - 1)
                Else
                    Call GiveInvItem(Index, GetPlayerBankItemNum(Index, BankTab, bankslot), 0)
                    Call SetPlayerBankItemNum(Index, BankTab, bankslot, 0)
                    Call SetPlayerBankItemValue(Index, BankTab, bankslot, 0)
                End If
                
                invSlot = FindOpenInvSlot(Index, GetPlayerBankItemNum(Index, BankTab, bankslot))
                If invSlot = 0 Then
                    Exit For
                End If
            Next
        End If
    End If
    
    SaveBank Index
    SavePlayer Index
    SendBank Index

End Sub

Public Sub KillPlayer(ByVal Index As Long)
    Call OnDeath(Index)
End Sub

Public Sub UseItem(ByVal Index As Long, ByVal invNum As Long)
Dim n As Long, i As Long, tempItem As Long, tempAmount As Long, x As Long, y As Long, ItemNum As Long, ItemAmount As Long

    If invNum < 1 Or invNum > MAX_ITEMS Then Exit Sub
    ItemNum = GetPlayerInvItemNum(Index, invNum)
    ItemAmount = GetPlayerInvItemValue(Index, invNum)
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    With item(ItemNum)
        ' Global requirements
        For i = 1 To Skills.Skill_Count - 1
            If GetPlayerSkillLevel(Index, i) < .SkillWearReq(i) Then
                PlayerMsg Index, "You need to have a " & GetSkillName(i) & " level of " & .SkillWearReq(i) & " to use this item.", BrightRed
                Exit Sub
            End If
        Next
        
        If GetPlayerAccess(Index) < .AccessRequired Then
            PlayerMsg Index, "You are not allowed to use this item.", BrightRed
            Exit Sub
        End If
        
        ' Do item scripts
        If UseItemScript(Index, ItemNum, invNum) Then Exit Sub
        
        Select Case .ItemType
            Case ITEM_TYPE_HELMET To ITEM_TYPE_RING
                
                If GetPlayerEquipment(Index, .ItemType) > 0 Then
                    tempItem = Player(Index).Equipment(.ItemType).Num
                    tempAmount = Player(Index).Equipment(.ItemType).value
                End If
                    
                SetPlayerEquipment Index, ItemNum, .ItemType, ItemAmount
                TakeInvItem Index, ItemNum, ItemAmount
                
                If tempItem > 0 Then
                    GiveInvItem Index, tempItem, tempAmount
                    tempItem = 0
                    tempAmount = 0
                End If
                
                If .ItemType = ITEM_TYPE_WEAPON Then
                    If .isTwoHander Then
                        ' Is there a shield? The weapon will be switched out automatically.
                        If GetPlayerEquipment(Index, Shield) > 0 Then
                            If FindHowManyOpenInvSlots(Index, GetPlayerEquipment(Index, Shield)) > 0 Then
                                GiveInvItem Index, GetPlayerEquipment(Index, Shield), Player(Index).Equipment(Equipment.Shield).value
                                SetPlayerEquipment Index, 0, Shield, 0
                            End If
                        End If
                        
                        If Player(Index).Gender = Male Then
                            ' Player(Index).Sprite = 3
                        Else
                            ' Player(Index).Sprite = 6
                        End If
                    Else
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
                    End If
                ElseIf .ItemType = ITEM_TYPE_SHIELD Then
                    ' Is there a weapon that's a two handed weapon?
                    If GetPlayerEquipment(Index, Weapon) > 0 Then
                        If item(GetPlayerEquipment(Index, Weapon)).isTwoHander Then
                            If FindHowManyOpenInvSlots(Index, GetPlayerEquipment(Index, Weapon)) > 0 Then
                                GiveInvItem Index, GetPlayerEquipment(Index, Weapon), Player(Index).Equipment(Equipment.Weapon).value
                                SetPlayerEquipment Index, 0, Weapon, 0
                            End If
                        End If
                    End If
                    
                    If Player(Index).Gender = Male Then
                        ' Player(Index).Sprite = 2
                    Else
                        ' Player(Index).Sprite = 5
                    End If
                End If
                
                Call SendUpdatePlayerSprite(Index)
                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
            Case ITEM_TYPE_CONSUME
                
                If TempPlayer(Index).UseItemTimer > 0 Then Exit Sub
            
                For i = 1 To Vitals.Vital_Count - 1
                    Player(Index).Skill(i).Level = .AddVital(i)
                    If .AddVital(i) > 0 Then SendActionMsg GetPlayerMap(Index), "NOPE" & " + " & .AddVital(i), White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                Next
                
                TakeInvItem Index, ItemNum, 1
                
                TempPlayer(Index).UseItemTimer = 500
            
            Case ITEM_TYPE_KEY
                x = GetPlayerX(Index)
                y = GetPlayerY(Index)
            
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        If y > 0 Then y = y - 1
                    Case DIR_DOWN
                        If y < Map(GetPlayerMap(Index)).MaxY Then y = y + 1
                    Case DIR_LEFT
                        If x > 0 Then x = x - 1
                    Case DIR_RIGHT
                        If x < Map(GetPlayerMap(Index)).MaxX Then x = x + 1
                End Select
                
                If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY Then
                    If ItemNum = Map(GetPlayerMap(Index)).Tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        SendMapKey Index, x, y, 1
                        Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(Index), item(ItemNum).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If Map(GetPlayerMap(Index)).Tile(x, y).Data2 = 1 Then
                            Call TakeInvItem(Index, ItemNum, 0)
                            Call PlayerMsg(Index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If
        End Select
    End With
    
    SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, ItemNum

End Sub

Public Function GetPlayerOffense(ByVal PlayerIndex As Long, ByVal Style As Byte) As Long
    If Style < 0 Or Style >= CombatStyles.Count Then Exit Function
    
    Dim i As Long
    Dim Amount As Long
    Amount = 0
    
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(PlayerIndex, i) > 0 Then
            If Style <> 0 Then Amount = Amount + item(GetPlayerEquipment(PlayerIndex, i)).Offense(Style)
        End If
    Next
    
    GetPlayerOffense = Amount
End Function


Public Function GetPlayerDefense(ByVal PlayerIndex As Long, ByVal Style As Byte) As Long
    If Style < 0 Or Style >= CombatStyles.Count Then Exit Function
    
    Dim i As Long
    Dim Amount As Long
    Amount = 0
    
    For i = 1 To Equipment.Equipment_Count
        If GetPlayerEquipment(PlayerIndex, i) > 0 Then
            Amount = Amount + item(GetPlayerEquipment(PlayerIndex, i)).Defense(Style)
        End If
    Next
    
    GetPlayerDefense = Amount
End Function

Public Function HasItemInInv(ByVal Index As Long, ByVal item As Long) As Long
Dim i As Long

    For i = 1 To MAX_INV
        If Player(Index).Inv(i).Num = item Then
            HasItemInInv = i
            Exit Function
        End If
    Next
    
    HasItemInInv = 0
End Function
