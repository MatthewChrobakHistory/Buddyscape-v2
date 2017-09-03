Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, x As Long
    Dim tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long, tmr10000 As Long, tmr250 As Long
    Dim tmr30000 As Long, StatsChanged As Boolean
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long

    ServerOnline = True

    Do While ServerOnline
        tick = GetTickCount
        ElapsedTime = tick - FrameTime
        FrameTime = tick
        
        If tick > tmr25 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If GetTickCount > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastingTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.Target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.Target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                End If
            Next
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr25 = GetTickCount + 25
        End If
        
        If tick > tmr250 Then
            For i = 1 To MAX_PLAYERS
                If TempPlayer(i).UseItemTimer > 0 Then
                    TempPlayer(i).UseItemTimer = TempPlayer(i).UseItemTimer - 250
                End If
                If TempPlayer(i).CanBeStunned > 0 Then
                    TempPlayer(i).CanBeStunned = TempPlayer(i).CanBeStunned - 250
                End If
            Next
            tmr250 = GetTickCount + 250
        End If

        ' Check for disconnections every half second
        If tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).state > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            UpdateMapLogic
            tmr500 = GetTickCount + 500
        End If

        If tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            tmr1000 = GetTickCount + 1000
        End If
        
        ' projectiles
        For i = 1 To MAX_MAPS
            For x = 1 To MAX_PROJECTILES
                If MapProjectile(i).Projectile(x).CreatorType > 0 Then
                    HandleProjecTile i, x
                End If
            Next
        Next

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If
        
        If tick > tmr10000 Then
            UpdateShopStocks
            tmr10000 = GetTickCount + 10000
        End If
        
        If tick > tmr30000 Then
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    StatsChanged = False
                    For x = 1 To Skills.Skill_Count - 1
                        If Player(i).Skill(x).Level > Player(i).Skill(x).MaxLevel Then
                            Player(i).Skill(x).Level = Player(i).Skill(x).Level - 1
                            StatsChanged = True
                        End If
                        If Player(i).Skill(x).Level < Player(i).Skill(x).MaxLevel Then
                            Player(i).Skill(x).Level = Player(i).Skill(x).Level + 1
                            StatsChanged = True
                        End If
                    Next
                    If StatsChanged Then SendSkills (i)
                    
                    If Player(i).SpecialAttack < 100 Then
                        If Player(i).SpecialAttack < 0 Then PlayerMsg i, "SPEC IS TOO LOW: " & Player(i).SpecialAttack, BrightRed
                        Player(i).SpecialAttack = Player(i).SpecialAttack + 10
                        If Player(i).SpecialAttack > 100 Then Player(i).SpecialAttack = 100
                        SendUpdateSpecial i
                    End If
                End If
            Next
            tmr30000 = tick + 30000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + ITEM_RESPAWN_RATE
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 300000
        End If

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < tick Then
            GameCPS = CPS
            TickCPS = tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub UpdateShopStocks()
Dim i As Long
Dim x As Long

    For i = 1 To MAX_SHOPS
        For x = 1 To MAX_TRADES
            With Shop(i).TradeItem(x)
                If .MaxStock <> -255 Then
                    If .Stock < .MaxStock Then
                        .Stock = .Stock + 1
                        Call UpdateShopStock(i, x)
                    End If
                End If
            End With
        Next
    Next
End Sub

Private Sub UpdateMapSpawnItems()
    Dim x As Long
    Dim y As Long
    Dim MapNum As Long
    Dim ItemNum As Long


    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For MapNum = 1 To MAX_MAPS
        For ItemNum = 1 To MAX_MAP_ITEMS
            With MapItem(MapNum, ItemNum)
                If Map(MapNum).Tile(.x, .y).Type = TILE_TYPE_ITEM Then
                    If .Num = Map(MapNum).Tile(.x, .y).Data1 Then
                        If .value = Map(MapNum).Tile(.x, .y).Data2 Then
                            Call ClearMapItem(ItemNum, MapNum)
                        End If
                    End If
                End If
            End With
        Next
    
    
        For x = 1 To Map(MapNum).MaxX
            For y = 1 To Map(MapNum).MaxY
                If Map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM Then
                    ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                    If item(Map(MapNum).Tile(x, y).Data1).Stackable And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                        Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                    Else
                        Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                    End If
                    
                End If
            Next
        Next
        
        Call SendMapItemsToAll(MapNum)
        DoEvents
    Next
    
End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, x As Long, MapNum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, NpcNum As Long
    Dim Target As Long, TargetType As Byte, DidWalk As Boolean, Buffer As clsBuffer, Resource_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean
    Dim Update As Boolean

    For MapNum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(MapNum, i).Num > 0 Then
                With MapItem(MapNum, i)
                
                    ' Make sure it isn't an item that spawns on the map.
                If Map(MapNum).Tile(.x, .y).Data1 = .Num And Map(MapNum).Tile(.x, .y).Type = TILE_TYPE_ITEM Then


                Else
                    If .tmr <= 0 Then
                        ' There is an item there.
                        Select Case .state
                            ' Only the owner could see it.
                            Case ITEM_DESPAWN_PLAYER
                                .state = ITEM_DESPAWN_GLOBAL
                                If item(.Num).Tradable Then .Owner = vbNullString
                                .tmr = ITEM_DESPAWN_TIME
                                Update = True
                            Case ITEM_DESPAWN_GLOBAL
                                ClearMapItem i, MapNum
                                Update = True
                        End Select
                    Else
                        .tmr = .tmr - 500
                    End If
                End If
                    

                    
                End With
            End If
        Next
        
        If Update Then
            SendMapItemsToAll (MapNum)
            Update = False
        End If
        
        '  Close the doors
        If False Then
            If TickCount > TempTile(MapNum).DoorTimer + 5000 Then
                For x1 = 0 To Map(MapNum).MaxX
                    For y1 = 0 To Map(MapNum).MaxY
                        If Map(MapNum).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(MapNum).DoorOpen(x1, y1) = YES Then
                            TempTile(MapNum).DoorOpen(x1, y1) = NO
                            SendMapKeyToMap MapNum, x1, y1, 0
                        End If
                    Next
                Next
            End If
        End If

        ' Respawning Resources
        If ResourceCache(MapNum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(MapNum).Resource_Count
                Resource_index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(i).x, ResourceCache(MapNum).ResourceData(i).y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(MapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(MapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(MapNum).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(MapNum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(MapNum).ResourceData(i).cur_health = RAND(Resource(Resource_index).minHealth, Resource(Resource_index).maxHealth)
                            SendResourceCacheToMap MapNum, i
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(MapNum) = YES Then
            TickCount = GetTickCount
            
            For x = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(MapNum).Npc(x).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
                
                    If MapNpc(MapNum).Npc(x).CanBeStunned > 0 Then
                        MapNpc(MapNum).Npc(x).CanBeStunned = MapNpc(MapNum).Npc(x).CanBeStunned - 500
                    End If

                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Type = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(NpcNum).Type = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(MapNum).Npc(x).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And MapNpc(MapNum).Npc(x).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        n = Npc(NpcNum).SightRange
                                        DistanceX = MapNpc(MapNum).Npc(x).x - GetPlayerX(i)
                                        DistanceY = MapNpc(MapNum).Npc(x).y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If Npc(NpcNum).Type = NPC_BEHAVIOUR_ATTACKONSIGHT Then
                                                MapNpc(MapNum).Npc(x).TargetType = 1 ' player
                                                MapNpc(MapNum).Npc(x).Target = i
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                    
                    If Npc(NpcNum).Team <> vbNullString Then
                        If MapNpc(MapNum).Npc(x).Target = 0 Then
                            If Not MapNpc(MapNum).Npc(x).StunDuration > 0 Then
                            
                                n = Npc(NpcNum).SightRange
                            
                                For i = 1 To MAX_MAP_NPCS
                                    If MapNpc(MapNum).Npc(i).Num > 0 And i <> x Then
                                        If Trim$(Npc(NpcNum).Team) <> Trim$(Npc(MapNpc(MapNum).Npc(i).Num).Team) Then
                                            
                                            
                                            If MapNpc(MapNum).Npc(x).x > MapNpc(MapNum).Npc(i).x Then
                                                DistanceX = MapNpc(MapNum).Npc(x).x - MapNpc(MapNum).Npc(i).x
                                            Else
                                                DistanceX = MapNpc(MapNum).Npc(i).x - MapNpc(MapNum).Npc(x).x
                                            End If
                                            
                                            If MapNpc(MapNum).Npc(x).y > MapNpc(MapNum).Npc(i).y Then
                                                DistanceY = MapNpc(MapNum).Npc(x).y - MapNpc(MapNum).Npc(i).y
                                            Else
                                                DistanceY = MapNpc(MapNum).Npc(i).y - MapNpc(MapNum).Npc(x).y
                                            End If
                                            
                                            If DistanceX <= n And DistanceY <= n Then
                                                If Npc(NpcNum).Type = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(NpcNum).Type = NPC_BEHAVIOUR_ATTACKWHENATTACKED Or Npc(NpcNum).Type = NPC_BEHAVIOUR_GUARD Then
                                                    MapNpc(MapNum).Npc(x).TargetType = TARGET_TYPE_NPC
                                                    MapNpc(MapNum).Npc(x).Target = i
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
                    If MapNpc(MapNum).Npc(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(MapNum).Npc(x).StunTimer + (MapNpc(MapNum).Npc(x).StunDuration * 1000) Then
                            MapNpc(MapNum).Npc(x).StunDuration = 0
                            MapNpc(MapNum).Npc(x).StunTimer = 0
                        End If
                    Else
                            
                        Target = MapNpc(MapNum).Npc(x).Target
                        TargetType = MapNpc(MapNum).Npc(x).TargetType
    
                        ' Check to see if its time for the npc to walk
                        If Npc(NpcNum).Type <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        
                            If TargetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If Target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = GetPlayerY(Target)
                                        TargetX = GetPlayerX(Target)
                                    Else
                                        MapNpc(MapNum).Npc(x).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(x).Target = 0
                                    End If
                                End If
                            
                            ElseIf TargetType = 2 Then 'npc
                                
                                If Target > 0 Then
                                    
                                    If MapNpc(MapNum).Npc(Target).Num > 0 Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = MapNpc(MapNum).Npc(Target).y
                                        TargetX = MapNpc(MapNum).Npc(Target).x
                                    Else
                                        MapNpc(MapNum).Npc(x).TargetType = 0 ' clear
                                        MapNpc(MapNum).Npc(x).Target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                
                                i = Int(Rnd * 5)
    
                                ' Lets move the npc
                                Select Case i
                                    Case 0
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_UP) Then
                                                Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_DOWN) Then
                                                Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_LEFT) Then
                                                Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 1
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_LEFT) Then
                                                Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_DOWN) Then
                                                Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_UP) Then
                                                Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 2
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_DOWN) Then
                                                Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_UP) Then
                                                Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_LEFT) Then
                                                Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 3
    
                                        ' Left
                                        If MapNpc(MapNum).Npc(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_LEFT) Then
                                                Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(MapNum).Npc(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_RIGHT) Then
                                                Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_UP) Then
                                                Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(MapNum, x, DIR_DOWN) Then
                                                Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                End Select
    
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(MapNum).Npc(x).x - 1 = TargetX And MapNpc(MapNum).Npc(x).y = TargetY Then
                                        If MapNpc(MapNum).Npc(x).dir <> DIR_LEFT Then
                                            Call NpcDir(MapNum, x, DIR_LEFT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(MapNum).Npc(x).x + 1 = TargetX And MapNpc(MapNum).Npc(x).y = TargetY Then
                                        If MapNpc(MapNum).Npc(x).dir <> DIR_RIGHT Then
                                            Call NpcDir(MapNum, x, DIR_RIGHT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(MapNum).Npc(x).x = TargetX And MapNpc(MapNum).Npc(x).y - 1 = TargetY Then
                                        If MapNpc(MapNum).Npc(x).dir <> DIR_UP Then
                                            Call NpcDir(MapNum, x, DIR_UP)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(MapNum).Npc(x).x = TargetX And MapNpc(MapNum).Npc(x).y + 1 = TargetY Then
                                        If MapNpc(MapNum).Npc(x).dir <> DIR_DOWN Then
                                            Call NpcDir(MapNum, x, DIR_DOWN)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    ' We could not move so Target must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
    
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
    
                                            If CanNpcMove(MapNum, x, i) Then
                                                Call NpcMove(MapNum, x, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
    
                            Else
                                i = Int(Rnd * 4)
    
                                If i = 1 Then
                                    i = Int(Rnd * 4)
    
                                    If CanNpcMove(MapNum, x, i) Then
                                        Call NpcMove(MapNum, x, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
                    Target = MapNpc(MapNum).Npc(x).Target
                    TargetType = MapNpc(MapNum).Npc(x).TargetType

                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                    
                        If TargetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = MapNum Then
                                If Not TryNpcAttackPlayer(x, Target) Then Call TryNpcAttackCustom(x, Target, MapNum)
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(MapNum).Npc(x).Target = 0
                                MapNpc(MapNum).Npc(x).TargetType = 0 ' clear
                            End If
                        Else
                            If Not TryNpcAttackNpc(x, Target, MapNum) Then Call TryNpcAttackCustom(x, Target, MapNum)
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(MapNum).Npc(x).stopRegen Then
                    If MapNpc(MapNum).Npc(x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(MapNum).Npc(x).Vital(Vitals.HitPoints) > 0 Then
                            MapNpc(MapNum).Npc(x).Vital(Vitals.HitPoints) = MapNpc(MapNum).Npc(x).Vital(Vitals.HitPoints) + GetNpcVitalRegen(NpcNum, Vitals.HitPoints)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(MapNum).Npc(x).Vital(Vitals.HitPoints) > GetNpcMaxVital(NpcNum, Vitals.HitPoints) Then
                                MapNpc(MapNum).Npc(x).Vital(Vitals.HitPoints) = GetNpcMaxVital(NpcNum, Vitals.HitPoints)
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HitPoints <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(MapNum).Npc(x).Num = 0 And Map(MapNum).Npc(x) > 0 Then
                    If TickCount > MapNpc(MapNum).Npc(x).SpawnWait + (Npc(Map(MapNum).Npc(x)).RespawnRate * 1000) Then
                        Call SpawnNpc(x, MapNum)
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

End Sub

Private Sub UpdatePlayerVitals()
Dim i As Long
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
        
            If TempPlayer(i).Overload > 0 Then
                TempPlayer(i).Overload = TempPlayer(i).Overload - 5000
                
                If TempPlayer(i).Overload > 0 Then
                    Player(i).Skill(Skills.Attack).Level = 125
                    Player(i).Skill(Skills.Strength).Level = 125
                    Player(i).Skill(Skills.Defense).Level = 125
                    Player(i).Skill(Skills.Range).Level = 125
                    Player(i).Skill(Skills.magic).Level = 125
                    SendSkills (i)
                    
                    Select Case TempPlayer(i).Overload
                        Case 120000
                            PlayerMsg i, "The effects of the overload will wear off in 2 minutes.", BrightCyan
                        Case 60000
                            PlayerMsg i, "The effects of the overload will wear off in 1 minute.", BrightCyan
                        Case 30000
                            PlayerMsg i, "The effects of the overload will wear off in 30 seconds.", BrightCyan
                        Case 10000
                            PlayerMsg i, "The effects of the overload will wear off in 10 secons.", BrightCyan
                    End Select
                Else
                    Player(i).Skill(Skills.Attack).Level = Player(i).Skill(Skills.Attack).MaxLevel
                    Player(i).Skill(Skills.Strength).Level = Player(i).Skill(Skills.Strength).MaxLevel
                    Player(i).Skill(Skills.Defense).Level = Player(i).Skill(Skills.Defense).MaxLevel
                    Player(i).Skill(Skills.Range).Level = Player(i).Skill(Skills.Range).MaxLevel
                    Player(i).Skill(Skills.magic).Level = Player(i).Skill(Skills.magic).MaxLevel
                
                    PlayerMsg i, "The effects of the overload have worn off.", BrightRed
                    TempPlayer(i).Overload = 0
                    SendSkills i
                End If
            End If
        
            If Not TempPlayer(i).stopRegen Then
            
                ' TODO
                'If GetPlayerVital(i, Vitals.HitPoints) <> GetPlayerMaxVital(i, Vitals.HitPoints) Then
                    'Call SetPlayerVital(i, Vitals.HitPoints, GetPlayerVital(i, Vitals.HitPoints) + GetPlayerVitalRegen(i, Vitals.HitPoints))
                    'Call SendVital(i, Vitals.HitPoints)
                'End If
                
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")

        For i = 1 To Player_HighIndex

            If IsPlaying(i) Then
                Call SavePlayer(i)
                Call SaveBank(i)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()

    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub
