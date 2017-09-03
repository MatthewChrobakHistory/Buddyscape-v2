Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerDamage(ByVal Index As Long) As Long
    GetPlayerDamage = 1000
End Function

Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim x As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case Vitals.HitPoints
            GetNpcMaxVital = Npc(NpcNum).Skill(Skills.HitPoints)
        Case Vitals.Prayer
            GetNpcMaxVital = Npc(NpcNum).Skill(Skills.Prayer)
        Case Vitals.Summoning
            GetNpcMaxVital = Npc(NpcNum).Skill(Skills.Summoning)
    End Select

End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case Vitals.HitPoints
            i = 1
        Case Vitals.Prayer
            i = 1
    End Select
    
    GetNpcVitalRegen = i

End Function

Function GetNpcDamage(ByVal NpcNum As Long) As Long
    ' EOC DONE
    GetNpcDamage = 100
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function GetPlayerDamageBonus(ByVal Index As Long, ByVal MapNpcNum As Long, ByVal CombatType As CombatStyles, ByVal Damage As Long) As Long
Dim NpcNum As Long

    NpcNum = MapNpc(GetPlayerMap(Index)).Npc(MapNpcNum).Num
    
    ' If void, and void matches combat type.
    
    ' Dharrok
    
End Function
Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
    CanPlayerDodge = False
End Function

Public Function CanPlayerBlock(ByVal Index As Long, ByVal CombatType As CombatStyles, ByVal Damage As Long) As Long
    CanPlayerBlock = Damage
End Function

Public Function CanNpcCrit(ByVal NpcNum As Long) As Boolean
    CanNpcCrit = False
End Function

Public Function CanNpcDodge(ByVal NpcNum As Long) As Boolean
    CanNpcDodge = False
End Function

Public Function CanNpcBlock(ByVal NpcNum As Long, ByVal CombatType As CombatStyles, ByVal Damage As Long) As Long
    CanNpcBlock = Damage
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal Index As Long, ByVal MapNpcNum As Long, ByVal Special As Boolean)
Dim blockAmount As Long
Dim NpcNum As Long
Dim MapNum As Long
Dim Damage As Long
Dim CombatStyle As CombatStyles
Dim COA As Long
Dim Offense As Long, Defense As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(Index, MapNpcNum) Then
    
        MapNum = GetPlayerMap(Index)
        NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
        ' projectiles
        If NpcNum < 1 Then Exit Sub
        
        If GetPlayerEquipment(Index, Weapon) > 0 Then
            CombatStyle = item(GetPlayerEquipment(Index, Weapon)).CombatType
        Else
            CombatStyle = CombatStyles.Melee
        End If
        
        Offense = GetPlayerOffense(Index, CombatStyle)
        Defense = (Npc(NpcNum).Skill(Skills.Defense) / 4) + (Npc(NpcNum).Defense(CombatStyle) / 6.5)
        COA = (Player(Index).Skill(Skills.Attack).Level / 4) + (Offense / 6.5)
        
        COA = RAND(1, COA)
        Defense = RAND(1, Defense)
        
        If COA < Defense Then
            SendActionMsg MapNum, "0", Blue, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
            Exit Sub
        End If
        
        ' Can we crit?
        Damage = GetPlayerDamageBonus(Index, MapNpcNum, CombatStyle, Damage)

        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' See if the npc can soak any damage
        Damage = CanNpcBlock(NpcNum, CombatStyle, Damage)
            
        If Damage > 0 Then
            Call PlayerAttackNpc(Index, MapNpcNum, Damage, Special)
        Else
            Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim AttackSpeed As Long
    Dim SlayerLevelRequirement As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HitPoints) <= 0 Then
        Exit Function
    End If
    
    If NpcNum = 0 Then
        CanPlayerAttackNpc = False
        Exit Function
    End If
    
    If Npc(NpcNum).Type = NPC_BEHAVIOUR_FRIENDLY Or Npc(NpcNum).Type = NPC_BEHAVIOUR_SHOPKEEPER Then
        CanPlayerAttackNpc = False
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
    
        ' Slayer Requirements
        SlayerLevelRequirement = 0
        
        Select Case NpcNum
            
        End Select
        
        If GetPlayerSkillLevel(Attacker, Skills.Slayer) < SlayerLevelRequirement Then
            PlayerMsg Attacker, "You need a slayer level of " & SlayerLevelRequirement & " to attack this monster.", BrightRed
            CanPlayerAttackNpc = False
            Exit Function
        End If
    
        ' exit out early
        If IsSpell Then
            CanPlayerAttackNpc = True
            Exit Function
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            AttackSpeed = item(GetPlayerEquipment(Attacker, Weapon)).Speed
        Else
            AttackSpeed = 1000
        End If

        If NpcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + AttackSpeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x + 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(MapNum).Npc(MapNpcNum).x - 1
                    NpcY = MapNpc(MapNum).Npc(MapNpcNum).y
            End Select

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then
                    CanPlayerAttackNpc = True
                End If
            End If
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim TrueKiller As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    Name = Trim$(Npc(NpcNum).Name)
    
    ' Check for weapon
    n = GetPlayerEquipment(Attacker, Weapon)
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount
    MapNpc(MapNum).Npc(MapNpcNum).PlayerDamage(Attacker) = MapNpc(MapNum).Npc(MapNpcNum).PlayerDamage(Attacker) + Damage

    If Damage >= MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HitPoints) Then
    
        SendActionMsg GetPlayerMap(Attacker), "-" & MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HitPoints), BrightRed, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound Attacker, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, item(GetPlayerEquipment(Attacker, Weapon)).Animation, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y)
            End If
        End If
        
        TrueKiller = Attacker
        
        For n = 1 To MAX_PLAYERS
            If IsPlaying(n) Then
                If MapNpc(MapNum).Npc(MapNpcNum).PlayerDamage(n) > MapNpc(MapNum).Npc(MapNpcNum).PlayerDamage(TrueKiller) Then
                    TrueKiller = n
                End If
            End If
        Next
        
        For n = 1 To MAX_NPC_DROPS
            If Npc(NpcNum).Drop(n).item > 0 Then
                If RAND(0, 100) <= Npc(NpcNum).Drop(n).Chance Then
                    Call SpawnItem(Npc(NpcNum).Drop(n).item, Npc(NpcNum).Drop(n).Amount, MapNum, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, Trim$(Player(n).Name))
                End If
            End If
        Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(MapNpcNum).Num = 0
        MapNpc(MapNum).Npc(MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HitPoints) = 0
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong MapNpcNum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = MapNum Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = MapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
        
        Call KilledNpc(Attacker, NpcNum)
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HitPoints) = MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HitPoints) - Damage

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound Attacker, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(MapNpcNum).TargetType = 1 ' player
        MapNpc(MapNum).Npc(MapNpcNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Type = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(MapNpcNum).Num Then
                    MapNpc(MapNum).Npc(i).Target = Attacker
                    MapNpc(MapNum).Npc(i).TargetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).Npc(MapNpcNum).stopRegen = True
        MapNpc(MapNum).Npc(MapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC MapNpcNum, MapNum, SpellNum
        End If
        
        SendMapNpcVitals MapNum, MapNpcNum
    End If

    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = GetTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Function TryNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
Dim MapNum As Long, NpcNum As Long, blockAmount As Long, Damage As Long

    TryNpcAttackPlayer = False

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNpcNum, Index) Then
        MapNum = GetPlayerMap(Index)
        NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(Index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(Index).x * 32), (Player(Index).y * 32)
            Exit Function
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NpcNum)

        If Damage > 0 Then
            Call NpcAttackPlayer(MapNpcNum, Index, Damage)
            TryNpcAttackPlayer = True
        End If
    End If
End Function

Public Function TryNpcAttackNpc(ByVal Attacker As Long, ByVal Victim As Long, ByVal MapNum As Long) As Boolean
Dim NpcNum As Long, Damage As Long

    TryNpcAttackNpc = False

    If CanNpcAttackNpc(Attacker, Victim, MapNum) Then
        If CanNpcDodge(MapNpc(MapNum).Npc(Victim).Num) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, MapNpc(MapNum).Npc(Victim).x * 32, MapNpc(MapNum).Npc(Victim).y * 32
            Exit Function
        End If
        
        Damage = GetNpcDamage(MapNpc(MapNum).Npc(Attacker).Num)
        
        If Damage > 0 Then
            Call NpcAttackNpc(Attacker, Victim, MapNum, Damage)
            TryNpcAttackNpc = True
        End If
    End If


End Function

Function CanNpcAttackNpc(ByVal Attacker As Long, ByVal Victim As Long, ByVal MapNum As Long) As Boolean
    
    If Attacker <= 0 Or Victim <= 0 Then Exit Function
    If MapNpc(MapNum).Npc(Victim).Vital(Vitals.HitPoints) <= 0 Then Exit Function
    If GetTickCount < MapNpc(MapNum).Npc(Attacker).AttackTimer + 1000 Then Exit Function
    
    ' Melee Combat
    
    If MapNpc(MapNum).Npc(Victim).y + 1 = MapNpc(MapNum).Npc(Attacker).y And MapNpc(MapNum).Npc(Victim).x = MapNpc(MapNum).Npc(Attacker).x Then
        CanNpcAttackNpc = True
    End If
    
    If MapNpc(MapNum).Npc(Victim).y - 1 = MapNpc(MapNum).Npc(Attacker).y And MapNpc(MapNum).Npc(Victim).x = MapNpc(MapNum).Npc(Attacker).x Then
        CanNpcAttackNpc = True
    End If
    
    If MapNpc(MapNum).Npc(Victim).y = MapNpc(MapNum).Npc(Attacker).y And MapNpc(MapNum).Npc(Victim).x + 1 = MapNpc(MapNum).Npc(Attacker).x Then
        CanNpcAttackNpc = True
    End If
    
    If MapNpc(MapNum).Npc(Victim).y = MapNpc(MapNum).Npc(Attacker).y And MapNpc(MapNum).Npc(Victim).x - 1 = MapNpc(MapNum).Npc(Attacker).x Then
        CanNpcAttackNpc = True
    End If
End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Index)).Npc(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    MapNum = GetPlayerMap(Index)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HitPoints) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum).Npc(MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NpcNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(Index) = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(Index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(Index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim exp As Long
    Dim MapNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Victim)
    Name = Trim$(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong MapNpcNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(MapNum).Npc(MapNpcNum).AttackTimer = GetTickCount
    MapNpc(MapNum).Npc(MapNpcNum).stopRegen = True
    MapNpc(MapNum).Npc(MapNpcNum).stopRegenTimer = GetTickCount

    If Damage >= Player(Victim).Skill(Skills.HitPoints).Level Then
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Player(Victim).Skill(Skills.HitPoints).Level, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNpcNum).Num
        
        ' kill player
        KillPlayer Victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(MapNum).Npc(MapNpcNum).Target = 0
        MapNpc(MapNum).Npc(MapNpcNum).TargetType = 0
    Else
        ' Player not dead, just do the damage
        Player(Victim).Skill(Skills.HitPoints).Level = Player(Victim).Skill(Skills.HitPoints).Level - Damage
        Call SendSkill(Victim, Skills.HitPoints)
        Call SendAnimation(MapNum, Npc(MapNpc(GetPlayerMap(Victim)).Npc(MapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        
        ' send the sound
        SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNpcNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
    End If

End Sub

Sub NpcAttackNpc(ByVal Attacker As Long, ByVal Victim As Long, ByVal MapNum As Long, ByVal Damage As Long)
Dim Buffer As clsBuffer
Dim NpcNum As Long
Dim TrueKiller As Long
Dim KillerName As String
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcAttack
    Buffer.WriteLong Attacker
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    MapNpc(MapNum).Npc(Attacker).AttackTimer = GetTickCount
    MapNpc(MapNum).Npc(Victim).stopRegen = True
    MapNpc(MapNum).Npc(Victim).stopRegenTimer = GetTickCount
    MapNpc(MapNum).Npc(Attacker).stopRegen = True
    MapNpc(MapNum).Npc(Attacker).stopRegenTimer = GetTickCount
    
    SendBlood MapNum, MapNpc(MapNum).Npc(Victim).x, MapNpc(MapNum).Npc(Victim).y
    
    If Damage >= MapNpc(MapNum).Npc(Victim).Vital(Vitals.HitPoints) Then
        SendActionMsg MapNum, "-" & MapNpc(MapNum).Npc(Victim).Vital(Vitals.HitPoints), BrightRed, 1, MapNpc(MapNum).Npc(Victim).x * 32, MapNpc(MapNum).Npc(Victim).y * 32
        
        NpcNum = MapNpc(MapNum).Npc(Victim).Num
        
        TrueKiller = 0
        KillerName = vbNullString
        
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If TrueKiller = 0 Then
                    If MapNpc(MapNum).Npc(Victim).PlayerDamage(i) > 0 Then
                        TrueKiller = i
                        KillerName = Trim$(Player(i).Name)
                    End If
                Else
                    If MapNpc(MapNum).Npc(Victim).PlayerDamage(i) > MapNpc(MapNum).Npc(Victim).PlayerDamage(TrueKiller) Then
                        TrueKiller = i
                        KillerName = Trim$(Player(i).Name)
                    End If
                End If
            End If
        Next
        
        For i = 1 To MAX_NPC_DROPS
            If Npc(NpcNum).Drop(i).item > 0 Then
                If RAND(0, 100) <= Npc(NpcNum).Drop(i).Chance Then
                    Call SpawnItem(Npc(NpcNum).Drop(i).item, Npc(NpcNum).Drop(i).Amount, MapNum, MapNpc(MapNum).Npc(Victim).x, MapNpc(MapNum).Npc(Victim).y, KillerName)
                End If
            End If
        Next
        
        MapNpc(MapNum).Npc(Victim).Num = 0
        MapNpc(MapNum).Npc(Victim).SpawnWait = GetTickCount
        MapNpc(MapNum).Npc(Victim).Vital(Vitals.HitPoints) = 0
        MapNpc(MapNum).Npc(Victim).Target = 0
        MapNpc(MapNum).Npc(Victim).TargetType = 0
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong SNpcDead
        Buffer.WriteLong Victim
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = MapNum Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        MapNpc(MapNum).Npc(Victim).Vital(Vitals.HitPoints) = MapNpc(MapNum).Npc(Victim).Vital(Vitals.HitPoints) - Damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, MapNpc(MapNum).Npc(Victim).x * 32, MapNpc(MapNum).Npc(Victim).y * 32
        
        If Npc(MapNpc(MapNum).Npc(Victim).Num).Type = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(Victim).Num Then
                    MapNpc(MapNum).Npc(i).Target = Attacker
                    MapNpc(MapNum).Npc(i).TargetType = TARGET_TYPE_NPC
                End If
            Next
        End If
        
        SendMapNpcVitals MapNum, Victim
    End If

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Special As Long)
Dim blockAmount As Long
Dim NpcNum As Long
Dim MapNum As Long
Dim Damage As Long

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(Attacker, Victim) Then
    
        MapNum = GetPlayerMap(Attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(Victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            Exit Sub
        End If
        'If CanPlayerParry(Victim) Then
            'SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
            'Exit Sub
        'End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
    
        If Damage > 0 Then
            Call PlayerAttackPlayer(Attacker, Victim, Damage)
        Else
            Call PlayerMsg(Attacker, "Your attack does nothing.", BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean

    If Not IsSpell And Not IsProjectile Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(Attacker).AttackTimer + item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function

    If Not IsSpell And Not IsProjectile Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
        Exit Function
    End If

    ' Make sure they have more then 0 hp
    If Player(Victim).Skill(Skills.HitPoints).Level <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim exp As Long
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).stopRegen = True
    TempPlayer(Attacker).stopRegenTimer = GetTickCount

    If Damage >= Player(Victim).Skill(Skills.HitPoints).Level Then
        SendActionMsg GetPlayerMap(Victim), "-" & Player(Victim).Skill(Skills.HitPoints).Level, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, SpellNum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If Player(i).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).Target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Player(Victim).Skill(Skills.HitPoints).Level = Player(Victim).Skill(Skills.HitPoints).Level - Damage
        Call SendSkill(Victim, Skills.HitPoints)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, SpellNum
        
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer Victim, SpellNum
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal Index As Long, ByVal spellslot As Long)
Dim MapNum As Long
Dim SpellNum As Long
Dim SpellCastType As Byte
Dim SpellType As Byte
Dim HasBuffered As Boolean
Dim Range As Long
Dim TargetType As Long
Dim Target As Long

    
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    SpellNum = GetPlayerSpell(Index, spellslot)
    MapNum = GetPlayerMap(Index)
    
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If
    
    ' Check for magic level and runes
    If CanPlayerCastSpell(Index, SpellNum) Then
        If Spell(SpellNum).Range > 0 Then
            If Spell(SpellNum).AoE Then
                SpellType = 2
            Else
                SpellType = 3
            End If
        Else
            If Spell(SpellNum).AoE Then
                SpellType = 0
            Else
                SpellType = 1
            End If
        End If
        
        TargetType = TempPlayer(Index).TargetType
        Target = TempPlayer(Index).Target
        Range = Spell(Index).Range
        HasBuffered = False
        
        Select Case SpellType
            Case 0, 1
                HasBuffered = True
            Case 2, 3
                If Target = 0 Then
                    PlayerMsg Index, "You do not have a target.", BrightRed
                    Exit Sub
                End If
                If TargetType = TARGET_TYPE_PLAYER Then
                    If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(Target), GetPlayerY(Target)) Then
                        PlayerMsg Index, "Target not in range.", BrightRed
                        Exit Sub
                    Else
                        HasBuffered = True
                    End If
                ElseIf TargetType = TARGET_TYPE_NPC Then
                    If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(MapNum).Npc(Target).x, MapNpc(MapNum).Npc(Target).y) Then
                        PlayerMsg Index, "Target not in range.", BrightRed
                        Exit Sub
                    Else
                        HasBuffered = True
                    End If
                End If
        End Select
        
        If HasBuffered Then
            SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index
            SendActionMsg MapNum, "Casting " & Trim$(Spell(SpellNum).Name) & "!", White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
            TempPlayer(Index).spellBuffer.Spell = spellslot
            TempPlayer(Index).spellBuffer.Timer = GetTickCount
            TempPlayer(Index).spellBuffer.Target = TempPlayer(Index).Target
            TempPlayer(Index).spellBuffer.tType = TempPlayer(Index).TargetType
            Exit Sub
        Else
            SendClearSpellBuffer Index
        End If
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal Target As Long, ByVal TargetType As Byte)
Dim SpellNum As Long
Dim Buffer As clsBuffer
Dim DidCast As Boolean
Dim x As Long, y As Long
Dim MapNum As Long
Dim i As Long
Dim AoE As Long

    DidCast = False
    MapNum = GetPlayerMap(Index)
    
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    SpellNum = GetPlayerSpell(Index, spellslot)
    
    Select Case SpellNum
        Case ENCHANT_CROSSBOW
            CastCustomSpell Index, SpellNum
            
        Case ENCHANT_SAPHIRE, ENCHANT_EMERALD, ENCHANT_RUBY, ENCHANT_DIAMOND, ENCHANT_DRAGONSTONE, ENCHANT_ONYX
            CastCustomSpell Index, SpellNum
            
        Case LOW_LEVEL_ALCHEMY, HIGH_LEVEL_ALCHEMY
            CastCustomSpell Index, SpellNum
            
        Case SUPERHEAT
            CastCustomSpell Index, SpellNum
            
        Case TELEPORT_BLOCK
            CastCustomSpell Index, SpellNum
            
        Case TELEPORT_HOME, TELEPORT_VARROCK, TELEPORT_LUMBRIDGE, TELEPORT_FALADOR, TELEPORT_HOUSE, TELEPORT_CAMELOT, TELEPORT_ARDOUGNE
            CastCustomSpell Index, SpellNum
            
        Case TELE_OTHER_LUMB, TELE_OTHER_FALADOR
            CastCustomSpell Index, SpellNum
            
        ' Combat spells left
        Case Else
            If TargetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
            AoE = Spell(SpellNum).AoE
            
            If Target = Index And TargetType = TARGET_TYPE_PLAYER Then
                PlayerMsg Index, "Why would I attack myself?...", BrightRed
                Exit Sub
            End If
            
            If TargetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            ElseIf TargetType = TARGET_TYPE_NPC Then
                x = MapNpc(MapNum).Npc(Target).x
                y = MapNpc(MapNum).Npc(Target).y
            End If
            
            If Not isInRange(Range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                SendClearSpellBuffer Index
                Exit Sub
            End If
            
            If Spell(SpellNum).AoE Then
                DidCast = True
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) And i <> Index Then
                        If GetPlayerMap(Index) = GetPlayerMap(i) Then
                            If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                If CanPlayerAttackPlayer(Index, i, True) Then
                                    SendAnimation MapNum, Spell(SpellNum).HitAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                    PlayerAttackPlayer Index, i, 0, SpellNum
                                End If
                            End If
                        End If
                    End If
                Next
                For i = 1 To MAX_MAP_NPCS
                    If MapNpc(MapNum).Npc(i).Num > 0 Then
                        If MapNpc(MapNum).Npc(i).Vital(Vitals.HitPoints) > 0 Then
                            If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                If CanPlayerAttackNpc(Index, i, True) Then
                                    SendAnimation MapNum, Spell(SpellNum).HitAnim, 0, 0, TARGET_TYPE_NPC, i
                                    PlayerAttackNpc Index, i, 0, SpellNum
                                End If
                            End If
                        End If
                    End If
                Next
            Else
                If TargetType = TARGET_TYPE_PLAYER Then
                    If CanPlayerAttackPlayer(Index, Target, True) Then
                        SendAnimation MapNum, Spell(SpellNum).HitAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                        PlayerAttackPlayer Index, Target, 0, SpellNum
                        DidCast = True
                    End If
                ElseIf TargetType = TARGET_TYPE_NPC Then
                    If CanPlayerAttackNpc(Index, Target, True) Then
                        SendAnimation MapNum, Spell(SpellNum).HitAnim, 0, 0, TARGET_TYPE_NPC, Target
                        PlayerAttackNpc Index, Target, 0, SpellNum
                        DidCast = True
                    End If
                End If
            End If
    End Select
    
    If DidCast Then
        Call TakeRunes(Index, SpellNum)
        TempPlayer(Index).SpellCD(spellslot) = GetTickCount + (Spell(SpellNum).CooldownTime * 1000)
        Call SendCooldown(Index, spellslot)
        SendActionMsg MapNum, Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    'If Damage > 0 Then
    '    If increment Then
    '        sSymbol = "+"
    '        If Vital = Vitals.HitPoints Then Colour = BrightGreen
    '        If Vital = Vitals.Prayer Then Colour = BrightBlue
    '    Else
    '        sSymbol = "-"
    '        Colour = Blue
    '    End If
   '
   '     SendAnimation GetPlayerMap(Index), Spell(SpellNum).HitAnim, 0, 0, TARGET_TYPE_PLAYER, Index
   '     SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
   '
   '     ' send the sound
   '     SendMapSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
   '
   '     If increment Then
   '         SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
   '         If Spell(SpellNum).Duration > 0 Then
    '            AddHoT_Player Index, SpellNum
   '         End If
   '     ElseIf Not increment Then
   '         SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
   '     End If
   ' End If
   MsgBox ("SpellPlayer_Effect was called.")
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal MapNum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    'If Damage > 0 Then
    '    If increment Then
    '        sSymbol = "+"
    '        If Vital = Vitals.HitPoints Then Colour = BrightGreen
    '        If Vital = Vitals.Prayer Then Colour = BrightBlue
    '    Else
    '        sSymbol = "-"
    '        Colour = Blue
    '    End If
   '
   '     SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
   '     SendActionMsg MapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(Index).x * 32, MapNpc(MapNum).Npc(Index).y * 32
   '
   '     ' send the sound
   '     SendMapSound Index, MapNpc(MapNum).Npc(Index).x, MapNpc(MapNum).Npc(Index).y, SoundEntity.seSpell, SpellNum
   '
   '     If increment Then
   '         MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) + Damage
   '         If Spell(SpellNum).Duration > 0 Then
   '             AddHoT_Npc MapNum, Index, SpellNum
   '         End If
   '     ElseIf Not increment Then
   '         MapNpc(MapNum).Npc(Index).Vital(Vital) = MapNpc(MapNum).Npc(Index).Vital(Vital) - Damage
   '     End If
   ' End If
   MsgBox ("SpellNpc_Effect was called.")
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal SpellNum As Long)
    
    If TempPlayer(Index).CanBeStunned > 0 Then
        Exit Sub
    End If

    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(Index).StunTimer = GetTickCount
        TempPlayer(Index).CanBeStunned = GetTickCount + (TempPlayer(Index).StunDuration * 2)
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal MapNum As Long, ByVal SpellNum As Long)
     
    If MapNpc(MapNum).Npc(Index).CanBeStunned > 0 Then
        Exit Sub
    End If
    
    ' check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(MapNum).Npc(Index).StunDuration = Spell(SpellNum).StunDuration
        MapNpc(MapNum).Npc(Index).StunTimer = GetTickCount
        MapNpc(MapNum).Npc(Index).CanBeStunned = GetTickCount + (Spell(SpellNum).StunDuration * 2)
    End If
End Sub
