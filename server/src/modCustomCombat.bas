Attribute VB_Name = "modCustomCombat"
Option Explicit

Public Function IsClearPath(ByVal MapNum As Long, ByVal xStart As Long, ByVal xFinish As Long, ByVal yStart As Long, ByVal yFinish As Long) As Boolean
Dim x As Long
Dim y As Long
Dim dir As Byte

    IsClearPath = False
    
    With Map(MapNum)

        If xStart < 0 Then xStart = 0
        If yStart < 0 Then yStart = 0
        If xFinish > .MaxX Then xFinish = .MaxX
        If yFinish > .MaxY Then yFinish = .MaxY
        
        If xStart <> xFinish And yStart <> yFinish Then Exit Function
        
        If xStart < xFinish Then dir = DIR_RIGHT
        If xStart > xFinish Then dir = DIR_LEFT
        If yStart < yFinish Then dir = DIR_DOWN
        If yStart > yFinish Then dir = DIR_UP

        For x = xStart To xFinish
            For y = yStart To yFinish
                With .Tile(x, y)
                    If isDirBlocked(.DirBlock, dir) Then Exit Function
                    If .Type = TILE_TYPE_BLOCKED Or .Type = TILE_TYPE_OBJECT Then Exit Function
                End With
            Next
        Next
        
    End With
    
    IsClearPath = True

End Function

Public Sub CreateProjectile(ByVal Index As Long, ByVal Owner As Byte, Optional ByVal MapNum As Long)
Dim i As Long
Dim dir As Byte
Dim x As Long
Dim y As Long
    
    If MapNum = 0 Then Exit Sub

    For i = 1 To MAX_PROJECTILES
        With MapProjectile(MapNum).Projectile(i)
            If .CreatorType = 0 Then
                Call ClearProjectile(MapNum, i)
                
                If Owner = 1 Then
                    If GetPlayerEquipment(Index, Weapon) > 0 Then
                        If item(GetPlayerEquipment(Index, Weapon)).Projectile.Pic > 0 Then
                            .Damage = item(GetPlayerEquipment(Index, Weapon)).Projectile.Damage
                            .Pic = item(GetPlayerEquipment(Index, Weapon)).Projectile.Pic
                            .Range = item(GetPlayerEquipment(Index, Weapon)).Projectile.Range
                            .Speed = item(GetPlayerEquipment(Index, Weapon)).Projectile.Speed
                        
                            .Direction = Player(Index).dir
                            .x = Player(Index).x
                            .y = Player(Index).y
                            .CreatorType = Owner
                            .Creator = Index
                            
                            SendProjectileToMap MapNum, i
                            
                            Npc(1).Projectile.Pic = 1
                            Npc(1).Projectile.Range = 10
                            Npc(1).Projectile.Speed = 25
                            Npc(2).Projectile.Pic = 1
                            Npc(2).Projectile.Range = 10
                            Npc(2).Projectile.Speed = 25
                            Npc(3).Projectile.Pic = 1
                            Npc(3).Projectile.Range = 10
                            Npc(3).Projectile.Speed = 25
                            
                            Npc(4).Projectile.Pic = 1
                            Npc(4).Projectile.Range = 10
                            Npc(4).Projectile.Speed = 25
                            Exit For
                        End If
                    End If
                Else
                    Dim NpcNum As Long
                    NpcNum = MapNpc(MapNum).Npc(Index).Num
                    
                    .Damage = Npc(NpcNum).Projectile.Damage
                    .Pic = Npc(NpcNum).Projectile.Pic
                    .Range = Npc(NpcNum).Projectile.Range
                    .Speed = Npc(NpcNum).Projectile.Speed
                    
                    .Direction = MapNpc(MapNum).Npc(Index).dir
                    .x = MapNpc(MapNum).Npc(Index).x
                    .y = MapNpc(MapNum).Npc(Index).y
                    
                    .CreatorType = Owner
                    .Creator = Index
                    
                    SendProjectileToMap MapNum, i
                    Exit For
                End If
            End If
        End With
    Next
End Sub

Public Sub TryNpcAttackCustom(ByVal Attacker, ByVal Victim As Long, ByVal MapNum As Long)
Dim NpcNum As Long
Dim x As Long
Dim y As Long
Dim Continue As Boolean
Dim tick As Long

    NpcNum = MapNpc(MapNum).Npc(Attacker).Num
    If GetTickCount < MapNpc(MapNum).Npc(Attacker).AttackTimer + 1000 Then Exit Sub
    MapNpc(MapNum).Npc(Attacker).AttackTimer = GetTickCount
    
    If MapNpc(MapNum).Npc(Attacker).TargetType = TARGET_TYPE_PLAYER Then
        x = GetPlayerX(Victim)
        y = GetPlayerY(Victim)
    Else
        x = MapNpc(MapNum).Npc(Victim).x
        y = MapNpc(MapNum).Npc(Victim).y
    End If
    
    With MapNpc(MapNum).Npc(Attacker)
        ' Try projectiles.
        If Npc(NpcNum).Projectile.Pic > 0 And Npc(NpcNum).Projectile.Range > 0 And Npc(NpcNum).Projectile.Speed > 0 Then
            Continue = True
        
            If .x - x < 0 And .dir <> DIR_RIGHT Then Continue = False
            If .x - x > 0 And .dir <> DIR_LEFT Then Continue = False
            If .y - y < 0 And .dir <> DIR_DOWN Then Continue = False
            If .y - y > 0 And .dir <> DIR_UP Then Continue = False
            
            If .x = x And Abs(.y - y) > Npc(NpcNum).Projectile.Range Then Continue = False
            If .y = y And Abs(.x - x) > Npc(NpcNum).Projectile.Range Then Continue = False
            
            If Continue Then If IsClearPath(MapNum, .x, x, .y, y) Then Call CreateProjectile(Attacker, 2, MapNum)
        End If
        
        ' Then spells, if ranged didn't work.
        If Npc(NpcNum).Spell > 0 Then
        
        End If
    End With
End Sub

Public Sub HandleProjecTile(ByVal MapNum As Long, ByVal Index As Long)
Dim i As Long
Dim Damage As Long

    If Index > MAX_PROJECTILES Or Index < 1 Then Exit Sub
    
    With MapProjectile(MapNum).Projectile(Index)
        If .Creator = 0 And .CreatorType = 0 And .Damage = 0 And .Pic = 0 Then Exit Sub
    
        If GetTickCount > .TravelTime Then
            Select Case .Direction
                Case DIR_DOWN
                    .y = .y + 1
                    If .y > Map(MapNum).MaxY Then ClearProjectile MapNum, Index: Exit Sub
                Case DIR_UP
                    .y = .y - 1
                    If .y < 0 Then ClearProjectile MapNum, Index: Exit Sub
                Case DIR_RIGHT
                    .x = .x + 1
                    If .x > Map(MapNum).MaxX Then ClearProjectile MapNum, Index: Exit Sub
                Case DIR_LEFT
                    .x = .x - 1
                    If .x < 0 Then ClearProjectile MapNum, Index: Exit Sub
            End Select
            .Traveled = .Traveled + 1
            .TravelTime = GetTickCount + .Speed
            
            If .Traveled > .Range Then ClearProjectile MapNum, Index: Exit Sub
        End If
        
        ' Did it hit a block?
        If Map(MapNum).Tile(.x, .y).Type = TILE_TYPE_OBJECT Then ClearProjectile MapNum, Index: Exit Sub
        If Map(MapNum).Tile(.x, .y).Type = TILE_TYPE_BLOCKED Then ClearProjectile MapNum, Index: Exit Sub
        
        ' Did an npc shoot it, or a player on a pvp map?
        ' If so, check if it hit a player
        If Map(MapNum).Moral <> MAP_MORAL_SAFE Or .CreatorType = 1 Then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    If .x = GetPlayerX(i) And .y = GetPlayerY(i) Then
                        If .CreatorType = 1 Then
                            Call PlayerAttackPlayer(.Creator, i, Damage)
                        Else
                            Call NpcAttackPlayer(.Creator, i, 500)
                        End If
                        Call ClearProjectile(MapNum, Index)
                        Exit Sub
                    End If
                End If
            Next
        End If
        
        If .CreatorType = 2 Then
            If MapNpc(MapNum).Npc(.Creator).Num <> 0 Then
                If Npc(MapNpc(MapNum).Npc(.Creator).Num).Team = vbNullString Then Exit Sub
            End If
        End If
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(MapNum).Npc(i).Num > 0 Then
                If .x = MapNpc(MapNum).Npc(i).x And .y = MapNpc(MapNum).Npc(i).y Then
                    If .CreatorType = 1 Then
                        Call PlayerAttackNpc(.Creator, i, 500)
                    Else
                        Call NpcAttackNpc(.Creator, i, MapNum, 500)
                    End If
                    Call ClearProjectile(MapNum, Index)
                    Exit Sub
                End If
            End If
        Next
    End With
    

End Sub

