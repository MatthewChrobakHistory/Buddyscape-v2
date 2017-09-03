Attribute VB_Name = "modSpells"
Option Explicit

Public Const NORMAL_SPELLBOOK As Byte = 0
Public Const ANCIENT_SPELLBOOK As Byte = 1

Public Const TELEPORT_HOME As Byte = 0
Public Const WIND_STRIKE As Byte = TELEPORT_HOME + 1
Public Const ENCHANT_CROSSBOW As Byte = WIND_STRIKE + 1
Public Const WATER_STRIKE As Byte = ENCHANT_CROSSBOW + 1
Public Const ENCHANT_SAPHIRE As Byte = WATER_STRIKE + 1
Public Const EARTH_STRIKE As Byte = ENCHANT_SAPHIRE + 1
Public Const FIRE_STRIKE As Byte = EARTH_STRIKE + 1
Public Const WIND_BOLT As Byte = FIRE_STRIKE + 1
Public Const BIND As Byte = WIND_BOLT + 1
Public Const LOW_LEVEL_ALCHEMY As Byte = BIND + 1
Public Const WATER_BOLT As Byte = LOW_LEVEL_ALCHEMY + 1
Public Const TELEPORT_VARROCK As Byte = WATER_BOLT + 1
Public Const ENCHANT_EMERALD As Byte = TELEPORT_VARROCK + 1
Public Const EARTH_BOLT As Byte = ENCHANT_EMERALD + 1
Public Const TELEPORT_LUMBRIDGE As Byte = EARTH_BOLT + 1
Public Const FIRE_BOLT As Byte = TELEPORT_LUMBRIDGE + 1
Public Const TELEPORT_FALADOR As Byte = FIRE_BOLT + 1
Public Const TELEPORT_HOUSE As Byte = TELEPORT_FALADOR + 1
Public Const WIND_BLAST As Byte = TELEPORT_HOUSE + 1
Public Const SUPERHEAT As Byte = WIND_BLAST + 1
Public Const TELEPORT_CAMELOT As Byte = SUPERHEAT + 1
Public Const WATER_BLAST As Byte = TELEPORT_CAMELOT + 1
Public Const ENCHANT_RUBY As Byte = WATER_BLAST + 1
Public Const SNARE As Byte = ENCHANT_RUBY + 1
Public Const MAGIC_DART As Byte = SNARE + 1
Public Const TELEPORT_ARDOUGNE As Byte = MAGIC_DART + 1
Public Const EARTH_BLAST As Byte = TELEPORT_ARDOUGNE + 1
Public Const HIGH_LEVEL_ALCHEMY As Byte = EARTH_BLAST + 1
Public Const ENCHANT_DIAMOND As Byte = HIGH_LEVEL_ALCHEMY + 1
Public Const FIRE_BLAST As Byte = ENCHANT_DIAMOND + 1
Public Const CLAWS_OF_GUTHIX As Byte = FIRE_BLAST + 1
Public Const FLAMES_OF_ZAMORAK As Byte = CLAWS_OF_GUTHIX + 1
Public Const SARADOMIN_STRIKE As Byte = FLAMES_OF_ZAMORAK + 1
Public Const WIND_WAVE As Byte = SARADOMIN_STRIKE + 1
Public Const WATER_WAVE As Byte = WIND_WAVE + 1
Public Const ENCHANT_DRAGONSTONE As Byte = WATER_WAVE + 1
Public Const EARTH_WAVE As Byte = ENCHANT_DRAGONSTONE + 1
Public Const TELE_OTHER_LUMB As Byte = EARTH_WAVE + 1
Public Const FIRE_WAVE As Byte = TELE_OTHER_LUMB + 1
Public Const ENTANGLE As Byte = FIRE_WAVE + 1
Public Const WIND_SURGE As Byte = ENTANGLE + 1
Public Const TELE_OTHER_FALADOR As Byte = WIND_SURGE + 1
Public Const TELEPORT_BLOCK As Byte = TELE_OTHER_FALADOR + 1
Public Const WATER_SURGE As Byte = TELEPORT_BLOCK + 1
Public Const ENCHANT_ONYX As Byte = WATER_SURGE + 1
Public Const EARTH_SURGE As Byte = ENCHANT_ONYX + 1
Public Const FIRE_SURGE As Byte = EARTH_SURGE + 1

' Ancients
Public Const SMOKE_RUSH As Byte = FIRE_SURGE + 1
Public Const SHADOW_RUSH As Byte = SMOKE_RUSH + 1
Public Const BLOOD_RUSH As Byte = SHADOW_RUSH + 1
Public Const ICE_RUSH As Byte = BLOOD_RUSH + 1

Public Const SMOKE_BURST As Byte = ICE_RUSH + 1
Public Const SHADOW_BURST As Byte = SMOKE_BURST + 1
Public Const BLOOD_BURST As Byte = SHADOW_BURST + 1
Public Const ICE_BURST As Byte = BLOOD_BURST + 1

Public Const SMOKE_BLITZ As Byte = ICE_BURST + 1
Public Const SHADOW_BLITZ As Byte = SMOKE_BLITZ + 1
Public Const BLOOD_BLITZ As Byte = SHADOW_BLITZ + 1
Public Const ICE_BLITZ As Byte = BLOOD_BLITZ + 1

Public Const SMOKE_BARRAGE As Byte = ICE_BLITZ + 1
Public Const SHADOW_BARRAGE As Byte = SMOKE_BARRAGE + 1
Public Const BLOOD_BARRAGE As Byte = SHADOW_BARRAGE + 1
Public Const ICE_BARRAGE As Byte = BLOOD_BARRAGE + 1

Public Sub ToggleSpellBook(ByVal Index As Long)

    If Player(Index).Spellbook = NORMAL_SPELLBOOK Then
        Player(Index).Spellbook = ANCIENT_SPELLBOOK
    Else
        Player(Index).Spellbook = NORMAL_SPELLBOOK
    End If
    
    Call LoadPlayerSpellbook(Index)
    Call SendPlayerSpells(Index)

End Sub

Public Sub LoadPlayerSpellbook(ByVal Index As Long)
Dim i As Long

Player(Index).Spellbook = 0

    For i = 1 To MAX_PLAYER_SPELLS
        Player(Index).Spell(i) = 0
    Next

    Select Case Player(Index).Spellbook
        Case ANCIENT_SPELLBOOK
            For i = SMOKE_BURST To ICE_BARRAGE
                Player(Index).Spell(i - (SMOKE_BURST - 1)) = i
            Next
        Case Else
            For i = 1 To FIRE_SURGE
                Player(Index).Spell(i) = i
            Next
    End Select
End Sub

Public Function CanPlayerCastSpell(ByVal Index As Long, ByVal SpellNum As Long) As Boolean

    CanPlayerCastSpell = True
End Function

Public Sub TakeRunes(ByVal Index As Long, ByVal SpellNum As Long)

End Sub

Public Sub CastCustomSpell(ByVal Index As Long, ByVal SpellNum As Long)
Dim ItemNum As Long
Dim DidCast As Boolean

    Select Case SpellNum
        Case ENCHANT_CROSSBOW
            
        '///////////////////
        ' ENCHANTING JEWELRY
        '///////////////////
        Case ENCHANT_SAPHIRE
            ItemNum = TempPlayer(Index).UseSpellOnInvSlot
            DidCast = True
        Case ENCHANT_EMERALD
            ItemNum = TempPlayer(Index).UseSpellOnInvSlot
            DidCast = True
        Case ENCHANT_RUBY
            ItemNum = TempPlayer(Index).UseSpellOnInvSlot
            DidCast = True
        Case ENCHANT_DIAMOND
            ItemNum = TempPlayer(Index).UseSpellOnInvSlot
            DidCast = True
        Case ENCHANT_DRAGONSTONE
            ItemNum = TempPlayer(Index).UseSpellOnInvSlot
            DidCast = True
        Case ENCHANT_ONYX
            ItemNum = TempPlayer(Index).UseSpellOnInvSlot
            DidCast = True
            
            
            
        Case LOW_LEVEL_ALCHEMY, HIGH_LEVEL_ALCHEMY
            
        Case SUPERHEAT
            
        Case TELEPORT_BLOCK
            
        Case TELEPORT_HOME, TELEPORT_VARROCK, TELEPORT_LUMBRIDGE, TELEPORT_FALADOR, TELEPORT_HOUSE, TELEPORT_CAMELOT, TELEPORT_ARDOUGNE
            
        Case TELE_OTHER_LUMB, TELE_OTHER_FALADOR
            
    End Select
    
End Sub
