Attribute VB_Name = "modPrayer"
Option Explicit

Public Const MAX_PRAYERS As Byte = 50

Public Prayer(1 To MAX_PRAYERS) As PrayerRec

Private Type PrayerRec
    Name As String * NAME_LENGTH
    Info As String * 255
    LevelRequired As Byte
    DrainRate As Double
End Type

' ********************
' [ NORMAL SPELLBOOK ]
' ********************
Public Const THICK_SKIN As Byte = 1 ' 1
Public Const BURST_OF_STRENGTH As Byte = THICK_SKIN + 1 ' 4
Public Const CLARITY_OF_THOUGHT As Byte = BURST_OF_STRENGTH + 1 ' 7
Public Const SHARP_EYE As Byte = CLARITY_OF_THOUGHT + 1 ' 8
Public Const MYSTIC_WILL As Byte = SHARP_EYE + 1 ' 9
Public Const ROCK_SKIN As Byte = MYSTIC_WILL + 1 ' 10
Public Const SUPERHUMAN_STRENGTH As Byte = ROCK_SKIN + 1 ' 13
Public Const IMPROVED_REFLEXES As Byte = SUPERHUMAN_STRENGTH + 1 ' 16
Public Const RAPID_RESTORE As Byte = IMPROVED_REFLEXES + 1 ' 19
Public Const RAPID_HEAL As Byte = RAPID_RESTORE + 1 ' 22
Public Const PROTECT_ITEM_NORMAL As Byte = RAPID_HEAL + 1 ' 25
Public Const HAWK_EYE As Byte = PROTECT_ITEM_NORMAL + 1 ' 26
Public Const MYSTIC_LORE As Byte = HAWK_EYE + 1 ' 27
Public Const STEEL_SKIN As Byte = MYSTIC_LORE + 1 ' 28
Public Const ULTIMATE_STRENGTH As Byte = STEEL_SKIN + 1 ' 31
Public Const INCREDIBLE_REFLEXES As Byte = ULTIMATE_STRENGTH + 1 ' 34
Public Const PROTECT_FROM_SUMMONING As Byte = INCREDIBLE_REFLEXES + 1 ' 35
Public Const PROTECT_FROM_MAGIC As Byte = PROTECT_FROM_SUMMONING + 1 ' 37
Public Const PROTECT_FROM_MISSILES As Byte = PROTECT_FROM_MAGIC + 1 ' 40
Public Const PROTECT_FROM_MELEE As Byte = PROTECT_FROM_MISSILES + 1 ' 43
Public Const EAGLE_EYE As Byte = PROTECT_FROM_MELEE + 1 ' 44
Public Const MYSTIC_MIGHT As Byte = EAGLE_EYE + 1 ' 45
Public Const RETRIBUTION As Byte = MYSTIC_MIGHT + 1 ' 46
Public Const REDEMPTION As Byte = RETRIBUTION + 1 ' 49
Public Const SMITE As Byte = REDEMPTION + 1 ' 52
Public Const CHIVALRY As Byte = SMITE + 1 ' 60
Public Const RAPID_RENEWAL As Byte = CHIVALRY + 1 ' 65
Public Const PIETY As Byte = RAPID_RENEWAL + 1 ' 70
Public Const RIGOUR As Byte = PIETY + 1 ' 74
Public Const AUGURY As Byte = RIGOUR + 1 ' 77

' **********
' [ CURSES ]
' **********

Public Const PROTECT_ITEM_CURSES As Byte = AUGURY + 1 ' 50
Public Const SAP_WARRIOR As Byte = PROTECT_ITEM_CURSES + 1 ' 50
Public Const SAP_RANGER As Byte = SAP_WARRIOR + 1 ' 52
Public Const SAP_MAGE As Byte = SAP_RANGER + 1 ' 54
Public Const SAP_SPIRIT As Byte = SAP_MAGE + 1 ' 56
Public Const BERSERKER As Byte = SAP_SPIRIT + 1 ' 59
Public Const DEFLECT_SUMMONING As Byte = BERSERKER + 1 ' 62
Public Const DEFLECT_MAGIC As Byte = DEFLECT_SUMMONING + 1 ' 65
Public Const DEFLECT_MISSILES As Byte = DEFLECT_MAGIC + 1 ' 68
Public Const DEFLECT_MELEE As Byte = DEFLECT_MISSILES + 1 ' 71
Public Const LEECH_ATTACK As Byte = DEFLECT_MELEE + 1 ' 74
Public Const LEECH_RANGED As Byte = LEECH_ATTACK + 1 ' 76
Public Const LEECH_MAGIC As Byte = LEECH_RANGED + 1 ' 78
Public Const LEECH_DEFENSE As Byte = LEECH_MAGIC + 1 ' 80
Public Const LEECH_STRENGTH As Byte = LEECH_DEFENSE + 1 ' 82
Public Const LEECH_ENERGY As Byte = LEECH_STRENGTH + 1 ' 84
Public Const LEECH_SPECIAL_ATTACK As Byte = LEECH_ENERGY + 1 ' 86
Public Const WRATH As Byte = LEECH_SPECIAL_ATTACK + 1 ' 89
Public Const SOUL_SPLIT As Byte = WRATH + 1 ' 92
Public Const TURMOIL As Byte = SOUL_SPLIT + 1 ' 95


