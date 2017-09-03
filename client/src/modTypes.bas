Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS + MAX_PLAYERS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public MapProjectile As MapProjectileRec

' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public MenuButton(1 To MAX_MENUBUTTONS) As ButtonRec
Public MainButton(1 To MAX_MAINBUTTONS) As ButtonRec
Public Clan As ClanRec

' options
Public Options As OptionsRec

Private Type SkillRec
    MaxLevel As Long
    Level As Long
    XP As Long
End Type

Private Type DataProjectileRec
    TravelTime As Long
    Pic As Long
    Range As Long
    Damage As Long
    Speed As Long
End Type

' projectiles
Public Type ProjectileRec
    TravelTime As Long
    Traveled As Long
    Direction As Long
    x As Long
    y As Long
    Pic As Long
    Range As Long
    Damage As Long
    Speed As Long
End Type

' Type recs
Private Type OptionsRec
    Game_Name As String
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    MenuMusic As String
    Music As Byte
    Sound As Byte
    Debug As Byte
End Type

Private Type ClanMemberRec
    playerIndex As Long
    Rank As String * NAME_LENGTH
End Type

Private Type ClanRec
    Member(1 To MAX_PLAYERS) As ClanMemberRec
    LootShare As Byte
End Type

Public Type PlayerInvRec
    num As Long
    Value As Long
    CustomID As Long
End Type

Private Type BankTabRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type BankRec
    BankTab(1 To MAX_BANK_TABS) As BankTabRec
End Type

Private Type SpellAnim
    spellnum As Long
    Timer As Long
    FramePointer As Long
End Type

Private Type SummoningCreatureRec
    NpcNum As Long
    Health As Long
    Special As Byte
End Type

Private Type PlayerRec
    ' General
    Name As String
    Gender As Byte
    Sprite As Long
    CombatLevel As Byte
    Access As Byte
    ' Stats
    Skill(1 To Skills.Skill_Count - 1) As SkillRec
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As PlayerInvRec
    ' Position
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte

    Spellbook As Byte
    SpecialAttack As Byte
    SummoningCreature As SummoningCreatureRec
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    DirBlock As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * NAME_LENGTH
    LoseItemsOnDeath As Byte
    Instance As Byte
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    Npc(1 To MAX_MAP_NPCS) As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    ItemType As Byte
    MonetaryValue As Long
    Tradable As Byte
    Stackable As Byte
    
    Description As String * 255
    Sound As String * NAME_LENGTH
    AccessRequired As Long
    
    Picture As Long
    Animation As Long
    Paperdoll(1 To Genders.Count - 1) As Long
    
    isTwoHander As Byte
    Damage As Long
    Speed As Long
    EquipmentType As Long
    CombatType As Byte
    SkillBonus(1 To Skills.Skill_Count - 1) As Long
    Offense(1 To CombatStyles.Count - 1) As Long
    Defense(1 To CombatStyles.Count - 1) As Long
    
    AddVital(1 To Vitals.Vital_Count - 1) As Long
   ' projectile
    Projectile As DataProjectileRec
    
    ' Requirement to wear
    SkillWearReq(1 To Skills.Skill_Count - 1) As Long
    SkillMakeRew(1 To Skills.Skill_Count - 1) As Long
    SkillMakeReq(1 To Skills.Skill_Count - 1) As Long
End Type

Private Type MapItemRec
    playerName As String
    num As Long
    Value As Long
    Frame As Byte
    x As Byte
    y As Byte
End Type

Private Type DropRec
    Item As Long
    Amount As Long
    Chance As Byte
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite As Long
    RespawnRate As Long
    Type As Byte
    SightRange As Byte
    
    Drop(1 To MAX_NPC_DROPS) As DropRec

    Skill(1 To Skills.Skill_Count - 1) As Long
    Offense(1 To CombatStyles.Count - 1) As Long
    Defense(1 To CombatStyles.Count - 1) As Long

    RewardXP As Long
    Animation As Long
    Damage As Long
    Level As Long
    AttackSpeed As Long
    Team As String * NAME_LENGTH
    Projectile As DataProjectileRec
    Spell As Long
End Type

Private Type MapNpcRec
    num As Long
    target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    x As Byte
    y As Byte
    Dir As Byte
    ' Client use only
    XOffset As Long
    YOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
End Type

Private Type TradeItemRec
    Item As Long
    CostItem(1 To MAX_SHOP_ITEM_COSTS) As Long
    CostValue(1 To MAX_SHOP_ITEM_COSTS) As Long
    GiveAndRequireXP As Byte
    MaxStock As Long
    Stock As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    BuyVerb As String * NAME_LENGTH
    ShopCurrency As Long
    InterfaceNum As Long
    OnlyBuyItemsInStock As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Description As String * 255
    Sound As String * NAME_LENGTH
    Icon As Long
    CastingTime As Long
    CooldownTime As Long
    Range As Long
    AoE As Byte
    AoERange As Long
    CastAnim As Long
    HitAnim As Long
    StunDuration As Long
    Damage As Long
End Type

Private Type TempTileRec
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
End Type

Public Type MapResourceRec
    x As Long
    y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    SkillType As Long
    
    ResourceImage As Long
    ExhaustedImage As Long
    ToolRequired As Long
    minHealth As Long
    maxHealth As Long
    RespawnTime As Long
    RequireWeapon As Byte
    Walkthrough As Byte
    Animation As Long
    
    RewardItem(1 To 10) As Long
    RewardValue(1 To 10) As Long
    RewardChance(1 To 10) As Long
    SkillReq(1 To Skills.Skill_Count - 1) As Long
    RewardXP(1 To Skills.Skill_Count - 1) As Long
End Type

Private Type ActionMsgRec
    message As String
    Created As Long
    Type As Long
    color As Long
    Scroll As Long
    x As Long
    y As Long
    Timer As Long
End Type

Private Type BloodRec
    Sprite As Long
    Timer As Long
    x As Long
    y As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    x As Long
    y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    Timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    FrameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type ButtonRec
    fileName As String
    state As Byte
End Type

Private Type MapProjectileRec
    Projectile(1 To MAX_PROJECTILES) As ProjectileRec
End Type
