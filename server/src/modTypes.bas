Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public MapCache(1 To MAX_MAPS) As Cache
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Bank(1 To MAX_PLAYERS) As BankRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Clan(1 To MAX_PLAYERS) As ClanRec
Public Options As OptionsRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public MapProjectile(1 To MAX_MAPS) As MapProjectileRec

Private Type SkillRec
    MaxLevel As Long
    Level As Long
    xp As Long
End Type

Private Type ClanMemberRec
    PlayerIndex As Long
    Rank As String * NAME_LENGTH
End Type

Private Type ClanRec
    Member(1 To MAX_PLAYERS) As ClanMemberRec
    LootShare As Byte
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
End Type

Public Type PlayerInvRec
    Num As Long
    value As Long
    CustomID As Long
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type BankTabRec
    item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type BankRec
    BankTab(1 To MAX_BANK_TABS) As BankTabRec
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Private Type DataProjectileRec
    TravelTime As Long
    Pic As Long
    Range As Long
    Damage As Long
    Speed As Long
End Type

' project tiles
Private Type ProjectileRec
    TravelTime As Long
    Traveled As Long
    Direction As Long
    x As Long
    y As Long
    Pic As Long
    Range As Long
    Damage As Long
    Speed As Long
    CreatorType As Byte
    Creator As Long
End Type

Private Type SummoningCreatureRec
    NpcNum As Long
    MapNpcNum As Long
    Health As Long
    Special As Byte
End Type

Private Type PlayerRec
    ' Account
    Login As String * ACCOUNT_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * ACCOUNT_LENGTH
    Gender As Byte
    Sprite As Long
    CombatLevel As Byte
    Access As Byte
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    
    ' Stats
    Skill(1 To Skills.Skill_Count - 1) As SkillRec
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As PlayerInvRec
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Position
    Map As Long
    x As Byte
    y As Byte
    dir As Byte
    Spellbook As Byte
    SpecialAttack As Byte
    SummoningCreature As SummoningCreatureRec
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    Target As Long
    tType As Byte
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    TargetType As Byte
    Target As Long
    GettingMap As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    ' trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerInvRec
    AcceptTrade As Boolean
    ' spell buffer
    spellBuffer As SpellBufferRec
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    ' clan
    inClan As Long
    UseItemTimer As Long
    UseSpellOnInvSlot As Long
    CanBeStunned As Long
    Overload As Long
End Type

Private Type TileDataRec
    x As Long
    y As Long
    Tileset As Long
End Type

Private Type TileRec
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
    Num As Long
    value As Long
    x As Byte
    y As Byte
    ' ownership + despawn
    Owner As String
    tmr As Long
    state As Byte
End Type

Private Type DropRec
    item As Long
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
    Num As Long
    Target As Long
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    PlayerDamage(1 To MAX_PLAYERS) As Long
    x As Byte
    y As Byte
    dir As Byte
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    ' regen
    stopRegen As Boolean
    stopRegenTimer As Long
    CanBeStunned As Long
End Type

Private Type TradeItemRec
    item As Long
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

Private Type TempTileRec
    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type MapDataRec
    Npc() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
    cur_health As Long
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
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

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
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

Private Type MapProjectileRec
    Projectile(1 To MAX_PROJECTILES) As ProjectileRec
End Type
