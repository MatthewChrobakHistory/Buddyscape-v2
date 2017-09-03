Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim f As Long

    If ServerLog Then
        filename = App.Path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            f = FreeFile
            Open filename For Output As #f
            Close #f
        End If

        f = FreeFile
        Open filename For Append As #f
        Print #f, Time & ": " & Text
        Close #f
    End If

End Sub

' gets a string from a text file
Public Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(file As String, Header As String, Var As String, value As String)
    Call WritePrivateProfileString$(Header, Var, value, file)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    
End Sub

Public Sub LoadOptions()
    
    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Website")
    
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim f As Long
    Dim i As Long
    filename = App.Path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    f = FreeFile
    Open filename For Append As #f
    Print #f, IP & "," & GetPlayerName(BannedByIndex)
    Close #f
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim f As Long
    Dim i As Long
    filename = App.Path & "data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next

    IP = Mid$(IP, 1, i)
    f = FreeFile
    Open filename For Append As #f
    Print #f, IP & "," & "Server"
    Close #f
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & "the Server" & "!", White)
    Call AddLog("The Server" & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & "The Server" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(filename) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    If AccountExist(Name) Then
        filename = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
    ClearPlayer Index
    
    Player(Index).Login = Name
    Player(Index).Password = Password

    Call SavePlayer(Index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal Index As Long) As Boolean

    If LenB(Trim$(Player(Index).Name)) > 0 Then
        CharExist = True
    End If

End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal Sprite As Long)
    Dim f As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(Player(Index).Name)) = 0 Then
        
        spritecheck = False
        
        Player(Index).Name = Name
        Player(Index).Gender = Sex
        
        If Player(Index).Gender = Male Then
            Player(Index).Sprite = 1
        Else
            Player(Index).Sprite = 2
        End If
        
        Player(Index).CombatLevel = 1
        
        For n = 1 To Skills.Skill_Count - 1
            Player(Index).Skill(n).MaxLevel = 1
            Player(Index).Skill(n).Level = 1
            Player(Index).Skill(n).xp = 0
        Next
        
        Player(Index).Skill(Skills.HitPoints).Level = 10
        Player(Index).Skill(Skills.HitPoints).MaxLevel = 10
        
        LoadPlayerSpellbook (Index)

        Player(Index).dir = DIR_DOWN
        Player(Index).Map = START_MAP
        Player(Index).x = START_X
        Player(Index).y = START_Y
        Player(Index).dir = DIR_DOWN
        
        ' Append name to file
        f = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Append As #f
        Print #f, Name
        Close #f
        Call SavePlayer(Index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim s As String
    f = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim filename As String
    Dim f As Long

    filename = App.Path & "\data\accounts\" & Trim$(Player(Index).Login) & ".bin"
    
    f = FreeFile
    
    Open filename For Binary As #f
    Put #f, , Player(Index)
    Close #f
    
    SavePlayerValue (Index)
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim f As Long
    Call ClearPlayer(Index)
    filename = App.Path & "\data\accounts\" & Trim(Name) & ".bin"
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Player(Index)
    Close #f
    
    LoadPlayerValue (Index)
    
    LoadPlayerSpellbook (Index)
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    Player(Index).Name = vbNullString

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim filename As String
    Dim f  As Long
    filename = App.Path & "\data\items\item" & ItemNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.Path & "\data\Items\Item" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , item(i)
        Close #f
    Next

End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(item(Index)), LenB(item(Index)))
    item(Index).Name = vbNullString
    item(Index).Description = vbNullString
    item(Index).Sound = "None."
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\shops\shop" & ShopNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim x As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.Path & "\data\shops\shop" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Shop(i)
        Close #f
        
        For x = 1 To MAX_TRADES
            If Shop(i).TradeItem(x).item > 0 And Shop(i).TradeItem(x).MaxStock <> -255 Then
                Shop(i).TradeItem(x).Stock = Shop(i).TradeItem(x).MaxStock
            End If
        Next
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal SpellNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\spells\spells" & SpellNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.Path & "\data\spells\spells" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Spell(i)
        Close #f
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Description = vbNullString
    Spell(Index).Sound = "None."
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal NpcNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\npcs\npc" & NpcNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Npc(NpcNum)
    Close #f
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.Path & "\data\npcs\npc" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Npc(i)
        Close #f
    Next

End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Npc(Index).Name = vbNullString
    Npc(Index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Resource(ResourceNum)
    Close #f
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.Path & "\data\resources\resource" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , Resource(i)
        Close #f
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

Sub ClearNpcProjectile(ByVal MapNum As Long, ByVal Index As Long, ByVal NpcProjectile As Long)
    ' clear the projectile
    'With MapNpc(MapNum).Npc(Index).Projectile(NpcProjectile)
    '    .Direction = 0
    '    .Pic = 0
    '    .TravelTime = 0
    '    .x = 0
    '    .y = 0
    '    .Range = 0
    '    .Damage = 0
    '    .Speed = 0
    'End With
    MsgBox ("ClearNpcProjectile was called.")
End Sub

Sub ClearProjectile(ByVal MapIndex As Long, ByVal ProjectileIndex As Long)
    ' clear the projectile
    With MapProjectile(MapIndex).Projectile(ProjectileIndex)
        .Direction = 0
        .Pic = 0
        .TravelTime = 0
        .Traveled = 0
        .x = 0
        .y = 0
        .Range = 0
        .Damage = 0
        .Speed = 0
        .CreatorType = 0
    End With
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim f As Long
    filename = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    f = FreeFile
    Open filename For Binary As #f
        Put #f, , Animation(AnimationNum)
    Close #f
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.Path & "\data\animations\animation" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
            Get #f, , Animation(i)
        Close #f
    Next

End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
    Animation(Index).Sound = "None."
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String
    Dim f As Long
    Dim x As Long
    Dim y As Long
    
    filename = App.Path & "\data\maps\map" & MapNum & ".dat"
    f = FreeFile
    
    Open filename For Binary As #f
    Put #f, , Map(MapNum).Name
    Put #f, , Map(MapNum).Music
    Put #f, , Map(MapNum).Revision
    Put #f, , Map(MapNum).Moral
    Put #f, , Map(MapNum).Up
    Put #f, , Map(MapNum).Down
    Put #f, , Map(MapNum).Left
    Put #f, , Map(MapNum).Right
    Put #f, , Map(MapNum).BootMap
    Put #f, , Map(MapNum).BootX
    Put #f, , Map(MapNum).BootY
    Put #f, , Map(MapNum).MaxX
    Put #f, , Map(MapNum).MaxY
    Put #f, , Map(MapNum).Instance

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            Put #f, , Map(MapNum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , Map(MapNum).Npc(x)
    Next
    Close #f
    
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To (MAX_MAPS)
        Call SaveMap(i)
    Next

End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim f As Long
    Dim x As Long
    Dim y As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        filename = App.Path & "\data\maps\map" & i & ".dat"
        f = FreeFile
        Open filename For Binary As #f
        Get #f, , Map(i).Name
        Get #f, , Map(i).Music
        Get #f, , Map(i).Revision
        Get #f, , Map(i).Moral
        Get #f, , Map(i).Up
        Get #f, , Map(i).Down
        Get #f, , Map(i).Left
        Get #f, , Map(i).Right
        Get #f, , Map(i).BootMap
        Get #f, , Map(i).BootX
        Get #f, , Map(i).BootY
        Get #f, , Map(i).MaxX
        Get #f, , Map(i).MaxY
        Get #f, , Map(i).Instance
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        For x = 0 To Map(i).MaxX
            For y = 0 To Map(i).MaxY
                Get #f, , Map(i).Tile(x, y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #f, , Map(i).Npc(x)
            MapNpc(i).Npc(x).Num = Map(i).Npc(x)
        Next

        Close #f
        
        ClearTempTile i
        CacheResources i
        DoEvents
    Next
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, Index)), LenB(MapItem(MapNum, Index)))
    MapItem(MapNum, Index).Owner = vbNullString
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    ReDim MapNpc(MapNum).Npc(1 To MAX_MAP_NPCS + MAX_PLAYERS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(Index)), LenB(MapNpc(MapNum).Npc(Index)))
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).Name = vbNullString
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Sub SaveBank(ByVal Index As Long)
    Dim filename As String
    Dim f As Long
    
    filename = App.Path & "\data\banks\" & Trim$(Player(Index).Login) & ".bin"
    
    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Bank(Index)
    Close #f
End Sub

Public Sub LoadBank(ByVal Index As Long, ByVal Name As String)
    Dim filename As String
    Dim f As Long

    Call ClearBank(Index)

    filename = App.Path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(Index)
        Exit Sub
    End If

    f = FreeFile
    Open filename For Binary As #f
        Get #f, , Bank(Index)
    Close #f

End Sub

Sub ClearBank(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(Index)), LenB(Bank(Index)))
End Sub

Sub ClearClan(ByVal clanNum As Long)
    Call ZeroMemory(ByVal VarPtr(Clan(clanNum)), LenB(Clan(clanNum)))
End Sub


Public Function GetSkillName(ByVal Skill As Long) As String

    Select Case Skill
        Case Skills.Agility
            GetSkillName = "Agility"
        Case Skills.Attack
            GetSkillName = "Attack"
        Case Skills.Construction
            GetSkillName = "Construction"
        Case Skills.Cooking
            GetSkillName = "Cooking"
        Case Skills.Crafting
            GetSkillName = "Crafting"
        Case Skills.Defense
            GetSkillName = "Defense"
        Case Skills.Dungeoneering
            GetSkillName = "Dungeoneering"
        Case Skills.Farming
            GetSkillName = "Farming"
        Case Skills.Firemaking
            GetSkillName = "Firemaking"
        Case Skills.Fishing
            GetSkillName = "Fishing"
        Case Skills.Fletching
            GetSkillName = "Fletching"
        Case Skills.Herblore
            GetSkillName = "Herblore"
        Case Skills.HitPoints
            GetSkillName = "HitPoints"
        Case Skills.Hunter
            GetSkillName = "Hunter"
        Case Skills.magic
            GetSkillName = "Magic"
        Case Skills.Mining
            GetSkillName = "Mining"
        Case Skills.Prayer
            GetSkillName = "Prayer"
        Case Skills.Range
            GetSkillName = "Range"
        Case Skills.Runecrafting
            GetSkillName = "Runecrafting"
        Case Skills.Slayer
            GetSkillName = "Slayer"
        Case Skills.Smithing
            GetSkillName = "Smithing"
        Case Skills.Strength
            GetSkillName = "Strength"
        Case Skills.Summoning
            GetSkillName = "Summoning"
        Case Skills.Thieving
            GetSkillName = "Thieving"
        Case Skills.Woodcutting
            GetSkillName = "Woodcutting"
    End Select
End Function
