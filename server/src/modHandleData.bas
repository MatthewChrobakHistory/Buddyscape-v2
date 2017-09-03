Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CProjecTileAttack) = GetAddress(AddressOf HandleProjecTileAttack)
    HandleDataSub(CMoveItemToNewBankTab) = GetAddress(AddressOf HandleMoveItemToNewBankTab)
    HandleDataSub(CProgressConversation) = GetAddress(AddressOf HandleProgressConversation)
    HandleDataSub(CRequestJoinClan) = GetAddress(AddressOf HandleRequestJoinClan)
    HandleDataSub(CLeaveClan) = GetAddress(AddressOf HandleLeaveClan)
End Sub

Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long
        
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), Index, Buffer.ReadBytes(Buffer.Length), 0, 0
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Then
                Call AlertMsg(Index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(Index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(Player(Index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar Index
                Else
                    ' send new char shit
                    If Not IsPlaying(Index) Then
                        Call SendCreateChar(Index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(Index, Name)

            If LenB(Trim$(Player(Index).Name)) > 0 Then
                Call DeleteName(Player(Index).Name)
            End If

            Call ClearPlayer(Index)
            ' Everything went ok
            Call Kill(App.Path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(Index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(Index, "Multiple account logins is not authorized.")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(Index, Name)
            ClearBank Index
            LoadBank Index, Name
            
            ' Check if character data has been created
            If LenB(Trim$(Player(Index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar Index
            Else
                ' send new char shit
                If Not IsPlaying(Index) Then
                    Call SendCreateChar(Index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            
            Set Buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Sprite As Long
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Sprite = Buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(Index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < 1) Or (Sex > 2) Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(Index) Then
            Call AlertMsg(Index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(Index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Sprite)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        HandleUseChar Index
        
        Set Buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(Index), Index, Msg, QBColor(White))
    
    Set Buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    s = "[Global]" & GetPlayerName(Index) & ": " & Msg
    Call SayMsg_Global(Index, Msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
            Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(GetPlayerName(Index), "Cannot message yourself.", BrightRed)
    End If
    
    Set Buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    dir = Buffer.ReadLong 'CLng(Parse(1))
    movement = Buffer.ReadLong 'CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Set Buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(Index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(Index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(Index).StunDuration > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(Index).InShop > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(Index) <> tmpX Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    If GetPlayerY(Index) <> tmpY Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    Call PlayerMove(Index, dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerDir
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerDir(Index)
    SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long
Dim Buffer As clsBuffer
    
    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong
    Set Buffer = Nothing

    UseItem Index, invNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long
    Dim Buffer As clsBuffer
    Dim Special As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Special = Buffer.ReadLong
    Set Buffer = Nothing
    
    ' can't attack whilst casting
    If TempPlayer(Index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    'SendAttack Index

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then
            TryPlayerAttackPlayer Index, i, Special
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc Index, i, Special
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = Map(GetPlayerMap(Index)).MaxY Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = Map(GetPlayerMap(Index)).MaxX Then Exit Sub
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index)
    End Select
    
    CheckResource Index, x, y
    
    CheckObject Index, x, y
End Sub
' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString 'Parse(1)
    Set Buffer = Nothing
    i = FindPlayer(Name)
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(Index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp to yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
            Call PlayerMsg(Index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index), True)
    Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The sprite
    n = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    Call SetPlayerSprite(Index, n)
    Call SendPlayerData(Index)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    dir = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If dir < DIR_UP Or dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(Index, dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim MapNum As Long
    Dim x As Long
    Dim y As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(Index)
    i = Map(MapNum).Revision + 1
    Call ClearMap(MapNum)
    
    Map(MapNum).Name = Buffer.ReadString
    Map(MapNum).Music = Buffer.ReadString
    Map(MapNum).Revision = i
    Map(MapNum).Moral = Buffer.ReadByte
    Map(MapNum).Up = Buffer.ReadLong
    Map(MapNum).Down = Buffer.ReadLong
    Map(MapNum).Left = Buffer.ReadLong
    Map(MapNum).Right = Buffer.ReadLong
    Map(MapNum).BootMap = Buffer.ReadLong
    Map(MapNum).BootX = Buffer.ReadByte
    Map(MapNum).BootY = Buffer.ReadByte
    Map(MapNum).MaxX = Buffer.ReadByte
    Map(MapNum).MaxY = Buffer.ReadByte
    Map(MapNum).Instance = Buffer.ReadByte
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(x, y).Layer(i).x = Buffer.ReadLong
                Map(MapNum).Tile(x, y).Layer(i).y = Buffer.ReadLong
                Map(MapNum).Tile(x, y).Layer(i).Tileset = Buffer.ReadLong
            Next
            Map(MapNum).Tile(x, y).Type = Buffer.ReadByte
            Map(MapNum).Tile(x, y).Data1 = Buffer.ReadLong
            Map(MapNum).Tile(x, y).Data2 = Buffer.ReadLong
            Map(MapNum).Tile(x, y).Data3 = Buffer.ReadLong
            Map(MapNum).Tile(x, y).DirBlock = Buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(x) = Buffer.ReadLong
        Call ClearMapNpc(x, MapNum)
    Next

    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    ' Save the map
    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    Call ClearTempTile(MapNum)
    Call CacheResources(MapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i

    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Get yes/no value
    s = Buffer.ReadLong 'Parse(1)
    Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)

    'send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next

    TempPlayer(Index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapDone
    SendDataTo Index, Buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(Index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim Amount As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong 'CLng(Parse(1))
    Amount = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing
    
    If TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(Index, invNum) < 1 Or GetPlayerInvItemNum(Index, invNum) > MAX_ITEMS Then Exit Sub
    
    If item(GetPlayerInvItemNum(Index, invNum)).Stackable = 1 Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(Index, invNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(Index, invNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next

    CacheResources GetPlayerMap(Index)
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If LenB(Trim$(Map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If

    Next

    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(Index, s, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(Index) & "!", White)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim f As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    f = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s
        Input #f, Name
        Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop

    Close #f
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim file As Long
    Dim f As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.Path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open filename For Output As #f
        Close #f
    End If

    Kill filename
    Call PlayerMsg(Index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    Set Buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEditMap
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SItemEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimationEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(Index) & " saved Animation #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SNpcEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim NpcNum As Long
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    NpcNum = Buffer.ReadLong

    ' Prevent hacking
    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = Buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(Npc(NpcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(NpcNum)
    Call SaveNpc(NpcNum)
    Call AddLog(GetPlayerName(Index) & " saved Npc #" & NpcNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SResourceEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(Index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SShopEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ShopNum As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    ShopNum = Buffer.ReadLong

    ' Prevent hacking
    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set Buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellEditor
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    SpellNum = Buffer.ReadLong

    ' Prevent hacking
    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call AddLog(GetPlayerName(Index) & " saved Spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString) 'Parse(1))
    ' The access
    i = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Invalid access level.", Red)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString) 'Parse(1))
    SaveOptions
    Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    x = Buffer.ReadLong 'CLng(Parse(1))
    y = Buffer.ReadLong 'CLng(Parse(2))
    Set Buffer = Nothing

    ' Prevent subscript out of range
    If x < 0 Or x > Map(GetPlayerMap(Index)).MaxX Or y < 0 Or y > Map(GetPlayerMap(Index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        ' Change target
                        If TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER And TempPlayer(Index).Target = i Then
                            TempPlayer(Index).Target = 0
                            TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                            ' send target to player
                            SendTarget Index
                        Else
                            TempPlayer(Index).Target = i
                            TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER
                            ' send target to player
                            SendTarget Index
                        End If
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next

    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(Index)).Npc(i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).Npc(i).x = x Then
                If MapNpc(GetPlayerMap(Index)).Npc(i).y = y Then
                    If TempPlayer(Index).Target = i And TempPlayer(Index).TargetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(Index).Target = 0
                        TempPlayer(Index).TargetType = TARGET_TYPE_NONE
                        ' send target to player
                        SendTarget Index
                    Else
                        ' Change target
                        TempPlayer(Index).Target = i
                        TempPlayer(Index).TargetType = TARGET_TYPE_NPC
                        ' send target to player
                        SendTarget Index
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong 'CLng(Parse(1))
    TempPlayer(Index).UseSpellOnInvSlot = Buffer.ReadLong
    Set Buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(Index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchInvSlots Index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    If TempPlayer(Index).InTrade > 0 Or TempPlayer(Index).InBank Or TempPlayer(Index).InShop Then Exit Sub
    
    If TempPlayer(Index).spellBuffer.Spell > 0 Then
        PlayerMsg Index, "You cannot swap spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(Index).SpellCD(n) > GetTickCount Then
            PlayerMsg Index, "You cannot swap spells whilst they're cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchSpellSlots Index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendPing
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleUnequip(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerUnequipItem Index, Buffer.ReadLong
    Set Buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData Index
End Sub

Sub HandleRequestItems(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems Index
End Sub

Sub HandleRequestAnimations(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations Index
End Sub

Sub HandleRequestNPCS(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs Index
End Sub

Sub HandleRequestResources(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources Index
End Sub

Sub HandleRequestSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells Index
End Sub

Sub HandleRequestShops(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops Index
End Sub

Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' item
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
        
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index)
    Set Buffer = Nothing
End Sub

Sub HandleForgetSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim spellslot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellslot = Buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(Index).spellBuffer.Spell = spellslot Then
        PlayerMsg Index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    Player(Index).Spell(spellslot) = 0
    SendPlayerSpells Index
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(Index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim ShopSlot As Long
    Dim ShopNum As Long
    Dim Amount As Long
    Dim ItemAmount(1 To MAX_SHOP_ITEM_COSTS) As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ShopSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    ' not in shop, exit out
    ShopNum = TempPlayer(Index).InShop
    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then Exit Sub

    With Shop(ShopNum).TradeItem(ShopSlot)
        ' check trade exists
        If .item < 1 Or .item > MAX_ITEMS Then Exit Sub
        
        If .Stock <= 0 And .MaxStock <> -255 Then
            PlayerMsg Index, "The shop does not have any more of this item!", BrightRed
            Exit Sub
        End If
        
        If .MaxStock <> -255 Then
            If .Stock < Amount Then
                Amount = .Stock
            End If
        End If
        
        If item(.item).Stackable = 1 Then
            If BuyItem(Index, ShopNum, ShopSlot, Amount) Then
                PlayerMsg Index, "Trade successful.", BrightGreen
            End If
        Else
            For i = 1 To Amount
                If BuyItem(Index, ShopNum, ShopSlot, 1) = False Then
                    Exit For
                End If
            Next
            PlayerMsg Index, "Trade successful.", BrightGreen
        End If
        
        If .MaxStock <> -255 Then
            .Stock = .Stock - Amount
            Call UpdateShopStock(ShopNum, ShopSlot)
        End If
        
                
        If .Stock < 0 Then
            PlayerMsg Index, "Stock is somehow less than 0.", BrightRed
        End If
        
    End With
    

    
    Set Buffer = Nothing
End Sub

Sub HandleSellItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    Dim ItemNum As Long
    Dim CanTrade As Boolean
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' Do they have the item?
    ItemNum = GetPlayerInvItemNum(Index, invSlot)
    CanTrade = False
    
    If ItemNum = 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' Does the shop want it at all?
    If Shop(TempPlayer(Index).InShop).OnlyBuyItemsInStock Then
        For i = 1 To MAX_TRADES
            If Shop(TempPlayer(Index).InShop).TradeItem(i).item = ItemNum Then
                CanTrade = True
            End If
        Next
    Else
        CanTrade = True
    End If
    
    If Shop(TempPlayer(Index).InShop).BuyRate <= 0 Then
        PlayerMsg Index, "This shop does not buy items.", BrightRed
        Exit Sub
    End If
    
    If CanTrade = False Then
        PlayerMsg Index, "This shop does not want that item.", BrightRed
        Exit Sub
    End If
    
    Call SellItem(Index, ItemNum, Amount)
    
    Set Buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    Dim BankTab As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankTab = Buffer.ReadLong
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots Index, BankTab, oldSlot, newSlot
    
    Set Buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim bankslot As Long
    Dim Amount As Long
    Dim BankTab As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankTab = Buffer.ReadLong
    bankslot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    TakeBankItem Index, BankTab, bankslot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    Dim BankTab As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankTab = Buffer.ReadLong
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    GiveBankItem Index, BankTab, invSlot, Amount
    
    Set Buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SaveBank Index
    SavePlayer Index
    
    TempPlayer(Index).InBank = False
    
    Set Buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long
    Dim y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    
    If GetPlayerAccess(Index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX Index, x
        SetPlayerY Index, y
        SendPlayerXYToMap Index
    End If
    
    Set Buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(Index).TargetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(Index).Target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = Index Then
        PlayerMsg Index, "You can't trade with yourself.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(Index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).x
    tY = Player(tradeTarget).y
    sX = Player(Index).x
    sY = Player(Index).y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg Index, "You need to be standing next to someone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg Index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = Index
    SendTradeRequest tradeTarget, Index
End Sub

Sub HandleAcceptTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim i As Long

    tradeTarget = TempPlayer(Index).TradeRequest
    ' let them know they're trading
    PlayerMsg Index, "You have accepted " & Trim$(GetPlayerName(tradeTarget)) & "'s trade request.", BrightGreen
    PlayerMsg tradeTarget, Trim$(GetPlayerName(Index)) & " has accepted your trade request.", BrightGreen
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
    TempPlayer(tradeTarget).TradeRequest = 0
    ' set that they're trading with each other
    TempPlayer(Index).InTrade = tradeTarget
    TempPlayer(tradeTarget).InTrade = Index
    ' clear out their trade offers
    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).value = 0
    Next
    ' Used to init the trade window clientside
    SendTrade Index, tradeTarget
    SendTrade tradeTarget, Index
    ' Send the offer data - Used to clear their client
    SendTradeUpdate Index, 0
    SendTradeUpdate Index, 1
    SendTradeUpdate tradeTarget, 0
    SendTradeUpdate tradeTarget, 1
End Sub

Sub HandleDeclineTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(Index).TradeRequest, GetPlayerName(Index) & " has declined your trade request.", BrightRed
    PlayerMsg Index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(Index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim ItemNum As Long
    
    TempPlayer(Index).AcceptTrade = True
    
    tradeTarget = TempPlayer(Index).InTrade
    
    ' if not both of them accept, then exit
    If Not TempPlayer(tradeTarget).AcceptTrade Then
        SendTradeStatus Index, 2
        SendTradeStatus tradeTarget, 1
        Exit Sub
    End If
    
    ' take their items
    For i = 1 To MAX_INV
        ' player
        If TempPlayer(Index).TradeOffer(i).Num > 0 Then
            ItemNum = Player(Index).Inv(TempPlayer(Index).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem(i).Num = ItemNum
                tmpTradeItem(i).value = TempPlayer(Index).TradeOffer(i).value
                ' take item
                TakeInvSlot Index, TempPlayer(Index).TradeOffer(i).Num, tmpTradeItem(i).value
            End If
        End If
        ' target
        If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
            ItemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            If ItemNum > 0 Then
                ' store temp
                tmpTradeItem2(i).Num = ItemNum
                tmpTradeItem2(i).value = TempPlayer(tradeTarget).TradeOffer(i).value
                ' take item
                TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).value
            End If
        End If
    Next
    
    ' taken all items. now they can't not get items because of no inventory space.
    For i = 1 To MAX_INV
        ' player
        If tmpTradeItem2(i).Num > 0 Then
            ' give away!
            GiveInvItem Index, tmpTradeItem2(i).Num, tmpTradeItem2(i).value, False
        End If
        ' target
        If tmpTradeItem(i).Num > 0 Then
            ' give away!
            GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).value, False
        End If
    Next
    
    SendInventory Index
    SendInventory tradeTarget
    
    ' they now have all the items. Clear out values + let them out of the trade.
    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).value = 0
    Next

    TempPlayer(Index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg Index, "Trade completed.", BrightGreen
    PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
    SendCloseTrade Index
    SendCloseTrade tradeTarget
End Sub

Sub HandleDeclineTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(Index).InTrade

    For i = 1 To MAX_INV
        TempPlayer(Index).TradeOffer(i).Num = 0
        TempPlayer(Index).TradeOffer(i).value = 0
        TempPlayer(tradeTarget).TradeOffer(i).Num = 0
        TempPlayer(tradeTarget).TradeOffer(i).value = 0
    Next

    TempPlayer(Index).InTrade = 0
    TempPlayer(tradeTarget).InTrade = 0
    
    PlayerMsg Index, "You declined the trade.", BrightRed
    PlayerMsg tradeTarget, GetPlayerName(Index) & " has declined the trade.", BrightRed
    
    SendCloseTrade Index
    SendCloseTrade tradeTarget
End Sub

Sub HandleTradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    Amount = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(Index, invSlot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, invSlot) Then
        Exit Sub
    End If
    
    If item(ItemNum).Tradable = False Then
        PlayerMsg Index, "You cannot trade this item.", BrightRed
        Exit Sub
    End If

    If item(ItemNum).Stackable = 1 Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invSlot Then
                ' add amount
                TempPlayer(Index).TradeOffer(i).value = TempPlayer(Index).TradeOffer(i).value + Amount
                ' clamp to limits
                If TempPlayer(Index).TradeOffer(i).value > GetPlayerInvItemValue(Index, invSlot) Then
                    TempPlayer(Index).TradeOffer(i).value = GetPlayerInvItemValue(Index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(Index).AcceptTrade = False
                TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
                
                SendTradeStatus Index, 0
                SendTradeStatus TempPlayer(Index).InTrade, 0
                
                SendTradeUpdate Index, 0
                SendTradeUpdate TempPlayer(Index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(Index).TradeOffer(i).Num = invSlot Then
                PlayerMsg Index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(Index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(Index).TradeOffer(EmptySlot).Num = invSlot
    TempPlayer(Index).TradeOffer(EmptySlot).value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(Index).AcceptTrade = False
    TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeSlot = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(Index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(Index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(Index).TradeOffer(tradeSlot).value = 0
    
    If TempPlayer(Index).AcceptTrade Then TempPlayer(Index).AcceptTrade = False
    If TempPlayer(TempPlayer(Index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(Index).InTrade).AcceptTrade = False
    
    SendTradeStatus Index, 0
    SendTradeStatus TempPlayer(Index).InTrade, 0
    
    SendTradeUpdate Index, 0
    SendTradeUpdate TempPlayer(Index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    sType = Buffer.ReadLong
    Slot = Buffer.ReadLong
    hotbarNum = Buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(Index).Hotbar(hotbarNum).Slot = 0
            Player(Index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(Index).Inv(Slot).Num > 0 Then
                    If Len(Trim$(item(GetPlayerInvItemNum(Index, Slot)).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Inv(Slot).Num
                        Player(Index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(Index).Spell(Slot) > 0 Then
                    If Len(Trim$(Spell(Player(Index).Spell(Slot)).Name)) > 0 Then
                        Player(Index).Hotbar(hotbarNum).Slot = Player(Index).Spell(Slot)
                        Player(Index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar Index
    
    Set Buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    
    Select Case Player(Index).Hotbar(Slot).sType
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If Player(Index).Inv(i).Num > 0 Then
                    If Player(Index).Inv(i).Num = Player(Index).Hotbar(Slot).Slot Then
                        UseItem Index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' spell
            For i = 1 To MAX_PLAYER_SPELLS
                If Player(Index).Spell(i) > 0 Then
                    If Player(Index).Spell(i) = Player(Index).Hotbar(Slot).Slot Then
                        BufferSpell Index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set Buffer = Nothing
End Sub

' projectile
Private Sub HandleProjecTileAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim CurEquipment As Long
    ' prevent subscript
    If Index > MAX_PLAYERS Or Index < 1 Then Exit Sub
    
    ' get the players current equipment
    CurEquipment = GetPlayerEquipment(Index, Weapon)
    
    ' check if they've got equipment
    If CurEquipment < 1 Or CurEquipment > MAX_ITEMS Then Exit Sub
    
    Call CreateProjectile(Index, 1, Player(Index).Map)
    
End Sub

Private Sub HandleMoveItemToNewBankTab(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim NewTab As Long, OldTab As Long, Slot As Long
    Dim item As Long, Amount As Long, newSlot As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    NewTab = Buffer.ReadLong
    OldTab = Buffer.ReadLong
    Slot = Buffer.ReadLong
    
    With Bank(Index)
        item = .BankTab(OldTab).item(Slot).Num
        Amount = .BankTab(OldTab).item(Slot).value
        newSlot = BankTabCanContainItem(Index, NewTab, item)
        
        If newSlot >= 0 Then
            .BankTab(OldTab).item(Slot).Num = 0
            .BankTab(OldTab).item(Slot).value = 0
            
            .BankTab(NewTab).item(newSlot).Num = item
            .BankTab(NewTab).item(newSlot).value = Amount + .BankTab(NewTab).item(newSlot).value
        End If
    End With
    
    SendBank (Index)
    
    Set Buffer = Nothing
End Sub

Private Sub HandleProgressConversation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Npc As Long, state As Long, Choice As Long
    
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Npc = Buffer.ReadLong
    state = Buffer.ReadLong
    Choice = Buffer.ReadLong
    
    Call HandleConversation(Index, Npc, state, Choice)
    Set Buffer = Nothing
End Sub

Public Function BankTabCanContainItem(ByVal Index As Long, ByVal BankTab As Long, ByVal ItemNum As Long) As Long
Dim i As Long

    For i = 1 To MAX_BANK
        If Bank(Index).BankTab(BankTab).item(i).Num = ItemNum Then
            BankTabCanContainItem = i
            Exit Function
        End If
    Next
    
    For i = 1 To MAX_BANK
        If Bank(Index).BankTab(BankTab).item(i).Num = 0 Then
            BankTabCanContainItem = i
            Exit Function
        End If
    Next
    
    Call PlayerMsg(Index, "This bank page is full!", BrightRed)
    BankTabCanContainItem = -1
End Function
