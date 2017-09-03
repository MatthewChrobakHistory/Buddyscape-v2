Attribute VB_Name = "modGUI"
Option Explicit

Public Sub SetupMainGUI()
Dim i As Long

    With frmMain

    .height = 9430
    .width = 11865
    .BackColor = RGB(44, 37, 31)
    
    .picBank.top = 10
    .picBank.Left = 10
    
    .picTrade.Left = 10
    .picTrade.top = 10
            
    '.picAdmin.top = 8
    '.picAdmin.Left = 800
    
    .picCover.top = 552
    .picCover.Left = 800
    
    .picSSMap.top = 552
    .picSSMap.Left = 824
    
    Dim amt As Long
    amt = 50
    
    .picClan.top = 283 - amt
    .picClan.Left = 541
    
    .picSpells.top = 283 - amt
    .picSpells.Left = 541
    
    .picOptions.top = 283 - amt
    .picOptions.Left = 541
    
    .picSkills.top = 283 - amt
    .picSkills.Left = 541
    
    .picInventory.top = 283 - amt
    .picInventory.Left = 541
    
    .picCharacter.top = 283 - amt
    .picCharacter.Left = 541
    
    .picCombat.top = 283 - amt
    .picCombat.Left = 541
    .imgSpec_Void.Left = 22
    .imgSpec_Void.top = 225
    .imgSpec.Left = 22
    .imgSpec.top = 225
    .imgSpec.width = 75
    .lblSpec.Left = 22
    .lblSpec.top = 225
    .lblSpec.width = 150
    .lblSpec.height = 17
    
    .picQuests.top = 283 - amt
    .picQuests.Left = 541
    
    .picShop.Left = 112
    .picShop.top = 32
    
    .picShopItem.top = 283
    .picShopItem.Left = 541
    
    .imgSpellUp.top = .picSpells.top + 30
    .imgSpellUp.Left = .picSpells.Left + .picSpells.width
    
    .imgSpellDown.top = .picSpells.top + .picSpells.height - .imgSpellDown.height
    .imgSpellDown.Left = .picSpells.Left + .picSpells.width
    
    .lblPing.Left = 568
    .lblPing.top = 128
    
    Dim chatColor As Long
    
    chatColor = RGB(13, 13, 13)
    
    .picConv.Left = 12
    .picConv.top = 442
    .txtMyChat.BackColor = chatColor
    
    .txtGlobalChat.Left = 12
    .txtGlobalChat.top = 442
    .txtGlobalChat.BackColor = chatColor
    .optGlobal.BackColor = RGB(44, 37, 31)
    
    .txtPrivateChat.Left = 12
    .txtPrivateChat.top = 442
    .txtPrivateChat.Visible = False
    .txtPrivateChat.BackColor = chatColor
    .optPrivate.BackColor = RGB(44, 37, 31)
    
    .txtGameChat.top = 442
    .txtGameChat.Left = 12
    .txtGameChat.Visible = False
    .txtGameChat.BackColor = chatColor
    .optGame.BackColor = RGB(44, 37, 31)
    
    .txtClanChat.top = 442
    .txtClanChat.Left = 12
    .txtClanChat.Visible = False
    .txtClanChat.BackColor = chatColor
    .optClan.BackColor = RGB(44, 37, 31)
    
    ' Tabs
    For i = 1 To 7
        .imgButton(i).width = 36
        .imgButton(i).height = 37
        .imgButton(i).top = 192
        .imgButton(i).Left = 512 + ((i - 1) * 36)
    Next
    
    For i = 8 To Tabs.Tab_Count - 1
        .imgButton(i).width = 36
        .imgButton(i).height = 37
        .imgButton(i).top = 510
        .imgButton(i).Left = 512 + ((i - 7) * 36)
    Next
    
    .imgHPBar.height = 16
    .imgPRBar.height = 16
    .imgSUMBar.height = 16
    
    .imgHPBar.width = 241
    .imgPRBar.width = 241
    .imgSUMBar.width = 241
    
    .imgHPBar.Left = 518
    .imgPRBar.Left = 518
    .imgSUMBar.Left = 518
    
    .imgHPBar.top = 28
    .imgPRBar.top = 50
    .imgSUMBar.top = 72
    

    
End With
    
End Sub
