Attribute VB_Name = "modAccountEditor"
Option Explicit

Public EditUserIndex As Byte

Public Sub AddInfo(ByVal Text As String)

frmAccountEditor.lblInfo.Caption = Text

End Sub

Public Sub AccountEditorInit(ByVal Index As Byte)
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    .FrameAccountDetails.Visible = True
    .txtUserName.Text = Trim$(Player(Index).Name)
    .txtPassword.Text = Trim$(Player(Index).Password)
    .txtAccess.Text = Trim$(Player(Index).Access)
    .txtSprite.Text = Player(Index).Sprite

    
    'bank
    .frameBank.Visible = True
    For i = 1 To 99
        If Bank(EditUserIndex).BankTab(frmAccountEditor.scrlBanktab.value).item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(item(Bank(EditUserIndex).BankTab(frmAccountEditor.scrlBanktab.value).item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).BankTab(frmAccountEditor.scrlBanktab.value).item(i).value)
    Next
    .lstBank.ListIndex = 0
    
    'inventory
    .frameInventory.Visible = True
    For i = 1 To 35
        If Player(Index).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(item(Player(Index).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & Player(Index).Inv(i).value)
    Next
    .lstInventory.ListIndex = 0
End With

End Sub

Public Sub BankEditorInit()
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    .lstBank.Clear
    For i = 1 To 99 '99 bank space
        If Bank(EditUserIndex).BankTab(frmAccountEditor.scrlBanktab.value).item(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(item(Bank(EditUserIndex).BankTab(frmAccountEditor.scrlBanktab.value).item(i).Num).Name)
        End If
        .lstBank.AddItem (i & ": " & ItemName & "  x  " & Bank(EditUserIndex).BankTab(frmAccountEditor.scrlBanktab.value).item(i).value)
    Next
    .lstBank.ListIndex = 0
End With

End Sub

Public Sub SaveEditPlayer(ByVal Index As Byte)

With Player(Index)
    .Name = frmAccountEditor.txtUserName.Text
    .Password = frmAccountEditor.txtPassword.Text
    .Access = frmAccountEditor.txtAccess.Text
    .Sprite = frmAccountEditor.txtSprite.Text
End With

Call SendPlayerData(Index)

Call PlayerMsg(Index, "Your account was edited by an admin!", Pink)

End Sub

Public Sub InvEditorInit()
Dim i As Byte
Dim ItemName As String

With frmAccountEditor
    'inventory
    .lstInventory.Clear
    .frameInventory.Visible = True
    For i = 1 To 35
        If Player(EditUserIndex).Inv(i).Num = 0 Then
            ItemName = "None"
        Else
            ItemName = Trim$(item(Player(EditUserIndex).Inv(i).Num).Name)
        End If
        .lstInventory.AddItem (i & ": " & ItemName & "  x  " & Player(EditUserIndex).Inv(i).value)
    Next
    .lstInventory.ListIndex = 0
End With

End Sub



