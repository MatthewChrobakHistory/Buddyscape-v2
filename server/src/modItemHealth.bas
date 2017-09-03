Attribute VB_Name = "modItemHealth"
Option Explicit

Public ItemHealth() As Long
Public LastItemID As Long

Public Sub LoadItemHealthList()
    Dim f As Long
    Dim File As String
    
    File = App.Path + "\data\ItemHealth.txt"
    f = FreeFile
    
    If FileExist("\data\ItemHealth.txt") = False Then
        LastItemID = 1
        SaveItemHealthList
    End If
    
    Open File For Input As #f
    Input #f, LastItemID
    Close #f
    
    ReDim ItemHealth(1 To LastItemID) As Long
End Sub

Public Sub SaveItemHealthList()
    Dim f As Long
    Dim File As String
    
    If LastItemID = 0 Then Exit Sub
    
    File = App.Path + "\data\ItemHealth.txt"
    f = FreeFile

    Open File For Output As #f
    Print #f, LastItemID
    Close #f
End Sub

Public Function GetNewItemID() As Long
    Dim i As Long
    
    For i = 1 To LastItemID
        If ItemHealth(i) = 0 Then
            GetNewItemID = i
        End If
    Next
    
    LastItemID = LastItemID + 1
    ReDim Preserve ItemHealth(1 To LastItemID) As Long
    GetNewItemID = LastItemID
End Function

Public Function IsDegradableItem(ByVal item As Long) As Long
    Dim amount As Long
    
    Select Case item
        Case Else
            amount = 0
    End Select
    
    IsDegradableItem = amount
End Function

Public Sub DegradeBankItem(ByVal index As Long, ByVal BankTab As Long, ByVal slot As Long)
Dim NewValue As Long

    ' Does the item degrade?
    If IsDegradableItem(Bank(index).BankTab(BankTab).item(slot).Num) > 0 Then
        NewValue = Bank(index).BankTab(BankTab).item(slot).CustomID - 1
        
        If NewValue <= 0 Then
            NewValue = 0
            Call DegradeItem(index, Bank(index).BankTab(BankTab).item(slot).Num, BankTab, slot)
        End If
        
        ItemHealth(Bank(index).BankTab(BankTab).item(slot).CustomID) = NewValue
    End If
End Sub

Public Sub DegradeInventoryItem(ByVal index As Long, ByVal slot As Long)
Dim NewValue As Long
    
    ' Does the item degrade?
    If IsDegradableItem(Player(index).Inv(slot).Num) > 0 Then
        NewValue = Player(index).Inv(slot).CustomID - 1
        
        If NewValue <= 0 Then
            NewValue = 0
            Call DegradeItem(index, Player(index).Inv(slot).Num, , slot)
        End If
        
        ItemHealth(Player(index).Inv(slot).CustomID) = NewValue
    End If
End Sub

Public Sub DegradeEquipmentItem(ByVal index As Long, ByVal slot As Long)
Dim NewValue As Long

    ' Does the item degrade?
    If IsDegradableItem(Player(index).Equipment(slot).Num) Then
        NewValue = Player(index).Equipment(slot).CustomID - 1
        
        If NewValue <= 0 Then
            NewValue = 0
            Call DegradeItem(index, Player(index).Equipment(slot).Num, , , slot)
        End If
        
        ItemHealth(Player(index).Equipment(slot).CustomID) = NewValue
    End If
End Sub

Public Sub DegradeItem(ByVal index As Long, ByVal item As Long, Optional ByVal BankTab As Byte = 0, Optional ByVal bankslot As Byte = 0, Optional ByVal inventoryslot As Byte = 0, Optional ByVal equipmentslot As Byte = 0)
    Dim NewItem As Long
    Dim NewValue As Long
    
    NewItem = 0
    NewValue = 1
    
    ' What item do we/should we get in return? 0 for item degrades.
    Select Case item
        Case 500
            NewItem = 1
    End Select
    
    If bankslot > 0 Then
        Bank(index).BankTab(BankTab).item(bankslot).Num = NewItem
        Bank(index).BankTab(BankTab).item(bankslot).CustomID = 0
        Bank(index).BankTab(BankTab).item(bankslot).value = 0
        If NewItem > 0 Then Bank(index).BankTab(BankTab).item(bankslot).value = NewValue
    ElseIf inventoryslot > 0 Then
        Player(index).Inv(inventoryslot).Num = NewItem
        Player(index).Inv(inventoryslot).CustomID = 0
        Player(index).Inv(inventoryslot).value = 0
        If NewItem > 0 Then Player(index).Inv(inventoryslot).value = NewValue
    ElseIf equipmentslot > 0 Then
        Player(index).Equipment(equipmentslot).Num = NewItem
        Player(index).Equipment(equipmentslot).CustomID = 0
        Player(index).Equipment(equipmentslot).value = 0
        If NewItem > 0 Then Player(index).Equipment(equipmentslot).value = NewValue
    End If
End Sub
