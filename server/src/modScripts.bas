Attribute VB_Name = "modScripts"
Option Explicit

Public Function UseItemScript(ByVal Index As Long, ByVal ItemNum As Long, ByVal invSlot As Long) As Boolean
    Select Case ItemNum
    
    End Select
End Function

Public Function CustomMapTile(ByVal Index As Long, ByVal x As Long, ByVal y As Long) As Boolean
Dim MapNum As Long
Dim CanMove As Boolean

    MapNum = Player(Index).Map
    CanMove = True
    
    If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_CUSTOM Then
        CustomMapTile = CanMove
        Exit Function
    End If
    
    Select Case MapNum
        Case 1
            If x = 5 And y = 4 Then
                CanMove = False
            End If
    End Select
    
    CustomMapTile = CanMove
End Function

Public Function CustomShop_Buy(ByVal Index As Long, ByVal ShopNum As Long, ByVal Slot As Long, ByVal Amount As Long) As Boolean
Dim ItemNum As Long

    ItemNum = Shop(ShopNum).TradeItem(Slot).item
    
    Select Case ShopNum
        Case Else
            CustomShop_Buy = False
    End Select

End Function

Public Function CustomShop_Sell(ByVal Index As Long, ByVal ShopNum As Long, ByVal Slot As Long, ByVal Amount As Long) As Boolean
Dim ItemNum As Long
Dim ItemAmount As Long

    ItemNum = GetPlayerInvItemNum(Index, Slot)
    ItemAmount = GetPlayerInvItemValue(Index, Slot)

    Select Case ShopNum
        Case Else
            CustomShop_Sell = False
    End Select

End Function

Public Sub KilledNpc(ByVal Index As Long, ByVal NpcNum As Long)

    Select Case NpcNum
        Case Else
            
    End Select
End Sub
