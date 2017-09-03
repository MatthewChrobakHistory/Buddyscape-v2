Attribute VB_Name = "modClans"
Option Explicit

Public Function GetMemberCount(ByVal ClanIndex As Long) As Long
Dim i As Long, amt As Long

    With Clan(ClanIndex)
        For i = 1 To MAX_PLAYERS
            If .Member(i).PlayerIndex > 0 And .Member(i).PlayerIndex <= MAX_PLAYERS Then
                amt = amt + 1
            End If
        Next
    End With
End Function
    
Public Sub HandleRequestJoinClan(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim ClanName As String * NAME_LENGTH
Dim ClanIndex As Long
Dim state As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ClanName = Buffer.ReadString
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If LCase$(Trim$(Player(i).Name)) = LCase$(Trim$(ClanName)) Then
                If TempPlayer(i).inClan = i Or index = i Then
                    ClanIndex = i
                    Exit For
                Else
                    PlayerMsg index, "The host is not in his clan.", BrightRed
                    Exit Sub
                End If
            End If
        End If
    Next
    
    If ClanIndex = 0 Then
        PlayerMsg index, "The host is not online.", BrightRed
        Exit Sub
    End If
    
    If TempPlayer(index).inClan = ClanIndex Then
        PlayerMsg index, "You are already in this clan.", BrightRed
        Exit Sub
    Else
        With Clan(ClanIndex)
            For i = 1 To MAX_PLAYERS
                If .Member(i).PlayerIndex = 0 Then
                    .Member(i).PlayerIndex = index
                    If ClanIndex = index Then
                        .Member(i).Rank = "[Owner] "
                    Else
                        .Member(i).Rank = "[m] "
                    End If
                    Exit For
                End If
            Next
        End With
    End If
    
    TempPlayer(index).inClan = ClanIndex
    Call UpdateClanList(ClanIndex)
    
    Set Buffer = Nothing
End Sub

Public Sub HandleLeaveClan(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim x As Long
Dim ClanIndex As Long

    ClanIndex = TempPlayer(index).inClan
    If ClanIndex < 1 Then Exit Sub
    
    
    For i = 1 To MAX_PLAYERS
        With Clan(ClanIndex).Member(i)
            If .PlayerIndex = index Then
            
                ' Are they the owner?
                If Trim$(.Rank) = "[Owner]" Then
                    For x = 1 To MAX_PLAYERS
                        With Clan(ClanIndex).Member(x)
                            If .PlayerIndex > 0 Then
                                ' Send them a null clan.
                                SendClanListTo .PlayerIndex, 0
                                
                                ' Now purge their values.
                                PlayerMsg .PlayerIndex, "Clan disbanded.", BrightRed
                                TempPlayer(.PlayerIndex).inClan = 0
                                .PlayerIndex = 0
                                .Rank = vbNullString
                            End If
                        End With
                    Next
                
                ' Nope
                Else
                    PlayerMsg index, "Left clan.", BrightRed
                    TempPlayer(.PlayerIndex).inClan = 0
                    .PlayerIndex = 0
                    .Rank = vbNullString
                    UpdateClanList (ClanIndex)
                    SendClanListTo index, 0
                End If
            
            End If
        End With
    Next
End Sub

Public Sub SendClanListTo(ByVal index As Long, ByVal ClanIndex As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateClanList
    
    If ClanIndex > 0 Then
        Buffer.WriteLong Clan(ClanIndex).LootShare
        For i = 1 To MAX_PLAYERS
            Buffer.WriteLong Clan(ClanIndex).Member(i).PlayerIndex
            Buffer.WriteString Clan(ClanIndex).Member(i).Rank
        Next
    Else
        Buffer.WriteLong -255
    End If
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub UpdateClanList(ByVal ClanIndex As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SUpdateClanList
    
    Buffer.WriteLong Clan(ClanIndex).LootShare
    For i = 1 To MAX_PLAYERS
        Buffer.WriteLong Clan(ClanIndex).Member(i).PlayerIndex
        Buffer.WriteString Clan(ClanIndex).Member(i).Rank
    Next
    
    For i = 1 To MAX_PLAYERS
        If Clan(ClanIndex).Member(i).PlayerIndex > 0 Then
            If IsPlaying(Clan(ClanIndex).Member(i).PlayerIndex) Then
                SendDataTo Clan(ClanIndex).Member(i).PlayerIndex, Buffer.ToArray()
            End If
        End If
    Next
    Set Buffer = Nothing
End Sub
