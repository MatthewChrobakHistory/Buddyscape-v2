Attribute VB_Name = "modConversations"
Option Explicit

Public Const OLDMAN As Byte = 1
Public Const END_CONVERSATION As Integer = -1
Public Const START_CONVERSATION As Byte = 0

Public Sub HandleConversation(ByVal Index As Long, ByVal Npc As Long, ByVal State As Long, ByVal Choice As Byte)
    Select Case Npc
        Case OLDMAN
            HandleOldMan Index, State, Choice
    End Select
End Sub

Public Sub StartConversation(ByVal Index As Long, ByVal Npc As Long)

    Select Case Npc
        Case OLDMAN
            Call SendProgressConversation(Index, Npc, START_CONVERSATION, "Welcome to my house!", "Thank you!")
    End Select

End Sub

Public Sub EndConversation(ByVal Index As Long)

    Call SendProgressConversation(Index, 0, END_CONVERSATION)

End Sub


'        Case ConversationPoint
'            Select Case Choice
'                Case 1
'
'                Case 2
'
'                Case 3
'
'                Case 4
'            End Select

Private Sub HandleOldMan(ByVal Index As Long, ByVal State As Long, ByVal Choice As Byte)
Dim Face As Long
    
    Face = 1
    
    Select Case State
        Case START_CONVERSATION
            Select Case Choice
                Case 1
                    EndConversation Index
            End Select
    End Select

End Sub

