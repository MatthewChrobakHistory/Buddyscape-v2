Attribute VB_Name = "modQuests"
Public Const MAX_QUESTS As Long = 1
Public QuestName(1 To MAX_QUESTS) As String

Public Const QUEST_A As Byte = 1
Public Const QUEST_B As Byte = QUEST_A + 1
Public Const QUEST_C As Byte = QUEST_B + 1

Public Sub LoadQuestNames()

    ' We want to load the quests alphabetically.
    QuestName(QUEST_A) = "QuestA"

End Sub


Public Sub LoadQuests()

    

End Sub
