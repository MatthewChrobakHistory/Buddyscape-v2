VERSION 5.00
Begin VB.Form frmEditor_NPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Npc Editor"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_NPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   606
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   16
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   15
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   8640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "NPC Properties"
      Height          =   8415
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2280
         TabIndex        =   52
         Text            =   "0"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Frame fraDrops 
         Caption         =   "Drop: 1"
         Height          =   2415
         Left            =   120
         TabIndex        =   43
         Top             =   5880
         Width           =   4815
         Begin VB.Frame fraDrop 
            Caption         =   "Drop"
            Height          =   1815
            Left            =   120
            TabIndex        =   45
            Top             =   600
            Width           =   4575
            Begin VB.TextBox txtDropValue 
               Height          =   270
               Left            =   1080
               TabIndex        =   51
               Top             =   1320
               Width           =   3375
            End
            Begin VB.HScrollBar scrlDropItem 
               Height          =   255
               Left            =   120
               Max             =   255
               TabIndex        =   47
               Top             =   960
               Value           =   1
               Width           =   4335
            End
            Begin VB.TextBox txtDropChance 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               TabIndex        =   46
               Text            =   "0"
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label lblValue 
               AutoSize        =   -1  'True
               Caption         =   "Value: "
               Height          =   180
               Left            =   120
               TabIndex        =   50
               Top             =   1320
               UseMnemonic     =   0   'False
               Width           =   540
            End
            Begin VB.Label lblDropItem 
               AutoSize        =   -1  'True
               Caption         =   "Item: None"
               Height          =   180
               Left            =   120
               TabIndex        =   49
               Top             =   600
               Width           =   2535
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Chance out of 100%"
               Height          =   180
               Left            =   120
               TabIndex        =   48
               Top             =   240
               UseMnemonic     =   0   'False
               Width           =   1560
            End
         End
         Begin VB.HScrollBar scrlDrop 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   44
            Top             =   240
            Value           =   1
            Width           =   4575
         End
      End
      Begin VB.Frame fraBonuses 
         BorderStyle     =   0  'None
         Caption         =   "Wearing Bonuses"
         Height          =   3495
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   4815
         Begin VB.HScrollBar scrlDefense 
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   33
            Top             =   3000
            Width           =   1575
         End
         Begin VB.HScrollBar scrlDefense 
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   32
            Top             =   2280
            Width           =   1575
         End
         Begin VB.HScrollBar scrlOffense 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   31
            Top             =   3000
            Width           =   1575
         End
         Begin VB.HScrollBar scrlOffense 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   30
            Top             =   2280
            Width           =   1575
         End
         Begin VB.HScrollBar scrlDefense 
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   29
            Top             =   1560
            Width           =   1575
         End
         Begin VB.HScrollBar scrlOffense 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txtBonus 
            Height          =   270
            Left            =   2880
            TabIndex        =   27
            Top             =   840
            Width           =   1935
         End
         Begin VB.HScrollBar scrlBonus 
            Height          =   255
            Left            =   120
            Max             =   25
            Min             =   1
            TabIndex        =   26
            Top             =   840
            Value           =   1
            Width           =   1695
         End
         Begin VB.HScrollBar scrlDamage 
            Height          =   255
            LargeChange     =   10
            Left            =   1320
            Max             =   255
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
         Begin VB.HScrollBar scrlSpeed 
            Height          =   255
            LargeChange     =   100
            Left            =   3720
            Max             =   3000
            Min             =   100
            SmallChange     =   100
            TabIndex        =   24
            Top             =   360
            Value           =   100
            Width           =   1095
         End
         Begin VB.Label lblTeam 
            Caption         =   "Team: "
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   0
            Width           =   3975
         End
         Begin VB.Label lblDefense 
            Caption         =   "Defense: 0"
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   42
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label lblDefense 
            Caption         =   "Defense: 0"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   41
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label lblOffense 
            Caption         =   "Magic Offense: 0"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   40
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label lblOffense 
            Caption         =   "Ranged Offense: 0"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   39
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label lblDefense 
            Caption         =   "Defense: 0"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   38
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label lblOffense 
            Caption         =   "Melee Offense: 0"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label lblBonus 
            Caption         =   "Attack:"
            Height          =   255
            Left            =   1920
            TabIndex        =   36
            Top             =   870
            Width           =   1215
         End
         Begin VB.Label lblDamage 
            AutoSize        =   -1  'True
            Caption         =   "Damage: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   1065
         End
         Begin VB.Label lblSpeed 
            AutoSize        =   -1  'True
            Caption         =   "Speed: 0.1 sec"
            Height          =   180
            Left            =   2520
            TabIndex        =   34
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   1140
         End
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   1440
         Width           =   1335
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cmbBehaviour 
         Height          =   300
         ItemData        =   "frmEditor_NPC.frx":3332
         Left            =   3480
         List            =   "frmEditor_NPC.frx":3345
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtEXP 
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Spawn Rate (in seconds)"
         Height          =   180
         Left            =   240
         TabIndex        =   53
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level:"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label lblAnimation 
         Caption         =   "Anim: None"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblSprite 
         AutoSize        =   -1  'True
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   660
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Behaviour:"
         Height          =   180
         Left            =   2640
         TabIndex        =   11
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   810
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Range: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Exp:"
         Height          =   180
         Left            =   2520
         TabIndex        =   9
         Top             =   1440
         Width           =   345
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NPC List"
      Height          =   8415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7980
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   8640
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbBehaviour_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Npc(EditorIndex).Type = cmbBehaviour.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbBehaviour_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearNPC EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlSprite.Max = NumCharacters
    scrlAnimation.Max = MAX_ANIMATIONS
    scrlDropItem.Max = MAX_ITEMS
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call NpcEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblTeam_Click()

    Npc(EditorIndex).Team = InputBox("Select a team name.")
    lblTeam.Caption = "Team: " & Npc(EditorIndex).Team

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    NpcEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlAnimation.Value = 0 Then sString = "None" Else sString = Trim$(Animation(scrlAnimation.Value).Name)
    lblAnimation.Caption = "Anim: " & sString
    Npc(EditorIndex).Animation = scrlAnimation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBonus_Change()

    lblBonus.Caption = GetSkillName(scrlBonus.Value) & ": "
    txtBonus.Text = Npc(EditorIndex).Skill(scrlBonus.Value)

End Sub

Private Sub scrlDamage_Change()

    Npc(EditorIndex).Damage = scrlDamage.Value
    lblDamage.Caption = "Damage: " & scrlDamage.Value

End Sub

Private Sub scrlDefense_Change(Index As Integer)

    lblDefense(Index).Caption = GetCombatName(Index) & " defense: " & scrlDefense(Index).Value
    Npc(EditorIndex).Defense(Index) = scrlDefense(Index).Value

End Sub

Private Sub scrlDrop_Change()

    fraDrops.Caption = "Drop: " & scrlDrop.Value
    
    txtDropChance.Text = Npc(EditorIndex).Drop(scrlDrop.Value).Chance
    scrlDropItem.Value = Npc(EditorIndex).Drop(scrlDrop.Value).Item
    txtDropValue.Text = Npc(EditorIndex).Drop(scrlDrop.Value).Amount

End Sub

Private Sub scrlDropItem_Change()

    If scrlDropItem.Value = 0 Then
        Npc(EditorIndex).Drop(scrlDrop.Value).Item = scrlDropItem.Value
        lblDropItem.Caption = "Item: None"
        Exit Sub
    End If

    Npc(EditorIndex).Drop(scrlDrop.Value).Item = scrlDropItem.Value
    lblDropItem.Caption = "Item: " & Trim$(Item(scrlDropItem.Value).Name)

End Sub

Private Sub scrlOffense_Change(Index As Integer)

    lblOffense(Index).Caption = GetCombatName(Index) & " offense: " & scrlOffense(Index).Value
    Npc(EditorIndex).Offense(Index) = scrlOffense(Index).Value

End Sub

Private Sub scrlSpeed_Change()

    Npc(EditorIndex).AttackSpeed = scrlSpeed.Value
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"

End Sub

Private Sub scrlSprite_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblSprite.Caption = "Sprite: " & scrlSprite.Value
    Npc(EditorIndex).Sprite = scrlSprite.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSprite_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRange.Caption = "Range: " & scrlRange.Value
    Npc(EditorIndex).SightRange = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtBonus_Change()

    If IsNumeric(txtBonus.Text) = False Then
        txtBonus.Text = "0"
    End If
    
    Npc(EditorIndex).Skill(scrlBonus.Value) = txtBonus.Text

End Sub

Private Sub txtDropChance_Change()

    If IsNumeric(txtDropChance.Text) = False Then
        txtDropChance.Text = "100"
    End If
    
    If txtDropChance.Text > 100 Or txtDropChance.Text < 0 Then
        txtDropChance.Text = "100"
    End If
    
    Npc(EditorIndex).Drop(scrlDrop.Value).Chance = txtDropChance.Text

End Sub

Private Sub txtDropValue_Change()

    If IsNumeric(txtDropValue.Text) = False Then
        txtDropValue.Text = "1"
    End If
    
    If txtDropValue.Text < 1 Or txtDropValue > MAX_LONG Then
        txtDropValue.Text = 1
    End If
    
    Npc(EditorIndex).Drop(scrlDrop.Value).Amount = txtDropValue.Text
    
End Sub

Private Sub txtEXP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtEXP.Text) > 0 Then Exit Sub
    If IsNumeric(txtEXP.Text) Then Npc(EditorIndex).RewardXP = Val(txtEXP.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtEXP_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtLevel.Text) > 0 Then Exit Sub
    If IsNumeric(txtLevel.Text) Then Npc(EditorIndex).Level = Val(txtLevel.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtlevel_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Npc(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Npc(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtSpawnSecs_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not Len(txtSpawnSecs.Text) > 0 Then Exit Sub
    Npc(EditorIndex).RespawnRate = Val(txtSpawnSecs.Text)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtSpawnSecs_Change", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Npc(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Npc(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_NPC", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
