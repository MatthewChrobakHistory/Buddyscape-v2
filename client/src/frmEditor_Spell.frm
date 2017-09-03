VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Spell Properties"
      Height          =   5295
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtDesc 
         Height          =   975
         Left            =   1440
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   4080
         Width           =   5295
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data"
         Height          =   3615
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   3255
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   3240
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   2640
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   2895
         End
         Begin VB.CheckBox chkAOE 
            Caption         =   "Area of Effect spell?"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAOE 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   3015
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label lblStun 
            Caption         =   "Stun Duration: None"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   3000
            Width           =   2895
         End
         Begin VB.Label lblAnim 
            Caption         =   "Animation: None"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label lblAnimCast 
            Caption         =   "Cast Anim: None"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Label lblAOE 
            Caption         =   "AoE: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label lblRange 
            Caption         =   "Range: Self-cast"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Basic Information"
         Height          =   3615
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txtDamage 
            Height          =   270
            Left            =   960
            TabIndex        =   34
            Text            =   "0"
            Top             =   2760
            Width           =   2175
         End
         Begin VB.ComboBox cmbSound 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   3120
            Width           =   2175
         End
         Begin VB.PictureBox picSprite 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2520
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   26
            Top             =   960
            Width           =   480
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   2295
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   14
            Top             =   2400
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   12
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label Label2 
            Caption         =   "Damage:"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Sound:"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblIcon 
            Caption         =   "Icon: None"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblCool 
            Caption         =   "Cooldown Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2160
            Width           =   2535
         End
         Begin VB.Label lblCast 
            Caption         =   "Casting Time: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Spell List"
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   4740
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
      Top             =   5520
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAOE_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Spell(EditorIndex).AoE = chkAOE.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkAOE_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearSpell EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnim.Value > 0 Then
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.Value).Name)
    Else
        lblAnim.Caption = "Animation: None"
    End If
    Spell(EditorIndex).HitAnim = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnimCast.Value > 0 Then
        lblAnimCast.Caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.Value).Name)
    Else
        lblAnimCast.Caption = "Cast Anim: None"
    End If
    Spell(EditorIndex).CastAnim = scrlAnimCast.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAOE_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAOE.Value > 0 Then
        lblAOE.Caption = "AoE: " & scrlAOE.Value & " tiles."
    Else
        lblAOE.Caption = "AoE: Self-cast"
    End If
    Spell(EditorIndex).AoERange = scrlAOE.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCast.Caption = "Casting Time: " & scrlCast.Value & "s"
    Spell(EditorIndex).CastingTime = scrlCast.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCool_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCool.Caption = "Cooldown Time: " & scrlCool.Value & "s"
    Spell(EditorIndex).CooldownTime = scrlCool.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlIcon_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlIcon.Value > 0 Then
        lblIcon.Caption = "Icon: " & scrlIcon.Value
    Else
        lblIcon.Caption = "Icon: None"
    End If
    Spell(EditorIndex).Icon = scrlIcon.Value
    EditorSpell_BltIcon
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlRange.Value > 0 Then
        lblRange.Caption = "Range: " & scrlRange.Value & " tiles."
    Else
        lblRange.Caption = "Range: Self-cast"
    End If
    Spell(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStun_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlStun.Value > 0 Then
        lblStun.Caption = "Stun Duration: " & scrlStun.Value & "s"
    Else
        lblStun.Caption = "Stun Duration: None"
    End If
    Spell(EditorIndex).StunDuration = scrlStun.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDamage_Change()

    If Not IsNumeric(txtDamage.Text) Then txtDamage.Text = 0
    
    Spell(EditorIndex).Damage = txtDamage.Text

End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).Description = txtDesc.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Spell(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Spell(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
