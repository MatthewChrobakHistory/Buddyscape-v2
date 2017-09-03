VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14610
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   974
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraReward 
      Caption         =   "Making Reward"
      Height          =   975
      Left            =   9240
      TabIndex        =   61
      Top             =   1200
      Width           =   5295
      Begin VB.TextBox txtReward 
         Height          =   270
         Left            =   3240
         TabIndex        =   63
         Top             =   360
         Width           =   1935
      End
      Begin VB.HScrollBar scrlReward 
         Height          =   255
         Left            =   120
         Max             =   25
         Min             =   1
         TabIndex        =   62
         Top             =   360
         Value           =   1
         Width           =   1695
      End
      Begin VB.Label lblReward 
         Caption         =   "Attack:"
         Height          =   255
         Left            =   1920
         TabIndex        =   64
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Frame fraMaking 
      Caption         =   "Making Requirement"
      Height          =   975
      Left            =   9240
      TabIndex        =   57
      Top             =   120
      Width           =   5295
      Begin VB.HScrollBar scrlMakingReq 
         Height          =   255
         Left            =   120
         Max             =   25
         Min             =   1
         TabIndex        =   59
         Top             =   360
         Value           =   1
         Width           =   1695
      End
      Begin VB.TextBox txtMakingReq 
         Height          =   270
         Left            =   3240
         TabIndex        =   58
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblMakingReq 
         Caption         =   "Attack:"
         Height          =   255
         Left            =   1920
         TabIndex        =   60
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Other Data"
      Height          =   4215
      Left            =   3360
      TabIndex        =   29
      Top             =   3600
      Width           =   5775
      Begin VB.Frame fraVitals 
         Caption         =   "Consume Data"
         Height          =   1455
         Left            =   1080
         TabIndex        =   42
         Top             =   1440
         Visible         =   0   'False
         Width           =   3735
         Begin VB.HScrollBar scrlAddHp 
            Height          =   255
            Left            =   1320
            Max             =   1000
            TabIndex        =   45
            Top             =   360
            Width           =   2175
         End
         Begin VB.HScrollBar scrlAddMP 
            Height          =   255
            Left            =   1320
            Max             =   1000
            TabIndex        =   44
            Top             =   720
            Width           =   2175
         End
         Begin VB.HScrollBar scrlAddExp 
            Height          =   255
            Left            =   1320
            Max             =   1000
            TabIndex        =   43
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblAddHP 
            AutoSize        =   -1  'True
            Caption         =   "Add HP: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   48
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   1140
         End
         Begin VB.Label lblAddMP 
            AutoSize        =   -1  'True
            Caption         =   "Add MP: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   720
            UseMnemonic     =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblAddExp 
            AutoSize        =   -1  'True
            Caption         =   "Add Exp: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   46
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   1200
         End
      End
      Begin VB.Frame fraEquipment 
         Caption         =   "Equipment Data"
         Height          =   2295
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   5535
         Begin VB.ComboBox cmbWeaponType 
            Height          =   300
            ItemData        =   "frmEditor_Item.frx":3332
            Left            =   2400
            List            =   "frmEditor_Item.frx":333F
            TabIndex        =   78
            Top             =   840
            Width           =   2895
         End
         Begin VB.Frame fraProjectiles 
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   1095
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   5295
            Begin VB.HScrollBar scrlProjectileSpeed 
               Height          =   255
               Left            =   2280
               TabIndex        =   38
               Top             =   840
               Width           =   2895
            End
            Begin VB.HScrollBar scrlProjectileRange 
               Height          =   255
               Left            =   2280
               TabIndex        =   37
               Top             =   480
               Width           =   2895
            End
            Begin VB.HScrollBar scrlProjectilePic 
               Height          =   255
               Left            =   2280
               TabIndex        =   36
               Top             =   120
               Width           =   2895
            End
            Begin VB.Label lblProjectilesSpeed 
               Caption         =   "Projectile Speed: 0"
               Height          =   225
               Left            =   120
               TabIndex        =   41
               Top             =   840
               Width           =   1980
            End
            Begin VB.Label lblProjectileRange 
               Caption         =   "Projectile Range: 0"
               Height          =   180
               Left            =   120
               TabIndex        =   40
               Top             =   480
               Width           =   1965
            End
            Begin VB.Label lblProjectilePiC 
               BackStyle       =   0  'Transparent
               Caption         =   "Projectile Image: 0"
               Height          =   270
               Left            =   120
               TabIndex        =   39
               Top             =   120
               Width           =   1875
            End
         End
         Begin VB.ComboBox cmbTool 
            Height          =   300
            ItemData        =   "frmEditor_Item.frx":3359
            Left            =   1320
            List            =   "frmEditor_Item.frx":3369
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   360
            Width           =   1215
         End
         Begin VB.HScrollBar scrlPaperdoll 
            Height          =   255
            Left            =   3960
            TabIndex        =   31
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Combat Type:"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Object Tool:"
            Height          =   180
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   945
         End
         Begin VB.Label lblPaperdoll 
            AutoSize        =   -1  'True
            Caption         =   "Paperdoll: 0"
            Height          =   180
            Left            =   2760
            TabIndex        =   33
            Top             =   360
            Width           =   915
         End
      End
   End
   Begin VB.Frame fraBonuses 
      Caption         =   "Wearing Bonuses"
      Height          =   3615
      Left            =   9240
      TabIndex        =   24
      Top             =   3360
      Width           =   5295
      Begin VB.HScrollBar scrlDefense 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   2760
         Max             =   500
         TabIndex        =   75
         Top             =   3000
         Width           =   2295
      End
      Begin VB.HScrollBar scrlDefense 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2760
         Max             =   500
         TabIndex        =   73
         Top             =   2280
         Width           =   2295
      End
      Begin VB.HScrollBar scrlOffense 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   240
         Max             =   500
         TabIndex        =   71
         Top             =   3000
         Width           =   2295
      End
      Begin VB.HScrollBar scrlOffense 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   240
         Max             =   500
         TabIndex        =   69
         Top             =   2280
         Width           =   2295
      End
      Begin VB.HScrollBar scrlDefense 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   2760
         Max             =   500
         TabIndex        =   66
         Top             =   1560
         Width           =   2295
      End
      Begin VB.HScrollBar scrlOffense 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   240
         Max             =   500
         TabIndex        =   65
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtBonus 
         Height          =   270
         Left            =   3240
         TabIndex        =   53
         Top             =   840
         Width           =   1935
      End
      Begin VB.HScrollBar scrlBonus 
         Height          =   255
         Left            =   120
         Max             =   25
         Min             =   1
         TabIndex        =   52
         Top             =   840
         Value           =   1
         Width           =   1695
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1440
         Max             =   255
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   3960
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   25
         Top             =   360
         Value           =   100
         Width           =   1095
      End
      Begin VB.Label lblDefense 
         Caption         =   "Defense: 0"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   76
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label lblDefense 
         Caption         =   "Defense: 0"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   74
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lblOffense 
         Caption         =   "Magic Offense: 0"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   72
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label lblOffense 
         Caption         =   "Ranged Offense: 0"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   70
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lblDefense 
         Caption         =   "Defense: 0"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   68
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblOffense 
         Caption         =   "Melee Offense: 0"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   67
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblBonus 
         Caption         =   "Attack:"
         Height          =   255
         Left            =   1920
         TabIndex        =   51
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   28
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   2640
         TabIndex        =   27
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info"
      Height          =   3375
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   5775
      Begin VB.CheckBox chkTwoHanded 
         Caption         =   "Two Handed?"
         Height          =   255
         Left            =   4200
         TabIndex        =   77
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox chkStackable 
         Caption         =   "Stackable"
         Height          =   180
         Left            =   2880
         TabIndex        =   50
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkTradable 
         Caption         =   "Tradable"
         Height          =   180
         Left            =   2880
         TabIndex        =   49
         Top             =   840
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   1440
         Max             =   5
         TabIndex        =   22
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1410
         Width           =   1935
      End
      Begin VB.TextBox txtDesc 
         Height          =   1335
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   1920
         Width           =   5415
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4080
         Max             =   30000
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   4560
         Max             =   5
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":338A
         Left            =   120
         List            =   "frmEditor_Item.frx":33B8
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access Req: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label4 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   17
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   3000
         TabIndex        =   16
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame fraWearing 
      Caption         =   "Wearing Requirements"
      Height          =   975
      Left            =   9240
      TabIndex        =   6
      Top             =   2280
      Width           =   5295
      Begin VB.TextBox txtWearingReq 
         Height          =   270
         Left            =   3240
         TabIndex        =   56
         Top             =   360
         Width           =   1935
      End
      Begin VB.HScrollBar scrlWearingReq 
         Height          =   255
         Left            =   120
         Max             =   25
         Min             =   1
         TabIndex        =   55
         Top             =   360
         Value           =   1
         Width           =   1695
      End
      Begin VB.Label lblWearingReq 
         Caption         =   "Attack:"
         Height          =   255
         Left            =   1920
         TabIndex        =   54
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7260
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub chkStackable_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Stackable = chkStackable.Value

End Sub

Private Sub chkTradable_Click()
        
        If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
        Item(EditorIndex).Tradable = chkTradable.Value
                
End Sub

Private Sub chkTwoHanded_Click()
    Item(EditorIndex).isTwoHander = chkTwoHanded.Value
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).Sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).Sound = "None."
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).EquipmentType = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbWeaponType_Click()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).CombatType = cmbWeaponType.ListIndex + 1

End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    scrlPic.Max = NumItems
    scrlAnim.Max = MAX_ANIMATIONS
    scrlPaperdoll.Max = NumPaperdolls
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If (cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
        fraProjectiles.Visible = True
        lblSpeed.Visible = True
        scrlSpeed.Visible = True
        chkTwoHanded.Visible = True
    Else
        fraProjectiles.Visible = False
        lblSpeed.Visible = False
        scrlSpeed.Visible = False
        chkTwoHanded.Visible = False
    End If

    If (cmbType.ListIndex > ITEM_TYPE_NONE) And (cmbType.ListIndex <= ITEM_TYPE_RING) Then
        fraEquipment.Visible = True
        fraEquipment.Visible = True
        fraBonuses.Visible = True
        fraWearing.Visible = True
        scrlDamage.Visible = True
        lblDamage.Visible = True
    Else
        fraEquipment.Visible = False
        fraEquipment.Visible = False
        lblDamage.Visible = False
        scrlDamage.Visible = False
        fraBonuses.Visible = False
        fraWearing.Visible = False
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If
    
    Item(EditorIndex).ItemType = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddVital(Vitals.HitPoints) = scrlAddHp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddVital(Vitals.Prayer) = scrlAddMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddVital(Vitals.Summoning) = scrlAddExp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.Value).Name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBonus_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblBonus.Caption = GetSkillName(scrlBonus.Value) & ":"
    txtBonus.Text = Item(EditorIndex).SkillBonus(scrlBonus.Value)

End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    Item(EditorIndex).Damage = scrlDamage.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDefense_Change(Index As Integer)

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Defense(Index) = scrlDefense(Index).Value
    lblDefense(Index).Caption = GetCombatName(Index) & " Defense: " & scrlDefense(Index).Value
    
End Sub

Private Sub scrlMakingReq_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblMakingReq.Caption = GetSkillName(scrlMakingReq.Value) & ":"
    txtMakingReq.Text = Item(EditorIndex).SkillMakeReq(scrlMakingReq.Value)

End Sub

Private Sub scrlOffense_Change(Index As Integer)

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Offense(Index) = scrlOffense(Index).Value
    lblOffense(Index).Caption = GetCombatName(Index) & " Offense: " & scrlOffense(Index).Value

End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll(1) = scrlPaperdoll.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.Value
    Item(EditorIndex).Picture = scrlPic.Value
    Call EditorItem_BltItem
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPrice_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "Price: " & scrlPrice.Value
    Item(EditorIndex).MonetaryValue = scrlPrice.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilePiC.Caption = "Projectile Image: " & scrlProjectilePic.Value
    Item(EditorIndex).ProjecTile.Pic = scrlProjectilePic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProjectileRange_Change()
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = "Range: " & scrlProjectileRange.Value
    Item(EditorIndex).ProjecTile.Range = scrlProjectileRange.Value
End Sub

Private Sub scrlProjectileSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilesSpeed.Caption = "Speed: " & scrlProjectileSpeed.Value
    Item(EditorIndex).ProjecTile.Speed = scrlProjectileSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlReward_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblReward.Caption = GetSkillName(scrlReward.Value) & ":"
    txtReward.Text = Item(EditorIndex).SkillMakeRew(scrlReward.Value)

End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).Speed = scrlSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlWearingReq_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblWearingReq.Caption = GetSkillName(scrlWearingReq.Value) & ":"
    txtWearingReq.Text = Item(EditorIndex).SkillWearReq(scrlWearingReq.Value)

End Sub

Private Sub txtBonus_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If IsNumeric(txtBonus.Text) Then
        Item(EditorIndex).SkillBonus(scrlBonus.Value) = txtBonus.Text
    Else
        txtBonus.Text = "0"
        Item(EditorIndex).SkillBonus(scrlBonus.Value) = 0
    End If

End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Description = txtDesc.Text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMakingReq_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If IsNumeric(txtMakingReq.Text) Then
        Item(EditorIndex).SkillMakeReq(scrlMakingReq.Value) = txtMakingReq.Text
    Else
        txtMakingReq.Text = "0"
        Item(EditorIndex).SkillMakeReq(scrlMakingReq.Value) = 0
    End If

End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtReward_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If IsNumeric(txtReward.Text) Then
        Item(EditorIndex).SkillMakeRew(scrlReward.Value) = txtReward.Text
    Else
        txtReward.Text = "0"
        Item(EditorIndex).SkillMakeRew(scrlReward.Value) = 0
    End If

End Sub

Private Sub txtWearingReq_Change()

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If IsNumeric(txtWearingReq.Text) Then
        Item(EditorIndex).SkillWearReq(scrlWearingReq.Value) = txtWearingReq.Text
    Else
        txtWearingReq.Text = "0"
        Item(EditorIndex).SkillWearReq(scrlWearingReq.Value) = 0
    End If
    
End Sub
