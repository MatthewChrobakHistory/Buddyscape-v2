VERSION 5.00
Begin VB.Form frmEditor_Shop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop Editor"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
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
   Icon            =   "frmEditor_Shop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   743
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Shop Properties"
      Height          =   6615
      Left            =   3360
      TabIndex        =   5
      Top             =   120
      Width           =   7695
      Begin VB.HScrollBar scrlCurrency 
         Height          =   255
         Left            =   4800
         TabIndex        =   43
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox chkOnlyBuyItemsInStock 
         Caption         =   "Buy Stock Only"
         Height          =   180
         Left            =   3000
         TabIndex        =   41
         Top             =   840
         Width           =   1815
      End
      Begin VB.HScrollBar scrlGUI 
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtMaxStock 
         Height          =   270
         Left            =   6240
         TabIndex        =   38
         Text            =   "infinity"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtBuyVerb 
         Height          =   285
         Left            =   5760
         MaxLength       =   12
         TabIndex        =   35
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkGiveAndRequireXP 
         Caption         =   "Give and require xp?"
         Height          =   180
         Left            =   5280
         TabIndex        =   34
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   4560
         TabIndex        =   31
         Text            =   "1"
         Top             =   3240
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Index           =   4
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   4560
         TabIndex        =   27
         Text            =   "1"
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Index           =   3
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   4560
         TabIndex        =   23
         Text            =   "1"
         Top             =   2520
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Index           =   2
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtCostValue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   4560
         TabIndex        =   19
         Text            =   "1"
         Top             =   2160
         Width           =   615
      End
      Begin VB.ComboBox cmbCostItem 
         Height          =   300
         Index           =   1
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CommandButton cmdDeleteTrade 
         Caption         =   "Delete"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   3720
         Width           =   2415
      End
      Begin VB.HScrollBar scrlBuy 
         Height          =   255
         Left            =   120
         Max             =   1000
         Min             =   1
         TabIndex        =   15
         Top             =   840
         Value           =   100
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   720
         MaxLength       =   12
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.ListBox lstTradeItem 
         Height          =   2400
         ItemData        =   "frmEditor_Shop.frx":3332
         Left            =   120
         List            =   "frmEditor_Shop.frx":334E
         TabIndex        =   9
         Top             =   4080
         Width           =   7335
      End
      Begin VB.ComboBox cmbItem 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtItemValue 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Text            =   "1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label lblCurrency 
         Caption         =   "Currency:"
         Height          =   255
         Left            =   4800
         TabIndex        =   42
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblGUI 
         Caption         =   "GUI: 0"
         Height          =   255
         Left            =   2760
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "MaxStock:"
         Height          =   255
         Left            =   5280
         TabIndex        =   37
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Verb:"
         Height          =   180
         Left            =   5160
         TabIndex        =   36
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Index           =   4
         Left            =   3960
         TabIndex        =   33
         Top             =   3240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   3240
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Index           =   3
         Left            =   3960
         TabIndex        =   29
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Index           =   2
         Left            =   3960
         TabIndex        =   25
         Top             =   2520
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   2520
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Index           =   1
         Left            =   3960
         TabIndex        =   21
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label lblBuy 
         AutoSize        =   -1  'True
         Caption         =   "Buy Rate: 100%"
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   180
         Left            =   3960
         TabIndex        =   11
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Shop List"
      Height          =   6615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6180
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Change Array Size"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6840
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditor_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkGiveAndRequireXP_Click()

    Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).GiveAndRequireXP = chkGiveAndRequireXP.Value

End Sub

Private Sub chkOnlyBuyItemsInStock_Click()

    Shop(EditorIndex).OnlyBuyItemsInStock = chkOnlyBuyItemsInStock.Value

End Sub

Private Sub cmdSave_Click()
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If LenB(Trim$(txtName)) = 0 Then
        Call MsgBox("Name required.")
    Else
        Call ShopEditorOk
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ShopEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdUpdate_Click()
Dim Index As Long
Dim tmpPos As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    tmpPos = lstTradeItem.ListIndex
    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = cmbItem.ListIndex
        For i = 1 To MAX_SHOP_ITEM_COSTS
            .CostItem(i) = cmbCostItem(i).ListIndex
            .CostValue(i) = Val(txtCostValue(i).Text)
        Next
    End With
    UpdateShopTrade tmpPos
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdUpdate_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDeleteTrade_Click()
Dim Index As Long
Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Index = lstTradeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    With Shop(EditorIndex).TradeItem(Index)
        .Item = 0
        .Stock = 0
        For i = 1 To MAX_SHOP_ITEM_COSTS
            .CostItem(i) = 0
            .CostValue(i) = 0
        Next
    End With
    Call UpdateShopTrade
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDeleteTrade_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()

    scrlCurrency.Max = MAX_ITEMS

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ShopEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstTradeItem_Click()
Dim i As Long

    With Me
        .cmbItem.ListIndex = Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).Item
        .txtItemValue.Text = 1
        .chkGiveAndRequireXP.Value = Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).GiveAndRequireXP
        
        If Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).MaxStock = -255 Then
            .txtMaxStock.Text = "infinity"
        Else
            .txtMaxStock.Text = Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).MaxStock
        End If
        
        For i = 1 To MAX_SHOP_ITEM_COSTS
            .cmbCostItem(i).ListIndex = Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).CostItem(i)
            .txtCostValue(i).Text = Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).CostValue(i)
        Next
    End With
End Sub

Private Sub scrlBuy_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblBuy.Caption = "Buy Rate: " & scrlBuy.Value & "%"
    Shop(EditorIndex).BuyRate = scrlBuy.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlBuy_Change", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCurrency_Change()

    If scrlCurrency.Value < 1 Then scrlCurrency.Value = 1
    lblCurrency.Caption = "Currency: " & Trim$(Item(scrlCurrency.Value).Name)
    Shop(EditorIndex).ShopCurrency = scrlCurrency.Value

End Sub

Private Sub scrlGUI_Change()

    lblGUI.Caption = "GUI: " & scrlGUI.Value
    Shop(EditorIndex).InterfaceNum = scrlGUI.Value

End Sub

Private Sub txtBuyVerb_Change()

    Shop(EditorIndex).BuyVerb = txtBuyVerb.Text

End Sub

Private Sub txtMaxStock_Change()

    If IsNumeric(txtMaxStock.Text) Then
        If txtMaxStock.Text > 2147483647 Or txtMaxStock.Text < 1 Then
            txtMaxStock.Text = "infinity"
            Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).MaxStock = -255
        Else
            Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).MaxStock = txtMaxStock.Text
        End If
    Else
        txtMaxStock.Text = "infinity"
        Shop(EditorIndex).TradeItem(lstTradeItem.ListIndex + 1).MaxStock = -255
    End If
    

End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Shop(EditorIndex).Name = Trim$(txtName.Text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Shop(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Shop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
