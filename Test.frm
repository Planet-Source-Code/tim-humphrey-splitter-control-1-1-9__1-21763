VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Splitter Control Demo"
   ClientHeight    =   6765
   ClientLeft      =   1770
   ClientTop       =   2250
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   8505
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   42
      Top             =   6480
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4683
            MinWidth        =   4683
            Text            =   "CurrRatioFromTop: "
            TextSave        =   "CurrRatioFromTop: "
            Key             =   "CurrRatioFromTop"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9789
            MinWidth        =   3201
            Text            =   "CurrSplitterPos: "
            TextSave        =   "CurrSplitterPos: "
            Key             =   "CurrSplitterPos"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraControlPanel 
      Caption         =   "Splitter Control Panel"
      Height          =   6255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3735
      Begin VB.Frame fraControls 
         BorderStyle     =   0  'None
         Caption         =   "Behavior"
         Height          =   1215
         Index           =   1
         Left            =   960
         TabIndex        =   45
         Top             =   3240
         Visible         =   0   'False
         Width           =   2655
         Begin VB.ComboBox cmbLiveUpdate 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox cmbMaintain 
            Height          =   315
            ItemData        =   "Test.frx":0000
            Left            =   1440
            List            =   "Test.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbAllowResize 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Maintain"
            Height          =   255
            Left            =   0
            TabIndex        =   51
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "LiveUpdate"
            Height          =   255
            Left            =   0
            TabIndex        =   50
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "AllowResize"
            Height          =   255
            Left            =   0
            TabIndex        =   49
            Top             =   30
            Width           =   1455
         End
      End
      Begin VB.ComboBox cmbSplitter 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame fraControls 
         BorderStyle     =   0  'None
         Caption         =   "Appearance"
         Height          =   1455
         Index           =   0
         Left            =   960
         TabIndex        =   31
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
         Begin VB.ComboBox cmbBorderStyle 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   0
            Width           =   1215
         End
         Begin VB.ComboBox cmbSplitterAppearance 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox cmbSplitterBorder 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdSplitterColor 
            Caption         =   "Color"
            Height          =   315
            Left            =   1800
            TabIndex        =   32
            Top             =   1080
            Width           =   855
         End
         Begin VB.Shape shpSplitterColor 
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   1440
            Top             =   1110
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "BorderStyle"
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   30
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "SplitterAppearance"
            Height          =   255
            Left            =   0
            TabIndex        =   38
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "SplitterBorder"
            Height          =   255
            Left            =   0
            TabIndex        =   37
            Top             =   750
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "SplitterColor"
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   1110
            Width           =   1455
         End
      End
      Begin VB.Frame fraControls 
         BorderStyle     =   0  'None
         Caption         =   "Misc"
         Height          =   735
         Index           =   2
         Left            =   960
         TabIndex        =   26
         Top             =   4560
         Visible         =   0   'False
         Width           =   2655
         Begin VB.TextBox txtChild2 
            Height          =   285
            Left            =   1440
            TabIndex        =   28
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtChild1 
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Child1"
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   15
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Child2"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   375
            Width           =   1455
         End
      End
      Begin VB.Frame fraControls 
         BorderStyle     =   0  'None
         Caption         =   "Position"
         Height          =   3255
         Index           =   3
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
         Begin VB.TextBox txtSplitterSize 
            Height          =   285
            Left            =   1440
            TabIndex        =   16
            Top             =   2880
            Width           =   1215
         End
         Begin VB.TextBox txtSplitterPos 
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtRatioFromTop 
            Height          =   285
            Left            =   1440
            TabIndex        =   14
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtMinSizeAux 
            Height          =   285
            Left            =   1440
            TabIndex        =   13
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtMinSize2 
            Height          =   285
            Left            =   1440
            TabIndex        =   12
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtMinSize1 
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtMaxSize 
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Top             =   0
            Width           =   1215
         End
         Begin VB.ComboBox cmbOrientation 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1800
            Width           =   1215
         End
         Begin VB.ComboBox cmbMaxSizeAppliesTo 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "RatioFromTop"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   2175
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Orientation"
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   1830
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "MinSizeAux"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   1455
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "MinSize2"
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   1095
            Width           =   1455
         End
         Begin VB.Label Label13 
            Caption         =   "MinSize1"
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   735
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "MaxSizeAppliesTo"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   390
            Width           =   1455
         End
         Begin VB.Label Label11 
            Caption         =   "SplitterPos"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   2535
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "SplitterSize"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   2895
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "MaxSize"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   15
            Width           =   1455
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3855
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   6800
         HotTracking     =   -1  'True
         Separators      =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Appearance"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Behavior"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Misc"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Position"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label Label19 
         Caption         =   $"Test.frx":0004
         Height          =   855
         Left            =   240
         TabIndex        =   44
         Top             =   5280
         Width           =   3255
      End
      Begin VB.Label Label18 
         Caption         =   "-- Press enter in text boxes to accept values, escape to cancel."
         Height          =   495
         Left            =   240
         TabIndex        =   43
         Top             =   4800
         Width           =   3255
      End
      Begin VB.Label Label20 
         Caption         =   "Splitter"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   390
         Width           =   615
      End
   End
   Begin Project1.Splitter Splitter1 
      Height          =   3225
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5689
      SplitterPos     =   1790
      RatioFromTop    =   0.41
      Child1          =   "treeview1"
      Child2          =   "splitter2"
      Begin Project1.Splitter Splitter2 
         Height          =   3225
         Left            =   1865
         TabIndex        =   2
         Top             =   0
         Width           =   2590
         _ExtentX        =   4577
         _ExtentY        =   5689
         Orientation     =   1
         SplitterPos     =   1001
         RatioFromTop    =   0.322
         Child1          =   "listview1"
         Child2          =   "frame2"
         MaxSize         =   2000
         MinSize2        =   1000
         MinSizeAux      =   2000
         Begin MSComctlLib.ListView ListView1 
            Height          =   1030
            Left            =   -15
            TabIndex        =   4
            Top             =   -15
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   1826
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Column 1"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Column 2"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   2150
            Left            =   0
            TabIndex        =   3
            Top             =   1075
            Width           =   2595
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3225
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1790
         _ExtentX        =   3149
         _ExtentY        =   5689
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim gsplitCurr As Splitter
Dim gintCurrTab As Integer

Function HandleKeyPress(txtBox As TextBox, KeyAscii As Integer, property As String) As Boolean
    Dim result As Boolean
    
    result = False
    Select Case KeyAscii
    Case vbKeyReturn
        txtBox.SelStart = 0
        txtBox.SelLength = Len(txtBox.Text)
        property = txtBox.Text
        
        KeyAscii = 0
        result = True
    Case vbKeyEscape
        txtBox.Text = property
        txtBox.SelStart = 0
        txtBox.SelLength = Len(txtBox.Text)
        
        KeyAscii = 0
        result = True
    End Select
    
    HandleKeyPress = result
End Function

Sub SetSplitterControlPanel()
    With gsplitCurr
        'Appearance
        cmbBorderStyle.ListIndex = .BorderStyle
        cmbSplitterAppearance.ListIndex = .SplitterAppearance
        cmbSplitterBorder.ListIndex = .SplitterBorder
        shpSplitterColor.FillColor = .SplitterColor
        
        'Behavior
        If gsplitCurr.AllowResize Then
            cmbAllowResize.ListIndex = 1
        Else
            cmbAllowResize.ListIndex = 0
        End If
        
        If gsplitCurr.LiveUpdate Then
            cmbLiveUpdate.ListIndex = 1
        Else
            cmbLiveUpdate.ListIndex = 0
        End If
        
        cmbMaintain.ListIndex = .Maintain
        
        'Misc
        txtChild1.Text = .Child1
        txtChild2.Text = .Child2
        
        'Position
        txtMaxSize.Text = .MaxSize
        cmbMaxSizeAppliesTo.ListIndex = .MaxSizeAppliesTo
        txtMinSize1.Text = .MinSize1
        txtMinSize2.Text = .MinSize2
        txtMinSizeAux.Text = .MinSizeAux
        cmbOrientation.ListIndex = .Orientation
        txtRatioFromTop.Text = .RatioFromTop
        txtSplitterPos.Text = .SplitterPos
        txtSplitterSize.Text = .SplitterSize
    End With
End Sub

Sub UpdateSplitterCoords(number As Integer)
    If number = cmbSplitter.ListIndex + 1 Then
        txtRatioFromTop.Text = gsplitCurr.RatioFromTop
        txtSplitterPos.Text = gsplitCurr.SplitterPos
        
        StatusBar1.Panels("CurrRatioFromTop").Text = "CurrRatioFromTop: " & gsplitCurr.CurrRatioFromTop
        StatusBar1.Panels("CurrSplitterPos").Text = "CurrSplitterPos: " & gsplitCurr.CurrSplitterPos
    End If
End Sub

Private Sub cmbAllowResize_Click()
    Select Case cmbAllowResize.ListIndex
    Case 0
        gsplitCurr.AllowResize = False
    Case 1
        gsplitCurr.AllowResize = True
    End Select
End Sub

Private Sub cmbBorderStyle_Click()
    gsplitCurr.BorderStyle = cmbBorderStyle.ListIndex
End Sub

Private Sub cmbLiveUpdate_Click()
    Select Case cmbLiveUpdate.ListIndex
    Case 0
        gsplitCurr.LiveUpdate = False
    Case 1
        gsplitCurr.LiveUpdate = True
    End Select
End Sub

Private Sub cmbMaintain_Click()
    gsplitCurr.Maintain = cmbMaintain.ListIndex
End Sub

Private Sub cmbMaxSizeAppliesTo_Click()
    gsplitCurr.MaxSizeAppliesTo = cmbMaxSizeAppliesTo.ListIndex
End Sub

Private Sub cmbOrientation_Click()
    gsplitCurr.Orientation = cmbOrientation.ListIndex
End Sub

Private Sub cmbSplitter_Click()
    Select Case cmbSplitter.ListIndex
    Case 0
        Set gsplitCurr = Splitter1
        SetSplitterControlPanel
    Case 1
        Set gsplitCurr = Splitter2
        SetSplitterControlPanel
    End Select
End Sub

Private Sub cmbSplitterAppearance_Click()
    gsplitCurr.SplitterAppearance = cmbSplitterAppearance.ListIndex
End Sub

Private Sub cmbSplitterBorder_Click()
    gsplitCurr.SplitterBorder = cmbSplitterBorder.ListIndex
End Sub

Private Sub cmdSplitterColor_Click()
    On Error Resume Next
    
    CommonDialog1.Flags = cdlCCFullOpen + cdlCCRGBInit
    CommonDialog1.Color = gsplitCurr.SplitterColor
    CommonDialog1.ShowColor
    If Err.number = 0 Then
        gsplitCurr.SplitterColor = CommonDialog1.Color
        shpSplitterColor.FillColor = gsplitCurr.SplitterColor
    Else
        If Err.number <> cdlCancel Then
            With Err
                .Raise .number, .Source, .Description, .HelpFile, .HelpContext
            End With
        End If
    End If
End Sub

Private Sub Form_Load()
    '----- Populate Splitter Control Panel controls
    'Splitter combo
    With cmbSplitter
        .AddItem "Splitter1"
        .AddItem "Splitter2"
    End With
    
    'Appearance
    With cmbBorderStyle
        .AddItem "None"
        .AddItem "Fixed Single"
    End With
    
    With cmbSplitterAppearance
        .AddItem "Flat"
        .AddItem "3D"
    End With
    
    With cmbSplitterBorder
        .AddItem "None"
        .AddItem "Fixed Single"
    End With
    
    'Behavior
    With cmbAllowResize
        .AddItem "False"
        .AddItem "True"
    End With
    
    With cmbLiveUpdate
        .AddItem "False"
        .AddItem "True"
    End With
    
    With cmbMaintain
        .AddItem "MN_POS"
        .AddItem "MN_RATIO"
    End With
    
    'Position
    With cmbMaxSizeAppliesTo
        .AddItem "MX_CHILD1"
        .AddItem "MX_CHILD2"
    End With
    
    With cmbOrientation
        .AddItem "OC_HORIZONTAL"
        .AddItem "OC_VERTICAL"
    End With
    
    '----- Initialize Splitter Control Panel controls
    cmbSplitter.ListIndex = 0
    SetSplitterControlPanel
    gintCurrTab = 1
    TabStrip1_Click
    
    '----- Populate Splitter contained controls
    TreeView1.Nodes.Add , , , "Node 1"
    TreeView1.Nodes.Add 1, tvwChild, , "Node 2"
    TreeView1.Nodes.Add 1, tvwChild, , "Node 3"
    TreeView1.Nodes.Add , , , "Node 4"
    TreeView1.Nodes(1).Expanded = True
    
    ListView1.ListItems.Add , , "Item 1"
    ListView1.ListItems(1).SubItems(1) = "Subitem 1"
    ListView1.ListItems.Add , , "Item 2"
    ListView1.ListItems(2).SubItems(1) = "Subitem 2"
End Sub

Private Sub Form_Resize()
    Dim newWidth As Integer
    Dim newHeight As Integer
    
    With Splitter1
        newWidth = Form1.ScaleWidth - .Left
        If newWidth < 0 Then
            newWidth = 0
        End If
        
        newHeight = Form1.ScaleHeight - StatusBar1.Height - .Top
        If newHeight < 0 Then
            newHeight = 0
        End If
        
        .Move .Left, .Top, newWidth, newHeight
    End With
End Sub

Private Sub Splitter1_Resize()
    UpdateSplitterCoords 1
End Sub

Private Sub Splitter2_Resize()
    UpdateSplitterCoords 2
End Sub

Private Sub TabStrip1_Click()
    With TabStrip1
        fraControls(gintCurrTab).Visible = False
        gintCurrTab = .SelectedItem.Index - 1
        fraControls(gintCurrTab).Move .ClientLeft + 120, .ClientTop + 120, .ClientWidth - 120, .ClientHeight - 120
        fraControls(gintCurrTab).Visible = True
    End With
End Sub

Private Sub txtChild1_KeyPress(KeyAscii As Integer)
    Dim property As String
    
    With gsplitCurr
        property = .Child1
        If HandleKeyPress(txtChild1, KeyAscii, property) Then
            .Child1 = property
            txtChild1.Text = .Child1
        End If
    End With
End Sub

Private Sub txtChild2_KeyPress(KeyAscii As Integer)
    Dim property As String
    
    With gsplitCurr
        property = .Child2
        If HandleKeyPress(txtChild2, KeyAscii, property) Then
            .Child2 = property
            txtChild2.Text = .Child2
        End If
    End With
End Sub

Private Sub txtMaxSize_KeyPress(KeyAscii As Integer)
    Dim property As String
    
    With gsplitCurr
        property = .MaxSize
        If HandleKeyPress(txtMaxSize, KeyAscii, property) Then
            .MaxSize = property
            txtMaxSize.Text = .MaxSize
        End If
    End With
End Sub

Private Sub txtMinSize1_KeyPress(KeyAscii As Integer)
    Dim property As String
    
    With gsplitCurr
        property = .MinSize1
        If HandleKeyPress(txtMinSize1, KeyAscii, property) Then
            .MinSize1 = property
            txtMinSize1.Text = .MinSize1
        End If
    End With
End Sub

Private Sub txtMinSize2_KeyPress(KeyAscii As Integer)
    Dim property As String
    
    With gsplitCurr
        property = .MinSize2
        If HandleKeyPress(txtMinSize2, KeyAscii, property) Then
            .MinSize2 = property
            txtMinSize2.Text = .MinSize2
        End If
    End With
End Sub

Private Sub txtMinSizeAux_KeyPress(KeyAscii As Integer)
    Dim property As String
    
    With gsplitCurr
        property = .MinSizeAux
        If HandleKeyPress(txtMinSizeAux, KeyAscii, property) Then
            .MinSizeAux = property
            txtMinSizeAux = .MinSizeAux
        End If
    End With
End Sub

Private Sub txtRatioFromTop_KeyPress(KeyAscii As Integer)
    Dim property As String
    
    With gsplitCurr
        property = .RatioFromTop
        If HandleKeyPress(txtRatioFromTop, KeyAscii, property) Then
            .RatioFromTop = property
            txtRatioFromTop.Text = .RatioFromTop
            txtSplitterPos.Text = .SplitterPos
        End If
    End With
End Sub

Private Sub txtSplitterPos_KeyPress(KeyAscii As Integer)
    Dim property As String
    
    With gsplitCurr
        property = .SplitterPos
        If HandleKeyPress(txtSplitterPos, KeyAscii, property) Then
            .SplitterPos = property
            txtSplitterPos.Text = .SplitterPos
            txtRatioFromTop.Text = .RatioFromTop
        End If
    End With
End Sub

Private Sub txtSplitterSize_KeyPress(KeyAscii As Integer)
    Dim property As String
    
    With gsplitCurr
        property = .SplitterSize
        If HandleKeyPress(txtSplitterSize, KeyAscii, property) Then
            .SplitterSize = property
            txtSplitterSize.Text = .SplitterSize
        End If
    End With
End Sub
