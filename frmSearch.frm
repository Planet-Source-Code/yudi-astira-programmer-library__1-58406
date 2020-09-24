VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Load"
      Default         =   -1  'True
      Height          =   390
      Left            =   150
      TabIndex        =   11
      Top             =   3465
      Width           =   1260
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   1395
      Left            =   150
      TabIndex        =   3
      Top             =   75
      Width           =   5970
      Begin VB.CommandButton Command2 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5175
         TabIndex        =   10
         Top             =   960
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1545
         TabIndex        =   9
         Top             =   960
         Width           =   3540
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   570
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   165
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Source title"
         Height          =   195
         Left            =   615
         TabIndex        =   8
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Select type"
         Height          =   195
         Left            =   645
         TabIndex        =   6
         Top             =   645
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select categories"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   225
         Width           =   1230
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   390
      Left            =   4740
      TabIndex        =   2
      Top             =   3465
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Files"
      Height          =   1920
      Left            =   150
      TabIndex        =   0
      Top             =   1485
      Width           =   5955
      Begin MSComctlLib.ListView LV 
         Height          =   1635
         Left            =   105
         TabIndex        =   1
         Top             =   210
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   2884
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Author"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Categories"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Data type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date/Time"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3300
      Top             =   3390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nilai As String

Private Sub Combo1_Click()
    If Combo1.ListIndex = 0 Then
        Combo2.Enabled = False
        Exit Sub
    Else
        Combo2.Enabled = True
    End If
            
    TampilTipe
End Sub
Sub TampilTipe()
    If RS.State = 1 Then RS.Close
    
    RS.Open "SELECT * FROM TBL_TIPE WHERE CATEGORIES='" & Combo1.List(Combo1.ListIndex) & "'", DB, adOpenKeyset, adLockOptimistic
    
    With Combo2
        
        .Clear
        .AddItem "All tipe"
        
        While Not RS.EOF
            .AddItem RS!TIPE
            .ItemData(.NewIndex) = RS!ID_TIPE
            RS.MoveNext
        Wend
        
        .ListIndex = 0
        
        RS.Close
        
    End With
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim ITM As ListItem
    
    LV.ListItems.Clear
    If RS.State = 1 Then RS.Close
        
    If Combo1.ListCount = 1 Or (Combo2.ListCount = 1 And Combo1.ListIndex > 0) Then Exit Sub
    
    If Combo1.ListIndex = 0 Then
        RS.Open "SELECT * FROM QUERY1 WHERE JUDUL LIKE'" & Text1.Text & "%'", DB, adOpenKeyset, adLockOptimistic
        'MsgBox "SDF"
    ElseIf Combo1.ListIndex > 0 And Combo2.ListIndex > 0 Then
        RS.Open "SELECT * FROM QUERY1 WHERE JUDUL LIKE'" & Text1.Text & "%' AND CATEGORIES='" & _
            Combo1.List(Combo1.ListIndex) & "' AND ID_TIPE=" & Combo2.ItemData(Combo2.ListIndex), DB, adOpenKeyset, adLockOptimistic
        
    Else
        RS.Open "SELECT * FROM QUERY1 WHERE JUDUL LIKE'" & Text1.Text & "%' AND CATEGORIES='" & _
            Combo1.List(Combo1.ListIndex) & "'", DB, adOpenKeyset, adLockOptimistic
    End If
        LV.ListItems.Clear
        While Not RS.EOF
            Set ITM = LV.ListItems.Add(1, Chr(RS!ID_CODE), , 1, 1)
            ITM.Text = RS!JUDUL
            ITM.SubItems(1) = RS!AUTHOR
            ITM.SubItems(2) = RS!CATEGORIES
            ITM.SubItems(3) = RS!TIPE
            ITM.ListSubItems(3).Key = Chr(RS!ID_TIPE)
            ITM.SubItems(4) = RS!TANGGAL
            RS.MoveNext
        Wend
        
        RS.Close
End Sub

Private Sub Command3_Click()
    If LV.ListItems.Count = 0 Then Exit Sub
    Pekerjaan = "open"
    
    Nilai = LV.SelectedItem.Text
    ID_KODE = Asc(LV.SelectedItem.Key)
    PesanCat = LV.ListItems(LV.SelectedItem.Index).ListSubItems(2).Text
    PesanTip = Asc(LV.ListItems(LV.SelectedItem.Index).ListSubItems(3).Key)
      
    Unload Me
    
    Dim FrmD As frmDocument
    
    Set FrmD = New frmDocument
   
    FrmD.Caption = "Title : " & Nilai
    FrmD.Show
End Sub

Private Sub Form_Load()
    TampilCategories
    
End Sub
Sub TampilCategories()
With Combo1
    .Clear
    
    If RS.State = 1 Then RS.Close
    
    RS.Open "SELECT * FROM TBL_CATEGORIES", DB, adOpenKeyset, adLockOptimistic
    
    RS.Requery
        
        .AddItem "All Categories"
    While Not RS.EOF
        .AddItem RS!CATEGORIES
'        .ItemData(.NewIndex) = RS!ID_CATEGORIES
        RS.MoveNext
    Wend
    
    .ListIndex = 0
    
    RS.Close

End With
    
End Sub

Private Sub LV_DblClick()
    Command3_Click
End Sub
