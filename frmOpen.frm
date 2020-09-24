VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open dictionary"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   Icon            =   "frmOpen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   4410
      TabIndex        =   10
      Top             =   5040
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Load"
      Default         =   -1  'True
      Height          =   390
      Left            =   5700
      TabIndex        =   9
      Top             =   5040
      Width           =   1200
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   2190
      Top             =   5235
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen.frx":0666
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   150
      TabIndex        =   5
      Top             =   15
      Width           =   6750
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   5535
         TabIndex        =   8
         Top             =   195
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList4"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Details"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Large Icon"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "List"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   225
         Width           =   3465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Select Categories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   1680
      End
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1440
      Top             =   5220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen.frx":0778
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLine2 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   135
      ScaleHeight     =   30
      ScaleWidth      =   6765
      TabIndex        =   3
      Top             =   4935
      Width           =   6765
      Begin VB.PictureBox Picture1 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   10
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   11880
         TabIndex        =   4
         Top             =   10
         Width           =   11880
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000005&
      Height          =   4080
      Left            =   150
      ScaleHeight     =   4020
      ScaleWidth      =   2385
      TabIndex        =   0
      Top             =   810
      Width           =   2450
      Begin MSComctlLib.TreeView TV 
         Height          =   4020
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   7091
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "ImageList2"
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   165
      Top             =   5205
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen.frx":0BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpen.frx":0F64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   795
      Top             =   5190
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
            Picture         =   "frmOpen.frx":12FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4110
      Left            =   2640
      TabIndex        =   2
      Top             =   795
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   7250
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList3"
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
      NumItems        =   3
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
         Text            =   "Date/Time"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nilai As String
Private Sub Combo1_Click()
    LV.ListItems.Clear
    TampilTipe
End Sub

Private Sub Command1_Click()
    If LV.ListItems.Count = 0 Then Exit Sub
    Pekerjaan = "open"
    
    Nilai = LV.SelectedItem.Text
    ID_KODE = Asc(LV.SelectedItem.Key)
    PesanCat = Combo1.Text
    PesanTip = TV.SelectedItem.Text
    Unload Me
    
    Dim FrmD As frmDocument
    
    Set FrmD = New frmDocument
   
    FrmD.Caption = "Title : " & Nilai
    FrmD.Show
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    TampilCategories
    

End Sub

Sub TampilCategories()

    With Combo1
        .Clear
        If RS.State = 1 Then RS.Close
        
        RS.Open "SELECT * FROM TBL_CATEGORIES", DB, adOpenKeyset, adLockOptimistic
        
        While Not RS.EOF
            .AddItem RS!CATEGORIES
            '.ItemData(.NewIndex) = RS!ID_CATEGORIES
            RS.MoveNext
        Wend
        
        RS.Close
    End With
        
End Sub

Sub TampilTipe()
       
        If RS.State = 1 Then RS.Close
        
    With TV.Nodes
        .Clear
        
        RS.Open "SELECT * FROM TBL_TIPE WHERE CATEGORIES='" & Combo1.List(Combo1.ListIndex) & "'", DB, adOpenKeyset, adLockOptimistic
        While Not RS.EOF
            .Add , , Chr(RS!ID_TIPE), RS!TIPE, 1, 2
            RS.MoveNext
        Wend
        
        RS.Close
    End With
    
End Sub

Private Sub LV_DblClick()
    Command1_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: LV.View = lvwReport
            
        Case 2: LV.View = lvwIcon
            
        Case 3: LV.View = lvwList
    End Select
        
End Sub

Private Sub TV_Click()
   ' On Error Resume Next
    Dim ITM As ListItem
    With TV.Nodes
        
        If .Count = 0 Then Exit Sub
        
        If RS.State = 1 Then RS.Close
        
        RS.Open "SELECT * FROM TBL_CODE WHERE ID_TIPE=" & _
            Asc(TV.SelectedItem.Key), DB, adOpenKeyset, adLockOptimistic
            LV.ListItems.Clear
        While Not RS.EOF
            Set ITM = LV.ListItems.Add(1, Chr(RS!ID_CODE), , 1, 1)
            ITM.Text = RS!JUDUL
            ITM.SubItems(1) = RS!AUTHOR
            ITM.SubItems(2) = RS!TANGGAL
            RS.MoveNext
        Wend
        
        RS.Close
    End With
End Sub
