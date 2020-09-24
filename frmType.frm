VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change data type"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   Icon            =   "frmType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   2535
      TabIndex        =   1
      Top             =   60
      Width           =   3975
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   735
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   960
         TabIndex        =   10
         Top             =   1575
         Width           =   2805
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
         Height          =   390
         Left            =   1230
         TabIndex        =   6
         Top             =   2100
         Width           =   810
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   390
         Left            =   345
         TabIndex        =   5
         Top             =   2100
         Width           =   810
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete"
         Height          =   390
         Left            =   2970
         TabIndex        =   4
         Top             =   2100
         Width           =   810
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Add"
         Height          =   390
         Left            =   2100
         TabIndex        =   3
         Top             =   2100
         Width           =   810
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1155
         Width           =   2820
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Categories"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   795
         Width           =   750
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   3765
         Y1              =   2025
         Y2              =   2025
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Please insert new data type name"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   420
         Width           =   2385
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "After"
         Height          =   195
         Left            =   165
         TabIndex        =   8
         Top             =   1680
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Before"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   1200
         Width           =   465
      End
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   2490
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   4392
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList2"
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
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1530
      Top             =   2115
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
            Picture         =   "frmType.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmType.frx":07DC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Kode As String

Private Sub Command1_Click()
    Dim h As String
    
    If Combo1.ListIndex < 0 Then MsgBox "You have to select the data type want you edit", vbExclamation, "Error": Exit Sub
    If Text1.Text = "" Then MsgBox "You have to enter new name for data type " & Combo1.Text, vbExclamation, "Error": Exit Sub
    
    h = Text1.Text
    h = UCase(Mid(h, 1, 1)) & LCase(Mid(h, 2, Len(h)))
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM TBL_TIPE WHERE CATEGORIES='" & Kode & _
        "' AND ID_TIPE=" & Combo1.ItemData(Combo1.ListIndex)
        RS!TIPE = h
    RS.Update
    RS.Close
    
    TV_Click
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If Combo1.Text = "" Then MsgBox "Sorry you have to select the data type want you delete !", vbExclamation, "Error": Exit Sub
    
    If MsgBox("Are you sure want to delete " & Combo1.Text & " data type?" & _
    Chr(10) & Chr(10) & "Please remember, if you delete this data type, all source code for this will lost", vbExclamation + vbYesNo + vbDefaultButton2, "Warning") = vbNo Then Exit Sub
    
    DB.Execute "DELETE FROM TBL_TIPE WHERE CATEGORIES='" & Kode & "' AND ID_TIPE=" & Combo1.ItemData(Combo1.ListIndex)
    
    TV_Click
    
End Sub

Private Sub Command4_Click()
    On Error GoTo hand
    Dim h As String
    If Text2.Text = "" Then MsgBox "Please select the categories first", , Title: Exit Sub
    h = InputBox("Please insert new data type for " & Text2.Text, "Add new data type")
    If h <> "" Then
        h = UCase(Mid(h, 1, 1)) & LCase(Mid(h, 2, Len(h)))
        RS.Open "SELECT * FROM TBL_TIPE"
        RS.AddNew
            RS!CATEGORIES = Kode
            RS!TIPE = h
        RS.Update
        RS.Close
        
        TV_Click
    End If
    Exit Sub
hand:
    MsgBox "Sorry ... found an error, bellow the Description" & Chr(10) & Chr(10) & Err.Description, vbExclamation, "Error Number : " & Err.Number
End Sub

Private Sub Form_Load()
    TampilCategories
End Sub
Sub TampilCategories()
With TV.Nodes
    .Clear
    
    If RS.State = 1 Then RS.Close
    
    RS.Open "SELECT * FROM TBL_CATEGORIES", DB, adOpenKeyset, adLockOptimistic
    
    RS.Requery
    
    While Not RS.EOF
        .Add , , , RS!CATEGORIES, 1, 2
        RS.MoveNext
    Wend
    
    RS.Close

End With
End Sub


Private Sub TV_Click()
        With TV.Nodes
        
        If .Count = 0 Then Exit Sub
        
        If RS.State = 1 Then RS.Close
        
        RS.Open "SELECT * FROM TBL_TIPE WHERE CATEGORIES='" & TV.SelectedItem.Text & "'", DB, adOpenKeyset, adLockOptimistic
        
        Combo1.Clear
        
        While Not RS.EOF
            Combo1.AddItem RS!TIPE
            Combo1.ItemData(Combo1.NewIndex) = RS!ID_TIPE
            RS.MoveNext
        Wend
        If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
        RS.Close
        
        Text2.Text = TV.SelectedItem.Text
        
        Kode = TV.SelectedItem.Text
        
    End With
End Sub
