VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change categories"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "frmCategories.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1020
      Top             =   1425
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
            Picture         =   "frmCategories.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategories.frx":07DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2325
      Left            =   2475
      TabIndex        =   0
      Top             =   60
      Width           =   4275
      Begin VB.CommandButton Command4 
         Caption         =   "&Add"
         Height          =   390
         Left            =   2490
         TabIndex        =   10
         Top             =   1830
         Width           =   810
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete"
         Height          =   390
         Left            =   3360
         TabIndex        =   8
         Top             =   1830
         Width           =   810
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   390
         Left            =   735
         TabIndex        =   6
         Top             =   1830
         Width           =   810
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
         Height          =   390
         Left            =   1620
         TabIndex        =   5
         Top             =   1830
         Width           =   810
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   750
         TabIndex        =   4
         Top             =   1230
         Width           =   3420
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   1980
      End
      Begin VB.Line Line1 
         X1              =   255
         X2              =   4170
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Please insert new categories name"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   300
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "After"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1305
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Before"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   780
         Width           =   465
      End
   End
   Begin MSComctlLib.TreeView TV 
      Height          =   2220
      Left            =   75
      TabIndex        =   9
      Top             =   150
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   3916
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
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Kode As String

Private Sub Command1_Click()
    On Error GoTo hand
    
    Dim h As String
    
    If Text1.Text = "" Then MsgBox "Please select the categories want you Edit", vbExclamation, "Delete": Exit Sub
    If Text2.Text = "" Then MsgBox "Please insert new name for " & Text1.Text, , "Error": Exit Sub
    
    h = Text2.Text
    h = UCase(Mid(h, 1, 1)) & LCase(Mid(h, 2, Len(h)))
    
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM TBL_CATEGORIES WHERE CATEGORIES='" & Kode & "'", DB, adOpenKeyset, adLockOptimistic
      RS!CATEGORIES = h
    RS.Update
    RS.Close
    
    Text1.Text = ""
    Text2.Text = ""
    
    TampilCategories
    
    Exit Sub
hand:
    If Err.Number = -2147467259 Then
    MsgBox "Sorry ..." & Chr(13) & "You try to insert new data." & Chr(13) & _
            "But the data have duplicate", vbExclamation, "Error"
    End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If Text1.Text = "" Then MsgBox "Please select the categories want you delete", vbExclamation, "Delete": Exit Sub
        
    If MsgBox("Are you sure to delete " & Text1.Text & " categories?" & _
    Chr(10) & Chr(10) & "Please remember, if you delete the " & _
    Text1.Text & " categories, so all source code for " & _
    Text1.Text & " will LOST...!", vbExclamation + vbYesNo + vbDefaultButton2, "Warning") = vbNo Then Exit Sub
    
    DB.Execute "DELETE FROM TBL_CATEGORIES WHERE CATEGORIES='" & Kode & "'"
    
    Text1.Text = ""
    Text2.Text = ""
    
    TampilCategories
End Sub

Private Sub Command4_Click()
    On Error GoTo hand
    Dim h As String
    h = InputBox("Please insert new categories", "Add categories")
    If h <> "" Then
        h = UCase(Mid(h, 1, 1)) & LCase(Mid(h, 2, Len(h)))
        
        If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM TBL_CATEGORIES", DB, adOpenKeyset, adLockOptimistic
            RS.AddNew
            RS!CATEGORIES = h
            RS.Update
        RS.Close
        
        TampilCategories
    End If
    Exit Sub
    
hand:
    If Err.Number = -2147467259 Then
    MsgBox "Sorry ..." & Chr(13) & "You try to insert new data." & Chr(13) & _
            "But the data have duplicate", vbExclamation, "Error"
    End If
    Exit Sub
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
        
        Text1.Text = TV.SelectedItem.Text
        
        Kode = TV.SelectedItem.Text
        
    End With
End Sub
