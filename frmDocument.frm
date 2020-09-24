VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   6615
   Begin VB.Frame fr2 
      Height          =   3375
      Left            =   270
      TabIndex        =   7
      Top             =   405
      Width           =   6105
      Begin VB.PictureBox bingKai 
         BackColor       =   &H00FFC0C0&
         Height          =   2730
         Left            =   450
         ScaleHeight     =   2670
         ScaleWidth      =   5160
         TabIndex        =   8
         Top             =   345
         Width           =   5220
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   4410
            Picture         =   "frmDocument.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "New data type"
            Top             =   630
            Width           =   315
         End
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   4410
            Picture         =   "frmDocument.frx":07CC
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "New categories"
            Top             =   225
            Width           =   315
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Left            =   5055
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   630
            Visible         =   0   'False
            Width           =   3705
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   2145
            Width           =   2925
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1125
            TabIndex        =   13
            Top             =   1770
            Width           =   2910
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1125
            TabIndex        =   12
            Top             =   1380
            Width           =   2910
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1125
            TabIndex        =   11
            Top             =   1020
            Width           =   2910
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   630
            Width           =   3195
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1140
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   225
            Width           =   3195
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date / Time"
            Height          =   195
            Left            =   150
            TabIndex        =   21
            Top             =   2220
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail"
            Height          =   195
            Left            =   585
            TabIndex        =   20
            Top             =   1860
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
            Height          =   195
            Left            =   705
            TabIndex        =   19
            Top             =   1095
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Author"
            Height          =   195
            Left            =   555
            TabIndex        =   18
            Top             =   1455
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data type"
            Height          =   195
            Left            =   315
            TabIndex        =   17
            Top             =   705
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Categories"
            Height          =   195
            Left            =   270
            TabIndex        =   16
            Top             =   315
            Width           =   750
         End
      End
   End
   Begin VB.Frame fr1 
      Height          =   3345
      Left            =   3180
      TabIndex        =   5
      Top             =   -765
      Visible         =   0   'False
      Width           =   6090
      Begin RichTextLib.RichTextBox rtfText 
         Height          =   1995
         Left            =   135
         TabIndex        =   6
         Top             =   255
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   3519
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmDocument.frx":0B56
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   435
      Left            =   5370
      TabIndex        =   4
      Top             =   4080
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   135
      TabIndex        =   1
      Top             =   3930
      Width           =   1920
      Begin VB.CommandButton Command3 
         Caption         =   "&Delete"
         Height          =   390
         Left            =   1020
         TabIndex        =   3
         Top             =   180
         Width           =   810
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Edit"
         Height          =   405
         Left            =   75
         TabIndex        =   2
         Top             =   165
         Width           =   885
      End
   End
   Begin MSComctlLib.TabStrip tbS 
      Height          =   3810
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   6720
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&File description"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Code"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pesan As Long
Dim Nilai As String
Dim NilaiCat As String
Dim NilaiTip As String
Dim Hancurkan As Boolean

Sub KunciKontrol(En As Boolean)
    Combo1.Enabled = Not En
    Combo2.Enabled = Not En
    Text1.Locked = En
    Text2.Locked = En
    Text3.Locked = En
    rtfText.Locked = En
    Command4.Enabled = Not En
    Command6.Enabled = Not En
End Sub

Private Sub Combo1_Click()
    TampilTipe
End Sub

Private Sub Combo2_Click()
    Combo4.ListIndex = Combo2.ListIndex
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Public Sub Command2_Click()
    
    Select Case Command2.Caption
    Case "&Edit"
        
        KunciKontrol False
        Command2.Caption = "&Save"
        Command3.Caption = "&Cancel"
        
    Case "&Save"
        
        If Combo1.ListIndex < 0 Or Combo2.ListIndex < 0 Or _
            Text1.Text = "" Or Text2.Text = "" Or rtfText.Text = "" Then
            
            MsgBox "Your data is not complete", vbExclamation, "Warning"
            Exit Sub
            
        End If
        
        KunciKontrol True
        
        If RS.State = 1 Then RS.Close
        
        If Nilai = "open" Then
            RS.Open "SELECT * FROM TBL_CODE WHERE ID_CODE=" & Pesan, DB, adOpenKeyset, adLockOptimistic
             
             RS!ID_TIPE = Combo4.Text
             RS!JUDUL = Text1.Text
             RS!AUTHOR = Text2.Text
             RS!EMAIL = Text3.Text
             RS!TANGGAL = Date & " " & Time
             RS!CODE = rtfText.TextRTF
            
            RS.Update
        Else
        
            RS.Open "SELECT * FROM TBL_CODE", DB, adOpenKeyset, adLockOptimistic
            RS.AddNew
            
             RS!ID_TIPE = Combo4.Text
             RS!JUDUL = Text1.Text
             RS!AUTHOR = Text2.Text
             RS!EMAIL = Text3.Text
             RS!TANGGAL = Date & " " & Time
             RS!CODE = rtfText.TextRTF
            
            RS.Update
            
            Nilai = "open"
            
            If RS.State = 1 Then RS.Close
            RS.Open "SELECT ID_CODE FROM TBL_CODE WHERE ID_TIPE=" & Combo4.Text, DB, adOpenKeyset, adLockOptimistic
            
            RS.MoveLast
            
            Pesan = RS!ID_CODE
            
        End If
        RS.Close
            
        Me.Caption = "Title : " & Text1.Text
        
        Command2.Caption = "&Edit"
        
        Command3.Caption = "&Delete"
        
    End Select
End Sub

Private Sub Command3_Click()
    Dim h As Long
    Select Case Command3.Caption
    Case "&Delete"
        h = MsgBox("Are you sure to delete " & Text1.Text, vbQuestion + vbYesNo, "Confirmation")
        If h = vbYes Then DB.Execute "DELETE FROM TBL_CODE WHERE ID_CODE=" & Pesan
    End Select
    Hancurkan = True
    Unload Me
End Sub

Private Sub Command4_Click()
    Unload Me
    frmType.Show vbModal, frmMain
End Sub

Private Sub Command6_Click()
    Unload Me
    frmCategories.Show vbModal, frmMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim h As Long
    If Command2.Caption = "&Save" Then
      If Not Hancurkan Then
        h = MsgBox("Do you want to save the changes you made to Database", vbQuestion + vbYesNoCancel, "Programmer library")
        If h = vbYes Then
            Command2_Click
            Cancel = 7
            Exit Sub
        ElseIf h = vbNo Then
            Command2.Caption = "&Edit"
            Unload Me
        Else
            Cancel = 7
        End If
      End If
    End If
    h = 0
    PesanCat = ""
    PesanTip = ""
    Pekerjaan = ""
    Nilai = ""
    NilaiCat = ""
    NilaiTip = ""
    Pesan = 0
End Sub

Private Sub rtfText_SelChange()
    fMainForm.tbToolBar.Buttons("Bold").Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Italic").Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Underline").Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Align Left").Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Center").Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    fMainForm.tbToolBar.Buttons("Align Right").Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
    

End Sub

Private Sub Form_Load()
    Combo1.Enabled = False
    Combo2.Enabled = False
    Interface
    Form_Resize
    Me.Width = 6100
    Me.Height = 5000
    rtfText.SelIndent = 50
End Sub

Sub Interface()
On Error Resume Next

    Nilai = Pekerjaan
    Select Case Nilai
    Case "open"
        Pesan = ID_KODE
        
        tbS.Tabs(2).Selected = True
        
        Command2.Caption = "&Edit"
        Command3.Caption = "&Delete"
        
        If RS.State = 1 Then RS.Close
        
        NilaiCat = PesanCat
                       
        Combo1.AddItem NilaiCat
        Combo1.ListIndex = 0
        
        If RS.State = 1 Then RS.Close
        
        RS.Open "SELECT * FROM TBL_CODE WHERE ID_CODE=" & Pesan, DB, adOpenKeyset, adLockOptimistic
        NilaiTip = PesanTip
        
        Combo2.AddItem NilaiTip
        Combo2.ListIndex = 0
        
        Text1.Text = RS!JUDUL
        Text2.Text = RS!AUTHOR
        Text3.Text = RS!EMAIL
        Text4.Text = Format(RS!TANGGAL, "dd-MM-yyyy hh:mm:ss")
        rtfText.TextRTF = RS!CODE
        
        KunciKontrol True
        
        If RS.State = 1 Then RS.Close
    
    Case "new"
        TampilCategories
        TampilTipe
        KunciKontrol False
        Command2.Caption = "&Save"
        Command3.Caption = "C&ancel"
        Text4.Text = Format(Date, "dd-MM-yyyy") & " " & Format(Time, "hh:mm:ss")
    End Select
End Sub

Sub TampilCategories()
    If RS.State = 1 Then RS.Close
    
    RS.Open "SELECT * FROM TBL_CATEGORIES", DB, adOpenKeyset, adLockOptimistic
    
    With Combo1
        .Clear
        
        While Not RS.EOF
            .AddItem RS!CATEGORIES
            RS.MoveNext
        Wend
        
        RS.Close
        
        If Combo1.ListCount > -1 Then Combo1.ListIndex = 0
    
    End With
End Sub

Sub TampilTipe()
    If RS.State = 1 Then RS.Close
    
    RS.Open "SELECT * FROM TBL_TIPE WHERE CATEGORIES='" & Combo1.List(Combo1.ListIndex) & "'", DB, adOpenKeyset, adLockOptimistic
    
    With Combo2
        .Clear
        Combo4.Clear
        
        While Not RS.EOF
            .AddItem RS!TIPE
            Combo4.AddItem RS!ID_TIPE
            
            
            RS.MoveNext
        Wend
        
        RS.Close
        
        If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
    
    
    End With
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If Me.Height < 5000 Then Me.Height = 5000
    If Me.Width < 6000 Then Me.Width = 6100
    tbS.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 900
    Frame1.Move 100, (tbS.Top + tbS.Height + 20)
    Command1.Move (tbS.Left + tbS.Width) - Command1.Width, (tbS.Top + tbS.Height + 150)
    fr2.Move (tbS.Left + 200), (tbS.Top + 400), (tbS.Width - 400), (tbS.Height - 600)
    bingKai.Move (fr2.Width / 2) - (bingKai.Width / 2), (fr2.Height / 2) - (bingKai.Height / 2 - 30)
    fr1.Move fr2.Left, fr2.Top, fr2.Width, fr2.Height
    rtfText.Move 100, 200, fr1.Width - 200, fr1.Height - 300
End Sub

Private Sub tbS_Click()
    If tbS.Tabs(1).Selected Then
        fr1.Visible = False
        fr2.Visible = True
    Else
        fr1.Visible = True
        fr2.Visible = False
    End If
End Sub

