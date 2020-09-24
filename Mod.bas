Attribute VB_Name = "Mod"
Public fMainForm As frmMain
Global DB As New ADODB.Connection
Global RS As New ADODB.Recordset
Global ID_KODE As Long
Global Pekerjaan As String
Global PesanCat As String
Global PesanTip As String

Const PasswordDb = "serventlord"

Sub main()
    Set DB = New Connection
    Set RS = New Recordset
    Set fMainForm = frmMain
    
    DB.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & _
        App.Path & "\DataCode.MDB;PERSIST SECURITY INFO=FALSE;" & _
        "JET OLEDB:DATABASE PASSWORD=" & PasswordDb & ";"
    RS.CursorLocation = adUseClient
    frmMain.Show
    
End Sub

Public Function Encrypt(ByRef Tex As String) As String
    Dim i As Long
    Dim Tmp As String
    
    For i = 1 To Len(Tex)
        Tmp = Tmp & Asc(Mid(Tex, i, 1)) & ":"
    Next i
    Encrypt = Tmp
    Tmp = ""
End Function

Public Function Decrypt(ByRef Tex As String) As String
    On Error GoTo er:
    Dim i As Long
    Dim Tmp As String
    Dim Dima() As String
    
    Dima = Split(Tex, ":")
    For i = 0 To Len(Tex)
        Tmp = Tmp & Chr(Dima(i))
    Next i
  
    Exit Function

er:
        Decrypt = Tmp
        Tmp = ""
End Function


