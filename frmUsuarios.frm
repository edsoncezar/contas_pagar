VERSION 5.00
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   3090
   ClientLeft      =   4200
   ClientTop       =   3495
   ClientWidth     =   7665
   Icon            =   "frmUsuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7665
   Begin VB.Data dtcUsuario 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5250
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   "Tabela_Usuarios"
      Top             =   2730
      Width           =   1230
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   510
      Left            =   5220
      TabIndex        =   14
      Top             =   2175
      Width           =   1260
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   510
      Left            =   4020
      TabIndex        =   13
      Top             =   2175
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   510
      Left            =   2790
      TabIndex        =   12
      Top             =   2175
      Width           =   1245
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "&Cadastrar"
      Height          =   510
      Left            =   4020
      TabIndex        =   11
      Top             =   2175
      Width           =   1200
   End
   Begin VB.Frame fraUsuario 
      Caption         =   " Usuário"
      ForeColor       =   &H00FF0000&
      Height          =   1650
      Left            =   75
      TabIndex        =   0
      Top             =   285
      Width           =   7560
      Begin VB.TextBox txtConfirmaSenha 
         Height          =   285
         Left            =   2700
         TabIndex        =   6
         Top             =   1245
         Width           =   1005
      End
      Begin VB.ComboBox cboPerfil 
         Height          =   315
         ItemData        =   "frmUsuarios.frx":0442
         Left            =   4260
         List            =   "frmUsuarios.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1245
         Width           =   1350
      End
      Begin VB.TextBox txtSenha 
         DataField       =   "Senha"
         DataSource      =   "dtcUsuario"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1245
         Width           =   1005
      End
      Begin VB.ComboBox cboNome 
         DataField       =   "Nome"
         DataSource      =   "dtcUsuario"
         Height          =   315
         Left            =   1545
         TabIndex        =   2
         Top             =   270
         Width           =   5955
      End
      Begin VB.ComboBox cboUsuarios 
         DataField       =   "Usuario"
         DataSource      =   "dtcUsuario"
         Height          =   315
         Left            =   165
         TabIndex        =   1
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label lblConfirma 
         AutoSize        =   -1  'True
         Caption         =   "Confirmar :"
         Height          =   195
         Left            =   1935
         TabIndex        =   5
         Top             =   1290
         Width           =   750
      End
      Begin VB.Label lblBloqueio 
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   6495
         TabIndex        =   10
         Top             =   1245
         Width           =   1035
      End
      Begin VB.Label lblBloq 
         AutoSize        =   -1  'True
         Caption         =   "Bloqueado: "
         Height          =   195
         Left            =   5640
         TabIndex        =   9
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label lblPerfil 
         AutoSize        =   -1  'True
         Caption         =   "Perfil :"
         Height          =   195
         Left            =   3795
         TabIndex        =   7
         Top             =   1275
         Width           =   435
      End
      Begin VB.Label lblSenha 
         AutoSize        =   -1  'True
         Caption         =   "Senha : "
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   1260
         Width           =   600
      End
   End
   Begin VB.Label lblBusca 
      AutoSize        =   -1  'True
      Caption         =   "Busca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   4620
      TabIndex        =   15
      Top             =   2805
      Width           =   540
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset
Function usuarioBloqueado(bloq As String) As String
    If bloq = "1" Then
        usuarioBloqueado = "Sim"
    Else
        usuarioBloqueado = "Não"
    End If
End Function
Private Sub cboUsuarios_DropDown()
   cboUsuarios.Clear
   adoResultado.Open "SELECT * FROM Tabela_Usuarios", cnaConexao, adOpenForwardOnly, adLockReadOnly
   If Not adoResultado.EOF Then
      While Not adoResultado.EOF
         cboUsuarios.AddItem adoResultado!Usuario
         adoResultado.MoveNext
      Wend
   End If
   adoResultado.Close
End Sub

Private Sub cboUsuarios_GotFocus()
   DadosLimpar
   ProximoUsuario
End Sub

Private Sub cboUsuarios_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 Then
      If KeyAscii < 48 Or KeyAscii > 57 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub cboUsuarios_LostFocus()
   BuscaUsuario (cboUsuarios.Text)
End Sub

Private Sub cboNome_DropDown()
   cboNome.Clear
   adoResultado.Open "SELECT Nome FROM Tabela_Usuarios", cnaConexao, adOpenForwardOnly, adLockReadOnly
   If Not adoResultado.EOF Then
      While Not adoResultado.EOF
         cboNome.AddItem adoResultado!Nome
         adoResultado.MoveNext
      Wend
   End If
   adoResultado.Close
End Sub
Private Sub cboNome_LostFocus()
 BuscaUsuario (cboUsuarios.Text)
End Sub
Private Sub cmdAlterar_Click()
   If DadosValidar Then
      cboNome.Text = Replace(cboNome.Text, "'", "#")
      txtSenha = Replace(txtSenha.Text, "'", "#")
      strSQL = "UPDATE Tabela_Usuarios SET Nome = '" & cboNome.Text & "', " & _
               "Senha = '" & txtSenha.Text & "',Bloqueio = '0' WHERE Usuario = " & cboUsuarios.Text
      cnaConexao.Execute strSQL
      Beep
      MsgBox "Alteração feita com sucesso.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      dtcUsuario.Refresh
      cboUsuarios.SetFocus
   End If
End Sub
Private Sub cmdCadastrar_Click()
   If DadosValidar Then
      cboNome.Text = Replace(cboNome.Text, "'", "#")
      txtSenha = Replace(txtSenha.Text, "'", "#")
      strSQL = "INSERT INTO Tabela_Usuarios (Usuario, Nome, Senha, Perfil, Bloqueio ) " & _
               "VALUES ( " & Trim(cboUsuarios.Text) & ",'" & cboNome.Text & "'," & _
               " '" & txtSenha & "','" & cboPerfil.ListIndex & "','0')"
     cnaConexao.Execute strSQL
     dtcUsuario.Refresh
     cboUsuarios.SetFocus
     DadosLimpar
   End If
End Sub
Private Sub cmdExcluir_Click()
   If idUsuario = Val(cboUsuarios.Text) Then
       MsgBox "Este é o usuário logado e não poderá ser excluido.", vbOKOnly + vbInformation, "Mensagem ao Usuário"
       Exit Sub
   End If
   If MsgBox("Deseja excluir este usuário?", vbQuestion + vbYesNo, "Mensagem ao Usuário") = vbYes Then
      strSQL = "DELETE FROM Tabela_Usuarios WHERE Usuario = " & cboUsuarios.Text
      cnaConexao.Execute strSQL
      dtcUsuario.Refresh
      cboUsuarios.SetFocus
   End If
End Sub
Private Sub cmdFechar_Click()
Unload Me
End Sub

Private Sub dtcUsuario_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cboUsuarios_LostFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Centraliza Me
If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Centraliza Me
dtcUsuario.DatabaseName = Caminho & nmBanco
End Sub
Private Function DadosValidar()
DadosValidar = True
   If Len(Trim(cboUsuarios.Text)) = 0 Then
      Beep
      MsgBox "Favor informar o número do usuário.", vbOKOnly + vbInformation, "Mensagem ao Usuário"
      cboUsuarios.SetFocus
      DadosValidar = False
      Exit Function
   End If
   If Len(Trim(cboNome.Text)) = 0 Then
      Beep
      MsgBox "É necessário informar o nome do usuário.", vbOKOnly + vbInformation, "Mensagem ao Usuário"
      cboNome.SetFocus
      DadosValidar = False
      Exit Function
   End If
   If Len(cboNome.Text) > 50 Then
      Beep
      MsgBox "O nome do usuário deve conter somente até 50 caracteres.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      cboNome.SetFocus
      DadosValidar = False
      Exit Function
   End If
   If txtSenha <> txtConfirmaSenha Then
      Beep
      MsgBox "A senha não confirma.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      txtConfirmaSenha.SetFocus
      DadosValidar = False
      Exit Function
   End If
End Function
Private Function ProximoUsuario()
   adoResultado.Open "SELECT IIF(ISNULL(MAX(Usuario)),0,MAX(Usuario)) + 1 AS Usuario FROM Tabela_Usuarios", cnaConexao, adOpenForwardOnly, adLockReadOnly
   cboUsuarios.Text = adoResultado!Usuario
   adoResultado.Close
   cboUsuarios.SelStart = 0
   cboUsuarios.SelLength = Len(cboUsuarios.Text)
End Function
Private Function DadosLimpar()
  cboUsuarios.Text = Space$(0)
  cboNome.Text = Space$(0)
  txtSenha = Space$(0)
  txtSenha.PasswordChar = "*"
  txtConfirmaSenha.Text = Space$(0)
  txtConfirmaSenha.PasswordChar = "*"
  lblBloqueio.Caption = Space$(0)
  cmdCadastrar.Visible = True
  cmdAlterar.Visible = False
  cmdExcluir.Visible = False
End Function
Private Function BuscaUsuario(strUsuario As String)
   If Len(Trim(strUsuario)) > 0 Then
      adoResultado.Open "SELECT * FROM Tabela_Usuarios WHERE Usuario = " & strUsuario, cnaConexao, adOpenForwardOnly, adLockReadOnly
         If Not adoResultado.EOF Then
            cboNome.Text = Replace(adoResultado!Nome, "#", "'")
            txtSenha = Replace(adoResultado!Senha, "#", "'")
            cboPerfil.Text = cboPerfil.List(Val(adoResultado!Perfil))
            lblBloqueio = usuarioBloqueado(adoResultado!bloqueio)
            cmdCadastrar.Visible = False
            cmdAlterar.Visible = True
            cmdExcluir.Visible = True
         End If
      adoResultado.Close
   End If
End Function

