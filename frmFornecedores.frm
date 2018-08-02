VERSION 5.00
Begin VB.Form frmFornecedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fornecedores"
   ClientHeight    =   3075
   ClientLeft      =   5175
   ClientTop       =   3975
   ClientWidth     =   6585
   Icon            =   "frmFornecedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6585
   Begin VB.Data dtcFornecedor 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4410
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Tabela_Fornecedores"
      Top             =   2730
      Width           =   1230
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   510
      Left            =   4380
      TabIndex        =   10
      Top             =   2175
      Width           =   1260
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "&Alterar"
      Height          =   510
      Left            =   3180
      TabIndex        =   9
      Top             =   2175
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   510
      Left            =   1950
      TabIndex        =   8
      Top             =   2175
      Width           =   1245
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "&Cadastrar"
      Height          =   510
      Left            =   3180
      TabIndex        =   7
      Top             =   2175
      Width           =   1200
   End
   Begin VB.Frame fraFornecedor 
      Caption         =   " Fornecedor "
      ForeColor       =   &H00FF0000&
      Height          =   1650
      Left            =   105
      TabIndex        =   0
      Top             =   300
      Width           =   6375
      Begin VB.TextBox txtEndereco 
         DataField       =   "Endereco"
         DataSource      =   "dtcFornecedor"
         Height          =   285
         Left            =   960
         MaxLength       =   50
         TabIndex        =   4
         Top             =   810
         Width           =   5130
      End
      Begin VB.ComboBox cboRazaoSocial 
         DataField       =   "RazaoSocial"
         DataSource      =   "dtcFornecedor"
         Height          =   315
         Left            =   1545
         TabIndex        =   2
         Top             =   270
         Width           =   4560
      End
      Begin VB.ComboBox cboFornecedores 
         DataField       =   "Fornecedor"
         DataSource      =   "dtcFornecedor"
         Height          =   315
         Left            =   165
         TabIndex        =   1
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label lblVRSaldo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Saldo"
         DataSource      =   "dtcFornecedor"
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
         Height          =   315
         Left            =   4725
         TabIndex        =   6
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblSaldo 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Atual :"
         Height          =   195
         Left            =   3630
         TabIndex        =   5
         Top             =   1260
         Width           =   900
      End
      Begin VB.Label lblEndereco 
         AutoSize        =   -1  'True
         Caption         =   "Endereço : "
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   810
         Width           =   825
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
      Left            =   3780
      TabIndex        =   11
      Top             =   2805
      Width           =   540
   End
End
Attribute VB_Name = "frmFornecedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset

Private Sub cboFornecedores_DropDown()
   
   cboFornecedores.Clear
   adoResultado.Open "SELECT * FROM Tabela_Fornecedores", cnaConexao, adOpenForwardOnly, adLockReadOnly
   
   If Not adoResultado.EOF Then
      While Not adoResultado.EOF
         cboFornecedores.AddItem adoResultado!Fornecedor
         adoResultado.MoveNext
      Wend
   End If
   adoResultado.Close
End Sub

Private Sub cboFornecedores_GotFocus()
   DadosLimpar
   ProximoFornecedor
End Sub

Private Sub cboFornecedores_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 Then
      If KeyAscii < 48 Or KeyAscii > 57 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub cboFornecedores_LostFocus()
   BuscaFornecedor (cboFornecedores.Text)
End Sub

Private Sub cboRazaoSocial_DropDown()
   
   cboRazaoSocial.Clear
   adoResultado.Open "SELECT * FROM Tabela_Fornecedores", cnaConexao, adOpenForwardOnly, adLockReadOnly
   
   If Not adoResultado.EOF Then
      While Not adoResultado.EOF
         cboRazaoSocial.AddItem adoResultado!RazaoSocial
         adoResultado.MoveNext
      Wend
   End If
   adoResultado.Close

End Sub

Private Sub cboRazaoSocial_LostFocus()
 BuscaRazaoSocial (Replace(cboRazaoSocial.Text, "'", "#"))
End Sub

Private Sub cmdAlterar_Click()
   If DadosValidar Then
      cboRazaoSocial.Text = Replace(cboRazaoSocial.Text, "'", "#")
      txtEndereco = Replace(txtEndereco.Text, "'", "#")
      strSQL = "UPDATE Tabela_Fornecedores SET RazaoSocial = '" & cboRazaoSocial.Text & "', " & _
               "Endereco = '" & txtEndereco.Text & "' WHERE Fornecedor = " & cboFornecedores.Text
      
      cnaConexao.Execute strSQL
      DoEvents
      MsgBox "Alteração feita com sucesso.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      AtualizaDados
      DadosLimpar
      cboFornecedores.SetFocus
      
      
   End If
End Sub

Private Sub cmdCadastrar_Click()
   If DadosValidar Then
      cboRazaoSocial.Text = Replace(cboRazaoSocial.Text, "'", "#")
      txtEndereco = Replace(txtEndereco.Text, "'", "#")
      strSQL = "INSERT INTO Tabela_Fornecedores (Fornecedor, RazaoSocial, Endereco, Saldo ) " & _
               "VALUES ( " & Trim(cboFornecedores.Text) & ",'" & cboRazaoSocial.Text & "'," & _
               " '" & txtEndereco & "',0)"
     
     cnaConexao.Execute strSQL
     DoEvents
     MsgBox "Registro cadastrado com sucesso.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
     AtualizaDados
     cboFornecedores.SetFocus
     DadosLimpar
     
   End If
End Sub

Private Sub cmdExcluir_Click()
  
   If MsgBox("Deseja excluir este fornecedor?", vbQuestion + vbYesNo, "Mensagem ao Usuário") = vbYes Then
      If Not BuscaMovimentos(cboFornecedores.Text) Then
         strSQL = "DELETE FROM Tabela_Fornecedores WHERE Fornecedor = " & cboFornecedores.Text
         
         cnaConexao.Execute strSQL
         DoEvents
         MsgBox "Registro excluído com sucesso.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
         AtualizaDados
         cboFornecedores.SetFocus
         
      End If
   End If

End Sub

Private Sub cmdFechar_Click()
Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Centraliza Me
If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Centraliza Me
dtcFornecedor.DatabaseName = Caminho & nmBanco
End Sub
Private Function DadosValidar()
DadosValidar = True

   If Len(Trim(cboFornecedores.Text)) = 0 Then
      Beep
      MsgBox "Favor informar o número do fornecedor.", vbOKOnly + vbInformation, "Mensagem ao Usuário"
      cboFornecedores.SetFocus
      DadosValidar = False
      Exit Function
   End If
   
   If Len(Trim(cboRazaoSocial.Text)) = 0 Then
      Beep
      MsgBox "É necessário informar a razão social deste fornecedor.", vbOKOnly + vbInformation, "Mensagem ao Usuário"
      cboRazaoSocial.SetFocus
      DadosValidar = False
      Exit Function
   End If
   
   If Len(cboRazaoSocial.Text) > 50 Then
      Beep
      MsgBox "A razão social deste fornecedor deve conter somente até 50 caracteres.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      cboRazaoSocial.SetFocus
      DadosValidar = False
      Exit Function
   End If
End Function
Private Function ProximoFornecedor()

   adoResultado.Open "SELECT IIF(ISNULL(MAX(Fornecedor)),0,MAX(Fornecedor)) + 1 AS Fornecedor FROM Tabela_Fornecedores", cnaConexao, adOpenForwardOnly, adLockReadOnly
   cboFornecedores.Text = adoResultado!Fornecedor
   adoResultado.Close
   cboFornecedores.SelStart = 0
   cboFornecedores.SelLength = Len(cboFornecedores.Text)
End Function
Private Function DadosLimpar()
  
  cboFornecedores.Text = Space$(0)
  cboRazaoSocial.Text = Space$(0)
  txtEndereco.Text = Space$(0)
  lblVRSaldo.Caption = Space$(0)
  cmdCadastrar.Visible = True
  cmdAlterar.Visible = False
  cmdExcluir.Visible = False

End Function
Private Function BuscaFornecedor(strFornecedor As String)

   If Len(Trim(strFornecedor)) > 0 Then
      adoResultado.Open "SELECT * FROM Tabela_Fornecedores WHERE Fornecedor = " & strFornecedor, cnaConexao, adOpenForwardOnly, adLockReadOnly
         If Not adoResultado.EOF Then
            cboRazaoSocial.Text = Replace(adoResultado!RazaoSocial, "#", "'")
            txtEndereco = Replace(adoResultado!Endereco, "#", "'")
            lblVRSaldo.Caption = adoResultado!Saldo
            cmdCadastrar.Visible = False
            cmdAlterar.Visible = True
            cmdExcluir.Visible = True
         End If
      adoResultado.Close
   End If
End Function
Private Function BuscaMovimentos(strFornecedor As String) As Boolean
   If Len(Trim(strFornecedor)) > 0 Then
      adoResultado.Open "SELECT TOP 1 * FROM Tabela_Pedidos WHERE Fornecedor = " & strFornecedor, cnaConexao, adOpenForwardOnly, adLockReadOnly
      If Not adoResultado.EOF Then
         Beep
         MsgBox "Este fornecedor possui pedidos e não poderá ser excluído.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
         BuscaMovimentos = True
         adoResultado.Close
         Exit Function
      End If
      adoResultado.Close
      
      adoResultado.Open "SELECT TOP 1 * FROM Tabela_Notas WHERE Fornecedor = " & strFornecedor, cnaConexao, adOpenForwardOnly, adLockReadOnly
      If Not adoResultado.EOF Then
         Beep
         MsgBox "Este fornecedor possui notas e não poderá ser excluído.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
         BuscaMovimentos = True
         adoResultado.Close
         Exit Function
      End If
      adoResultado.Close
   End If
End Function
Private Function BuscaRazaoSocial(strRazaoSocial As String)
   If Len(Trim(strRazaoSocial)) > 0 Then
      adoResultado.Open "SELECT * FROM Tabela_Fornecedores WHERE RazaoSocial = '" & strRazaoSocial & "'", cnaConexao, adOpenForwardOnly, adLockReadOnly
         If Not adoResultado.EOF Then
            cboFornecedores.Text = adoResultado!Fornecedor
            txtEndereco = Replace(adoResultado!Endereco, "#", "'")
            lblVRSaldo.Caption = adoResultado!Saldo
            cmdCadastrar.Visible = False
            cmdAlterar.Visible = True
            cmdExcluir.Visible = True
         End If
      adoResultado.Close
   End If
End Function
Private Function AtualizaDados()
   Dim intContador As Integer
   
   For intContador = 1 To 250
      dtcFornecedor.Refresh
   Next
End Function
