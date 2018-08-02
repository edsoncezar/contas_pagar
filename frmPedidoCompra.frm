VERSION 5.00
Begin VB.Form frmPedidoCompra 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido de Compra"
   ClientHeight    =   3810
   ClientLeft      =   2040
   ClientTop       =   2250
   ClientWidth     =   6165
   Icon            =   "frmPedidoCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.Data dtcPedido 
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4860
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Consulta_Pedidos"
      Top             =   3420
      Width           =   1245
   End
   Begin VB.ComboBox cboFornecedor 
      DataField       =   "Fornecedor"
      DataSource      =   "dtcPedido"
      Height          =   315
      Left            =   1350
      TabIndex        =   1
      Top             =   660
      Width           =   945
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   525
      Left            =   4800
      TabIndex        =   8
      Top             =   2730
      Width           =   1305
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   525
      Left            =   2190
      TabIndex        =   6
      Top             =   2730
      Width           =   1305
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   525
      Left            =   3510
      TabIndex        =   7
      Top             =   2730
      Width           =   1305
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "&Cadastrar"
      Height          =   525
      Left            =   3510
      TabIndex        =   5
      Top             =   2730
      Width           =   1305
   End
   Begin VB.TextBox txtPrecoUnitario 
      DataField       =   "PUnitario"
      DataSource      =   "dtcPedido"
      Height          =   285
      Left            =   1350
      TabIndex        =   4
      Top             =   2190
      Width           =   1695
   End
   Begin VB.TextBox txtDescricao 
      DataField       =   "Descricao"
      DataSource      =   "dtcPedido"
      Height          =   285
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1590
      Width           =   4755
   End
   Begin VB.TextBox txtQuantidade 
      DataField       =   "Quantidade"
      DataSource      =   "dtcPedido"
      Height          =   285
      Left            =   1350
      TabIndex        =   2
      Top             =   1110
      Width           =   1695
   End
   Begin VB.TextBox txtPedido 
      DataField       =   "Pedido"
      DataSource      =   "dtcPedido"
      Height          =   285
      Left            =   1350
      TabIndex        =   0
      Top             =   210
      Width           =   1695
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
      Left            =   4230
      TabIndex        =   15
      Top             =   3510
      Width           =   540
   End
   Begin VB.Label lblRazaoSocial 
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "RazaoSocial"
      DataSource      =   "dtcPedido"
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
      Left            =   2340
      TabIndex        =   14
      Top             =   660
      Width           =   3705
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Preço Unitário :"
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   13
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Descrição :"
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   12
      Top             =   1620
      Width           =   810
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Pedido :"
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   11
      Top             =   240
      Width           =   585
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade :"
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   10
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor :"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   9
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frmPedidoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboFornecedor_Click()
   Dim adoResultado As New ADODB.Recordset
   adoResultado.Open "SELECT RazaoSocial FROM Tabela_Fornecedores WHERE Fornecedor = " & cboFornecedor.Text, cnaConexao, adOpenForwardOnly, adLockReadOnly
   If Not adoResultado.EOF Then
      lblRazaoSocial = Replace(adoResultado!RazaoSocial, "#", "'")
   End If
   adoResultado.Close
End Sub

Private Sub cboFornecedor_DropDown()
   Dim adoResultado As New ADODB.Recordset
   cboFornecedor.Clear
   adoResultado.Open "SELECT Fornecedor FROM Tabela_Fornecedores", cnaConexao, adOpenForwardOnly, adLockReadOnly
   If Not adoResultado.EOF Then
      While Not adoResultado.EOF
      cboFornecedor.AddItem adoResultado!Fornecedor
      adoResultado.MoveNext
      Wend
   End If
   adoResultado.Close
End Sub
Private Sub cboFornecedor_LostFocus()
Dim adoResultado As New ADODB.Recordset
   If Len(Trim(cboFornecedor.Text)) > 0 Then
     adoResultado.Open "SELECT RazaoSocial FROM Tabela_Fornecedores WHERE Fornecedor = " & cboFornecedor.Text, cnaConexao, adOpenForwardOnly, adLockReadOnly
     If Not adoResultado.EOF Then
        lblRazaoSocial = Replace(adoResultado!RazaoSocial, "#", "'")
     Else
        lblRazaoSocial = Space$(0)
        BuscaFornecedor (Trim(cboFornecedor.Text))
     End If
     adoResultado.Close
  End If
End Sub
Private Sub cmdCadastrar_Click()
Dim dblValor  As Double
strSQL = Space$(0)
dblValor = Val(txtQuantidade) * Val(txtPrecoUnitario)
txtDescricao = Replace(txtDescricao, "'", "#")
txtQuantidade = Replace(txtQuantidade, ",", ".")
txtPrecoUnitario = Replace(txtPrecoUnitario, ",", ".")
strSQL = "INSERT INTO Tabela_Pedidos (Pedido,Fornecedor,Quantidade,Descricao,PUnitario,Saldo)" & _
         " VALUES (" & Trim(txtPedido) & "," & Trim(cboFornecedor.Text) & "," & _
                       Trim(txtQuantidade) & ",'" & txtDescricao & "'," & Trim(txtPrecoUnitario) & "," & dblValor & ")"
If DadosValidar Then
   cnaConexao.Execute strSQL
   DoEvents
   MsgBox "Registro cadastrado com sucesso.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
   AtualizaDados
   DadosLimpar
End If

End Sub
Private Sub cmdExcluir_Click()
strSQL = Space$(0)
Dim adoResultado As New ADODB.Recordset
If MsgBox("Deseja excluir este pedido?", vbQuestion + vbYesNo, "Mensagem ao Usuário") = vbYes Then
   adoResultado.Open "SELECT TOP 1 * FROM Tabela_Notas WHERE Pedido = " & Trim(txtPedido), cnaConexao, adOpenForwardOnly, adLockReadOnly
   If Not adoResultado.EOF Then
      Beep
      MsgBox "Este pedido possui nota(s) fiscal(is).Registro não pode ser excluído.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      adoResultado.Close
      Exit Sub
   End If
   adoResultado.Close
   strSQL = "DELETE FROM Tabela_Pedidos WHERE Pedido = " & Trim(txtPedido)
   cnaConexao.Execute strSQL
   DoEvents
   MsgBox "Registro excluído com sucesso.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
   AtualizaDados
   txtPedido.SetFocus
End If
End Sub
Private Sub cmdModificar_Click()
Dim dblValor As Double
strSQL = Space$(0)
If DadosValidar Then
   txtDescricao = Replace(txtDescricao, "'", "#")
   txtQuantidade = Replace(txtQuantidade, ",", ".")
   txtPrecoUnitario = Replace(txtPrecoUnitario, ",", ".")
   dblValor = Replace(Val(txtPrecoUnitario) * Val(txtQuantidade), ",", ".")
   strSQL = "UPDATE Tabela_Pedidos SET Quantidade = " & Trim(txtQuantidade) & _
            ",Descricao = '" & txtDescricao & "',PUnitario = " & Trim(txtPrecoUnitario) & _
            ", Saldo = " & dblValor & " WHERE Pedido = " & txtPedido
   cnaConexao.Execute strSQL
   txtDescricao = Replace(txtDescricao, "#", "'")
   txtQuantidade = Replace(txtQuantidade, ".", ",")
   txtPrecoUnitario = Replace(txtPrecoUnitario, ".", ",")
   DoEvents
   MsgBox "Dados modificados com sucesso. ", vbInformation + vbOKOnly, "Mensagem ao Usuário"
   AtualizaDados
   DadosLimpar
End If
End Sub
Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Centraliza Me
dtcPedido.DatabaseName = Caminho & nmBanco
End Sub

Private Sub txtPedido_GotFocus()
Dim adoResultado As New ADODB.Recordset
adoResultado.Open "SELECT IIF(MAX(Pedido) IS NULL, 0, MAX(Pedido)) + 1 as Pedido FROM Tabela_Pedidos", cnaConexao, adOpenForwardOnly, adLockReadOnly
If Not adoResultado.EOF Then
   txtPedido = adoResultado!Pedido
   txtPedido.SelStart = 0
   txtPedido.SelLength = Len(Trim(txtPedido))
End If
adoResultado.Close
DadosLimpar
txtQuantidade.Enabled = True
txtPrecoUnitario.Enabled = True
cboFornecedor.Enabled = True
cmdCadastrar.Visible = True
cmdExcluir.Visible = False
cmdModificar.Visible = False
End Sub
Private Function DadosValidar() As Boolean
DadosValidar = True
If Len(Trim(cboFornecedor.Text)) = 0 Then
   Beep
   MsgBox "Informe o número do fornecedor. ", vbInformation + vbOKOnly, "Mensagem ao Usuário"
   cboFornecedor.SetFocus
   DadosValidar = False
   Exit Function
Else
   If Not BuscaFornecedor(Trim(cboFornecedor.Text)) Then
      DadosValidar = False
      Exit Function
   End If
End If
If Len(Trim(txtPedido)) = 0 Then
   Beep
   MsgBox "Informe o número do pedido.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
   txtPedido.SetFocus
   DadosValidar = False
   Exit Function
End If
If Len(Trim(txtQuantidade)) = 0 Or Val(txtQuantidade) = 0 Then
   Beep
   MsgBox "Informe a quantidade corretamente.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
   txtQuantidade.SetFocus
   DadosValidar = False
   Exit Function
End If
If Len(Trim(txtPrecoUnitario)) = 0 Or Val(Replace(txtPrecoUnitario, ",", ".")) = 0 Then
   Beep
   MsgBox "Informe o preço unitário corretamente.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
   txtPrecoUnitario.SetFocus
   DadosValidar = False
   Exit Function
End If
End Function
Private Function BuscaFornecedor(Fornecedor As String) As Boolean
Dim adoResultado As New ADODB.Recordset
BuscaFornecedor = True
If Len(Trim(cboFornecedor.Text)) > 0 Then
  adoResultado.Open "SELECT RazaoSocial FROM Tabela_Fornecedores WHERE Fornecedor = " & Fornecedor, cnaConexao, adOpenForwardOnly, adLockReadOnly
  If adoResultado.EOF Then
     Beep
     MsgBox "Fornecedor não cadastrado.", vbInformation + vbOKOnly
     cboFornecedor.Text = ""
     BuscaFornecedor = False
  End If
  adoResultado.Close
Else
   BuscaFornecedor = False
End If
End Function
Private Sub txtPedido_LostFocus()
If Len(Trim(txtPedido)) > 0 Then
   DadosPreencher
End If
End Sub
Private Function DadosPreencher()
Dim adoResultado As New ADODB.Recordset
adoResultado.Open "SELECT Pedido, Fornecedor, Quantidade, Descricao, PUnitario FROM Tabela_Pedidos WHERE Pedido = " & txtPedido, cnaConexao, adOpenForwardOnly, adLockReadOnly
If Not adoResultado.EOF Then
   cboFornecedor.Text = adoResultado!Fornecedor
   cboFornecedor_LostFocus
   cboFornecedor.Enabled = False
   cmdCadastrar.Visible = False
   cmdExcluir.Visible = True
   cmdModificar.Visible = True
   txtQuantidade = adoResultado!Quantidade
   txtDescricao = adoResultado!Descricao
   txtDescricao = Replace(txtDescricao, "#", "'")
   txtQuantidade = Replace(txtQuantidade, ".", ",")
   txtPrecoUnitario = adoResultado!PUnitario
   txtPrecoUnitario = Replace(txtPrecoUnitario, ".", ",")
   txtQuantidade.Enabled = True
   txtPrecoUnitario.Enabled = True
End If
adoResultado.Close
End Function
Private Function DadosLimpar()
cboFornecedor.Clear
lblRazaoSocial.Caption = Space$(0)
txtQuantidade = Space$(0)
txtDescricao = Space$(0)
txtPrecoUnitario = Space$(0)
txtPedido.SetFocus
End Function
Private Function VerificaNotaFiscal() As Boolean
Dim adoResultado As New ADODB.Recordset
adoResultado.Open "SELECT TOP 1 * FROM Tabela_Notas WHERE Pedido = " & Trim(txtPedido), cnaConexao, adOpenForwardOnly, adLockReadOnly
If Not adoResultado.EOF Then
   VerificaNotaFiscal = True
End If
adoResultado.Close
End Function

Private Sub txtPrecoUnitario_KeyPress(KeyAscii As Integer)
   
   If KeyAscii <> 8 And KeyAscii <> 44 Then
      If KeyAscii < 48 Or KeyAscii > 57 Then
         KeyAscii = 0
      End If
   End If

End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
      If KeyAscii < 48 Or KeyAscii > 57 Then
         KeyAscii = 0
      End If
   End If
End Sub
Private Function AtualizaDados()
   Dim intContador As Integer
   
   For intContador = 1 To 250
      dtcPedido.Refresh
   Next
End Function
