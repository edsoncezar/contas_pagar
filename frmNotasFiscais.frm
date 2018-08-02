VERSION 5.00
Begin VB.Form frmNotasFiscais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Fiscais"
   ClientHeight    =   3405
   ClientLeft      =   4110
   ClientTop       =   3960
   ClientWidth     =   6585
   Icon            =   "frmNotasFiscais.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6585
   Begin VB.Data dtcNotas 
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Consulta_Notas"
      Top             =   2970
      Width           =   1140
   End
   Begin VB.TextBox txtNF 
      DataField       =   "NotaFiscal"
      DataSource      =   "dtcNotas"
      Height          =   285
      Left            =   1650
      TabIndex        =   0
      Top             =   255
      Width           =   1005
   End
   Begin VB.ComboBox cboPedido 
      DataField       =   "Pedido"
      DataSource      =   "dtcNotas"
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtPrecoUnitario 
      DataField       =   "PUnitario"
      DataSource      =   "dtcNotas"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1665
      TabIndex        =   5
      Top             =   1860
      Width           =   1695
   End
   Begin VB.TextBox txtDescricao 
      DataField       =   "Descricao"
      DataSource      =   "dtcNotas"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1665
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1470
      Width           =   4755
   End
   Begin VB.TextBox txtQuantidade 
      DataField       =   "Quantidade"
      DataSource      =   "dtcNotas"
      Height          =   285
      Left            =   1665
      TabIndex        =   3
      Top             =   1065
      Width           =   1695
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Height          =   525
      Left            =   5130
      TabIndex        =   8
      Top             =   2280
      Width           =   1305
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   525
      Left            =   3840
      TabIndex        =   6
      Top             =   2280
      Width           =   1305
   End
   Begin VB.CommandButton cmdCadastrar 
      Caption         =   "&Cadastrar"
      Height          =   525
      Left            =   3840
      TabIndex        =   7
      Top             =   2280
      Width           =   1305
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
      Left            =   4710
      TabIndex        =   15
      Top             =   3030
      Width           =   540
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Nota Fiscal :"
      Height          =   195
      Index           =   5
      Left            =   720
      TabIndex        =   14
      Top             =   285
      Width           =   885
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedor :"
      Height          =   165
      Index           =   0
      Left            =   720
      TabIndex        =   13
      Top             =   765
      Width           =   855
   End
   Begin VB.Label lblRazaoSocial 
      BackColor       =   &H8000000C&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "RazaoSocial"
      DataSource      =   "dtcNotas"
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
      Left            =   1650
      TabIndex        =   2
      Top             =   690
      Width           =   3660
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Preço Unitário :"
      Height          =   195
      Index           =   4
      Left            =   540
      TabIndex        =   12
      Top             =   1890
      Width           =   1095
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Descrição :"
      Height          =   195
      Index           =   3
      Left            =   765
      TabIndex        =   11
      Top             =   1500
      Width           =   810
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Pedido :"
      Height          =   195
      Index           =   2
      Left            =   3045
      TabIndex        =   10
      Top             =   285
      Width           =   585
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade :"
      Height          =   195
      Index           =   1
      Left            =   705
      TabIndex        =   9
      Top             =   1095
      Width           =   915
   End
End
Attribute VB_Name = "frmNotasFiscais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub txtNota_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys "{TAB}"
   End If
End Sub

Private Sub txtNota_LostFocus()
   If Len(Trim(txtNota)) > 0 Then
      DadosPreencher
   End If
End Sub

Private Sub cboFornecedor_Click()
   Dim adoResultado As New ADODB.Recordset
   adoResultado.Open "SELECT RazaoSocial FROM Tabela_Fornecedores WHERE Fornecedor = " & cboFornecedor.Text, cnaConexao, adOpenForwardOnly, adLockReadOnly
   If Not adoResultado.EOF Then
      lblRazaoSocial = adoResultado!RazaoSocial
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

Private Sub cboPedido_Click()
   BuscaFornecedor
End Sub

Private Sub cboPedido_DropDown()
 Dim adoResultado As New ADODB.Recordset
   
   adoResultado.Open "SELECT Pedido FROM Tabela_Pedidos", cnaConexao, adOpenForwardOnly, adLockReadOnly
   
   While Not adoResultado.EOF
      cboPedido.AddItem adoResultado!Pedido
      adoResultado.MoveNext
   Wend
   
   adoResultado.Close

End Sub

Private Sub cboPedido_LostFocus()
   BuscaFornecedor
End Sub

Private Sub cmdCadastrar_Click()
   Dim intFornecedor As Integer
   If DadosValidar Then
      intFornecedor = InStr(1, lblRazaoSocial.Caption, " - ")
      
      strSQL = Space$(0)
      txtDescricao = Replace(txtDescricao, "'", "#")
      txtQuantidade = Replace(txtQuantidade, ",", ".")
      txtPrecoUnitario = Replace(txtPrecoUnitario, ",", ".")
      
      strSQL = "INSERT INTO Tabela_Notas (NotaFiscal,Fornecedor,Pedido,Quantidade,Descricao,PUnitario) " & _
               " VALUES (" & Trim(txtNF) & "," & Mid(lblRazaoSocial.Caption, 1, intFornecedor - 1) & "," & _
               cboPedido.Text & "," & txtQuantidade & ",'" & txtDescricao & "'," & txtPrecoUnitario & ")"
      
      If VerificaSaldo < (Val(txtQuantidade) * Val(txtPrecoUnitario)) Then
         Beep
         MsgBox "Saldo do pedido insulficiente para emissão desta nota.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
         Exit Sub
      Else
         cnaConexao.Execute strSQL
         strSQL = "UPDATE Tabela_Fornecedores SET Saldo = Saldo + " & Replace((Val(txtQuantidade) * Val(txtPrecoUnitario)), ",", ".") & _
                " WHERE Fornecedor = " & Mid(lblRazaoSocial.Caption, 1, intFornecedor - 1)
         cnaConexao.Execute strSQL
         strSQL = "UPDATE Tabela_Pedidos SET Saldo = Saldo -  " & Replace((Val(txtQuantidade) * Val(txtPrecoUnitario)), ",", ".") & _
                  " WHERE Pedido = " & cboPedido.Text
         cnaConexao.Execute strSQL
      End If
   
      txtDescricao = Replace(txtDescricao, "#", "'")
      txtQuantidade = Replace(txtQuantidade, ".", ",")
      txtPrecoUnitario = Replace(txtPrecoUnitario, ".", ",")
      DoEvents
      MsgBox "Registro cadastrado com sucesso.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      AtualizaDados
      DadosLimpar
   End If
End Sub

Private Sub cmdExcluir_Click()
   Dim adoResultado As New ADODB.Recordset
   Dim intFornecedor As Integer
   intFornecedor = InStr(1, lblRazaoSocial.Caption, " - ")
   
   strSQL = Space$(0)
   If MsgBox("Deseja excluir esta Nota Fiscal?", vbQuestion + vbYesNo, "Mensagem ao Usuário") = vbYes Then
      strSQL = "DELETE FROM Tabela_Notas WHERE NotaFiscal = " & txtNF
      cnaConexao.Execute strSQL
      strSQL = "UPDATE Tabela_Pedidos SET Saldo = Saldo + " & (Val(txtQuantidade) * Val(txtPrecoUnitario)) & _
               " WHERE Pedido = " & cboPedido.Text
      cnaConexao.Execute strSQL
      strSQL = "UPDATE Tabela_Fornecedores SET Saldo = Saldo - " & (Val(txtQuantidade) * Val(txtPrecoUnitario)) & _
               " WHERE Fornecedor = " & Mid(lblRazaoSocial.Caption, 1, intFornecedor - 1)
      cnaConexao.Execute strSQL
      DoEvents
      MsgBox "Registro excluído com sucesso.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      AtualizaDados
      txtNF.SetFocus
      
   End If
End Sub
Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub txtNota_GotFocus()
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
  cboFornecedor.Enabled = True
  cmdCadastrar.Visible = True
  cmdExcluir.Visible = False
  
End Sub
Private Function DadosValidar() As Boolean
   DadosValidar = True
   If Len(Trim(txtNF.Text)) = 0 Then
      Beep
      MsgBox "Informe o número da Nota Fiscal.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      cboPedido.SetFocus
      DadosValidar = False
      Exit Function
   End If
   If Len(Trim(cboPedido.Text)) = 0 Then
      Beep
      MsgBox "Informe o número do pedido.", vbInformation + vbOKOnly, "Mensagem ao Usuário"
      cboPedido.SetFocus
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
Private Function DadosPreencher()
   Dim adoResultado As New ADODB.Recordset
   adoResultado.Open "SELECT * FROM Tabela_Notas WHERE NotaFiscal = " & txtNF, cnaConexao, adOpenForwardOnly, adLockReadOnly
   If Not adoResultado.EOF Then
      cboPedido.Text = adoResultado!Pedido
      cboPedido_LostFocus
      txtQuantidade = adoResultado!Quantidade
      txtDescricao = adoResultado!Descricao
      txtDescricao = Replace(txtDescricao, "#", "'")
      txtPrecoUnitario = adoResultado!PUnitario
      cmdCadastrar.Visible = False
      cmdExcluir.Visible = True
      txtQuantidade.Enabled = False
      cboPedido.Enabled = False
   End If
   adoResultado.Close
End Function


Private Function VerificaNotaFiscal() As Boolean
   Dim adoResultado As New ADODB.Recordset
   adoResultado.Open "SELECT TOP 1 * FROM Tabela_Notas WHERE Pedido = " & Trim(txtPedido), cnaConexao, adOpenForwardOnly, adLockReadOnly
   If Not adoResultado.EOF Then
      VerificaNotaFiscal = True
   End If
   adoResultado.Close
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys "{TAB}"
   End If
End Sub

Private Sub Form_Load()
Centraliza Me
dtcNotas.DatabaseName = Caminho & nmBanco
End Sub

Private Sub txtNF_GotFocus()
  Dim adoResultado As New ADODB.Recordset
  
  adoResultado.Open "SELECT IIF(MAX(NotaFiscal) IS NULL, 0, MAX(NotaFiscal)) + 1 as NotaFiscal FROM Tabela_Notas", cnaConexao, adOpenForwardOnly, adLockReadOnly
  
  If Not adoResultado.EOF Then
     txtNF = adoResultado!NotaFiscal
     txtNF.SelStart = 0
     txtNF.SelLength = Len(Trim(txtNF))
  End If
  adoResultado.Close
  
  DadosLimpar
  txtQuantidade.Enabled = True
  cboPedido.Enabled = True
  cmdCadastrar.Visible = True
  cmdExcluir.Visible = False
  
End Sub

Private Sub txtNF_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
      If KeyAscii < 48 Or KeyAscii > 57 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtNF_LostFocus()
   If Len(Trim(txtNF)) > 0 Then
      DadosPreencher
            
   End If
End Sub

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
Private Sub DadosLimpar()
   
   cboPedido.Clear
   lblRazaoSocial.Caption = Space$(0)
   txtQuantidade = Space$(0)
   txtDescricao = Space$(0)
   txtPrecoUnitario = Space$(0)
   txtNF.SetFocus


End Sub
   
Private Function BuscaFornecedor()
Dim adoResultado   As New ADODB.Recordset
   
   If Len(Trim(cboPedido.Text)) > 0 Then
      adoResultado.Open "SELECT Tabela_Fornecedores.Fornecedor,Tabela_Fornecedores.RazaoSocial,Tabela_Pedidos.PUnitario,Tabela_Pedidos.Descricao FROM Tabela_Fornecedores,Tabela_Pedidos WHERE Tabela_Pedidos.Fornecedor = Tabela_Fornecedores.Fornecedor AND Tabela_Pedidos.Pedido = " & cboPedido.Text, cnaConexao, adOpenForwardOnly, adLockReadOnly
      If Not adoResultado.EOF Then
         lblRazaoSocial.Caption = adoResultado!Fornecedor & " - " & adoResultado!RazaoSocial
         txtDescricao = Replace(adoResultado!Descricao, "#", "'")
         txtPrecoUnitario = adoResultado!PUnitario
      Else
         cboPedido.Text = Space$(0)
      End If
      adoResultado.Close
   End If
End Function
Private Function VerificaSaldo() As Double
   Dim adoResultado  As New ADODB.Recordset
   
    adoResultado.Open "SELECT Saldo FROM Tabela_Pedidos WHERE Pedido = " & cboPedido.Text, cnaConexao, adOpenForwardOnly, adLockReadOnly
    
    If Not adoResultado.EOF Then
       VerificaSaldo = adoResultado!Saldo
    Else
       VerificaSaldo = 0
    End If
   
    adoResultado.Close
   

End Function
Private Function AtualizaDados()
   Dim intContador As Integer
   
   For intContador = 1 To 500
      dtcNotas.Refresh
   Next
End Function
