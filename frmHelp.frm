VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuda do SIP - Sistema Integrado de Contas à Pagar"
   ClientHeight    =   6855
   ClientLeft      =   3630
   ClientTop       =   615
   ClientWidth     =   5550
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   5550
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   510
      Left            =   2153
      TabIndex        =   4
      Top             =   6225
      Width           =   1245
   End
   Begin VB.TextBox txtPesquisar 
      Height          =   295
      Left            =   1028
      TabIndex        =   0
      Top             =   745
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pesquisar"
      Height          =   975
      Left            =   788
      TabIndex        =   3
      Top             =   405
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Itens"
      Height          =   4215
      Left            =   788
      TabIndex        =   1
      Top             =   1740
      Width           =   3975
      Begin VB.ListBox lstItem 
         Height          =   3570
         ItemData        =   "frmHelp.frx":0442
         Left            =   240
         List            =   "frmHelp.frx":0444
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHelp As New ADODB.Recordset
Public Sub preencherDescricao()
If (rsHelp.State) = 0 Then
    rsHelp.Open "Tabela_Help", cnaConexao, adOpenForwardOnly, adLockReadOnly
End If
Set rsHelp = cnaConexao.Execute("SELECT IDHelp, Chave FROM Tabela_Help")
lstItem.Clear
While Not rsHelp.EOF
    lstItem.AddItem (Replace(rsHelp!Chave, "#", "'"))
    lstItem.ItemData(lstItem.NewIndex) = rsHelp!IDHelp
    rsHelp.MoveNext
Wend
End Sub
Private Sub cmdSair_Click()
rsHelp.Close
Unload Me
End Sub
Private Sub Form_Load()
Centraliza Me
preencherDescricao
End Sub
Private Sub lstItem_DblClick()
    IDHelp = lstItem.ItemData(lstItem.ListIndex)
    Unload frmHelp
    frmDescricao.Show
End Sub
Private Sub txtPesquisar_Change()
Chave = Replace(txtPesquisar, "'", "#")
Set rsHelp = cnaConexao.Execute("SELECT IDHelp, Chave FROM Tabela_Help WHERE Chave LIKE '" & Chave & "%'")
lstItem.Clear
While Not rsHelp.EOF
    lstItem.AddItem Replace(rsHelp!Chave, "#", "'")
    lstItem.ItemData(lstItem.NewIndex) = rsHelp!IDHelp
    rsHelp.MoveNext
Wend
End Sub
