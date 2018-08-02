VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMenu 
   BackColor       =   &H8000000C&
   Caption         =   "SIP - Sistema Integrado de Contas à Pagar"
   ClientHeight    =   5280
   ClientLeft      =   3630
   ClientTop       =   2865
   ClientWidth     =   8520
   Icon            =   "mdiMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stMenu 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Nome do usuário logado."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "06/11/2003"
            Object.ToolTipText     =   "Data"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "13:18"
            Object.ToolTipText     =   "Hora"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuUsuarios 
      Caption         =   "&Usuários"
      Begin VB.Menu mnuAbrirUSU 
         Caption         =   "&Abrir"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuConsultaUsuario 
         Caption         =   "&Consulta"
      End
   End
   Begin VB.Menu mnuFornecedores 
      Caption         =   "&Fornecedores"
      Begin VB.Menu mnuAbrirFOR 
         Caption         =   "&Abrir   "
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuConsultaFornecedores 
         Caption         =   "&Consulta"
      End
   End
   Begin VB.Menu mnuPedido 
      Caption         =   "&Pedidos"
      Begin VB.Menu mnuAbrirPED 
         Caption         =   "&Abrir"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuConsultaPedido 
         Caption         =   "&Consulta"
      End
   End
   Begin VB.Menu mnuNotas 
      Caption         =   "&Notas Fiscais"
      Begin VB.Menu mnuAbrirNF 
         Caption         =   "&Abrir"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuConsultaNF 
         Caption         =   "&Consulta"
      End
   End
   Begin VB.Menu mnuSairSIS 
      Caption         =   "&Sair"
      Begin VB.Menu mnuTrocarUSU 
         Caption         =   "&Trocar de Usuário"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuTracoSAIR 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "&Sair do Sistema"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuIndice 
         Caption         =   "&Índice"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuTracoAjuda 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "&Sobre SIP(Sistema Integrado)"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "mdiMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Activate()
stMenu.Panels(1) = "Usuário: " & nmUsuario
If perfilUsuario = 1 Then
    mnuUsuarios.Visible = False
Else
    mnuUsuarios.Visible = True
End If
End Sub

Private Sub MDIForm_Load()
Centraliza Me
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
cnaConexao.Close
End
End Sub

Private Sub mnuAbrirFOR_Click()
frmFornecedores.Show
End Sub

Private Sub mnuAbrirNF_Click()
frmNotasFiscais.Show
End Sub

Private Sub mnuAbrirPED_Click()
frmPedidoCompra.Show
End Sub

Private Sub mnuAbrirUSU_Click()
frmUsuarios.Show
End Sub

Private Sub mnuConsultaFornecedores_Click()
frmConsultaFornecedores.Show
End Sub

Private Sub mnuConsultaNF_Click()
frmConsultaNotas.Show
End Sub

Private Sub mnuConsultaPedido_Click()
frmConsultaPedidos.Show
End Sub

Private Sub mnuConsultaUsuario_Click()
frmConsultaUsuarios.Show
End Sub

Private Sub mnuIndice_Click()
frmHelp.Show
End Sub

Private Sub mnuSair_Click()
End
End Sub

Private Sub mnuSobre_Click()
frmAbout.Show
End Sub

Private Sub mnuTrocarUSU_Click()
frmLogin.Show
isLogado = True
End Sub
