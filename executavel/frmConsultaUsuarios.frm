VERSION 5.00
Begin VB.Form frmConsultaUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Fornecedores"
   ClientHeight    =   1755
   ClientLeft      =   4050
   ClientTop       =   5235
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   6465
   Begin VB.PictureBox GRIDFornecedores 
      Height          =   1725
      Left            =   -15
      ScaleHeight     =   1665
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   -15
      Width           =   6495
   End
   Begin VB.Data dtcFornecedores 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Contas a Pagar\dados\Base_CPagar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3705
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Tabela_Fornecedores"
      Top             =   -15
      Visible         =   0   'False
      Width           =   1230
   End
End
Attribute VB_Name = "frmConsultaUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Centraliza Me
   dtcFornecedores.DatabaseName = Caminho & nmBanco
End Sub

