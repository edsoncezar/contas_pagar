VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConsultaPedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Pedidos"
   ClientHeight    =   2160
   ClientLeft      =   3300
   ClientTop       =   4815
   ClientWidth     =   9600
   Icon            =   "frmConsultaPedidos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   9600
   Begin MSDBGrid.DBGrid GRIDPedidos 
      Bindings        =   "frmConsultaPedidos.frx":0442
      Height          =   2085
      Left            =   0
      OleObjectBlob   =   "frmConsultaPedidos.frx":045A
      TabIndex        =   0
      Top             =   0
      Width           =   9585
   End
   Begin VB.Data dtcPedido 
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\contas a pagar\dados\Base_CPagar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5310
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "Consulta_Pedidos"
      Top             =   30
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "frmConsultaPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Centraliza Me
dtcPedido.DatabaseName = Caminho & nmBanco
End Sub
