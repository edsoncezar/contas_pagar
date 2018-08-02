VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConsultaNotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Notas"
   ClientHeight    =   1590
   ClientLeft      =   3165
   ClientTop       =   5025
   ClientWidth     =   9600
   Icon            =   "frmConsultaNotas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   9600
   Begin MSDBGrid.DBGrid GRIDNotas 
      Bindings        =   "frmConsultaNotas.frx":0442
      Height          =   1545
      Left            =   30
      OleObjectBlob   =   "frmConsultaNotas.frx":0459
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Data dtcNotas 
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
      RecordSource    =   "Consulta_Notas"
      Top             =   30
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "frmConsultaNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Centraliza Me
dtcNotas.DatabaseName = Caminho & nmBanco
End Sub

