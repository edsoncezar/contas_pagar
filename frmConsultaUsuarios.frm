VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConsultaUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Usuários"
   ClientHeight    =   2010
   ClientLeft      =   2025
   ClientTop       =   1935
   ClientWidth     =   8040
   Icon            =   "frmConsultaUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   8040
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConsultaUsuarios.frx":0442
      Height          =   2010
      Left            =   0
      OleObjectBlob   =   "frmConsultaUsuarios.frx":045B
      TabIndex        =   0
      Top             =   0
      Width           =   8040
   End
   Begin VB.Data dtcUsuario 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\contas_pagar\dados\Base_CPagar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1170
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Consulta_Usuarios"
      Top             =   -30
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
dtcUsuario.DatabaseName = Caminho & nmBanco
End Sub
