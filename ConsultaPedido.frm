VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   2025
   ClientTop       =   1935
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8430
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3255
      TabIndex        =   3
      Top             =   2940
      Width           =   2160
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2325
      TabIndex        =   2
      Top             =   480
      Width           =   4230
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ConsultaPedido.frx":0000
      Left            =   510
      List            =   "ConsultaPedido.frx":000D
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   450
      Width           =   1725
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ConsultaPedido.frx":0033
      Height          =   1395
      Left            =   60
      OleObjectBlob   =   "ConsultaPedido.frx":0047
      TabIndex        =   0
      Top             =   1200
      Width           =   8025
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Contas a Pagar\Base_CPagar.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1170
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   2  'Snapshot
      RecordSource    =   "ConsultaGenerica_Pedidos"
      Top             =   3180
      Width           =   1260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub
