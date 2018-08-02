VERSION 5.00
Begin VB.Form frmDescricao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descrição do Item"
   ClientHeight    =   3120
   ClientLeft      =   5790
   ClientTop       =   4320
   ClientWidth     =   4530
   Icon            =   "frmDescricaoSobre.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4530
   Begin VB.ListBox lstDescricao 
      Height          =   1620
      ItemData        =   "frmDescricaoSobre.frx":0442
      Left            =   315
      List            =   "frmDescricaoSobre.frx":0444
      TabIndex        =   2
      Top             =   450
      Width           =   3900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descrição"
      Height          =   2220
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   4320
   End
   Begin VB.CommandButton cmdVoltar 
      Caption         =   "&Voltar"
      Height          =   435
      Left            =   1650
      TabIndex        =   0
      Top             =   2490
      Width           =   1170
   End
End
Attribute VB_Name = "frmDescricao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHelp As New ADODB.Recordset
Private Sub cmdVoltar_Click()
Unload frmDescricao
frmHelp.Show
End Sub
Private Sub Form_Load()
Centraliza Me
If (rsHelp.State) = 0 Then
    rsHelp.Open "Tabela_Help", cnaConexao, adOpenForwardOnly, adLockReadOnly
End If
Set rsHelp = cnaConexao.Execute("SELECT Descricao FROM Tabela_Help WHERE IDHelp = " & IDHelp)
lstDescricao.Clear
While Not rsHelp.EOF
    If Len(rsHelp!Descricao) > 45 Then
       lstDescricao.AddItem (Mid(rsHelp!Descricao, 1, 45))
       lstDescricao.AddItem (Mid(rsHelp!Descricao, 46, 45))
       lstDescricao.AddItem (Mid(rsHelp!Descricao, 91, 45))
       lstDescricao.AddItem (Mid(rsHelp!Descricao, 136, 45))
       lstDescricao.AddItem (Mid(rsHelp!Descricao, 181, 45))
       lstDescricao.AddItem (Mid(rsHelp!Descricao, 226, 45))
       lstDescricao.AddItem (Mid(rsHelp!Descricao, 271, 45))
       lstDescricao.AddItem (Mid(rsHelp!Descricao, 316, 45))
    Else
       lstDescricao.AddItem (rsHelp!Descricao)
    End If
    rsHelp.MoveNext
Wend
    
End Sub

