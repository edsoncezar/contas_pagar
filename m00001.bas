Attribute VB_Name = "m00001"
Option Explicit
 'Declaração das variáveis globais do sistema
Public cnaConexao As New ADODB.Connection
Public strSQL, nmUsuario, Caminho, Chave, nmBanco As String
Public IDHelp, perfilUsuario, idUsuario, tentativas As Integer
Public isLogado As Boolean
Dim Pausa, Inicio As Double
'Abertura de banco de dados
Public Function AbrirBD()
    cnaConexao.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & Caminho & nmBanco & ";"
End Function
'Iniciar o sistema
Public Sub Main()
 ' Tela de abertura do sistema
 frmSplash.Show
 frmSplash.Refresh
 
 'Nome do banco de dados
 nmBanco = "\Base_CPagar.mdb"
 
 ' Montar o caminho para o banco de dados do sistema
 montarCaminho
 
 'Abrir a conexão com banco de dados
 AbrirBD
  
 'Descarregar a tela de abertura
Pausa = 1   ' Duração.
Inicio = Timer   ' Começo.
Do While Timer < Inicio + Pausa
   DoEvents
Loop
Unload frmSplash

'Carregar a tela de login
frmLogin.Show vbModal
frmLogin.Refresh
End Sub
Public Function CampoValidarTexto(strCampo As String) As Boolean
   CampoValidarTexto = True
   If Len(Trim(strCampo)) = 0 Then
      CampoValidarTexto = False
  End If
End Function
Public Sub Centraliza(frm As Form)
With mdiMenu
    If frm.WindowState = vbNormal Then
        If TypeOf frm Is MDIForm Then
            frm.Top = (Screen.Height - frm.Height) / 2
            frm.Left = (Screen.Width - frm.Width) / 2
        Else
            If frm.MDIChild = True Then
                frm.Top = (.ScaleHeight - frm.Height) / 2
                frm.Left = (.ScaleWidth - frm.Width) / 2
            Else
                frm.Top = (Screen.Height - frm.Height) / 2
                frm.Left = (Screen.Width - frm.Width) / 2
            End If
        End If
    End If
End With
End Sub
Public Sub montarCaminho()
Caminho = LCase(App.Path)
If Right$(App.Path, 1) <> "\" Then
    Caminho = Caminho + "\dados"
End If
End Sub



