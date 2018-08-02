VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acesso ao Sistema"
   ClientHeight    =   1920
   ClientLeft      =   5025
   ClientTop       =   4035
   ClientWidth     =   5655
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1134.399
   ScaleMode       =   0  'User
   ScaleWidth      =   5309.739
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   900
      TabIndex        =   1
      Top             =   945
      Width           =   2265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   4275
      TabIndex        =   3
      Top             =   960
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4275
      TabIndex        =   4
      Top             =   1380
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   900
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1380
      Width           =   2250
   End
   Begin VB.Frame Frame 
      Height          =   120
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Top             =   600
      Width           =   5385
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuário: "
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   7
      Top             =   960
      Width           =   630
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   195
      Picture         =   "frmLogin.frx":0442
      Top             =   165
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "SIP - Sistema Integrado"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   525
      Left            =   1020
      TabIndex        =   6
      Top             =   60
      Width           =   4425
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Senha: "
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   0
      Top             =   1380
      Width           =   555
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTmp As New ADODB.Recordset

Private Sub cmdCancel_Click()
   nmUsuario = Space$(0)
   Unload Me
End Sub
Private Sub cmdOK_Click()
Dim intUsuario As Integer
If CampoValidarTexto(txtUser) Then
   If CampoValidarTexto(txtPassword) Then
        strSQL = "Select Usuario, Perfil, Nome, Senha, Bloqueio from Tabela_Usuarios Where Usuario = " & txtUser
        If (rsTmp.State) = 0 Then
            rsTmp.Open "Tabela_Usuarios", cnaConexao, adOpenForwardOnly, adLockReadOnly
        End If
        Set rsTmp = cnaConexao.Execute(strSQL)
        If Not rsTmp.EOF And Not rsTmp.BOF Then
           intUsuario = rsTmp!Usuario
           If UCase(txtPassword) <> UCase(rsTmp!Senha) Then
              tentativas = tentativas + 1
              MsgBox "Senha Inválida!", , "Login"
              txtPassword.SetFocus
              rsTmp.Close
              SendKeys "{Home}+{End}"
           Else
                If rsTmp!bloqueio <> "1" Then
                    idUsuario = rsTmp!Usuario
                    perfilUsuario = Val(rsTmp!Perfil)
                    nmUsuario = rsTmp!Nome
                    Unload Me
                Else
                    MsgBox "Usuário bloqueado, entre em contato com o administrador!", , "Login"
                    bloqueio
                End If
           End If
        Else
           tentativas = tentativas + 1
           MsgBox "Usuário Inválido!", , "Login"
           txtUser.SetFocus
           rsTmp.Close
           SendKeys "{Home}+{End}"
        End If
     Else
         MsgBox "Digite a senha", , "Login"
         txtPassword.SetFocus
         SendKeys "{Home}+{End}"
     End If
 Else
    MsgBox "Digite o usuário.", , "Login"
    txtUser.SetFocus
    SendKeys "{Home}+{End}"
 End If
 If tentativas = 3 And intUsuario <> 0 Then
    MsgBox "Usuário bloqueado, foi gerado uma senha padrão, entre em contato com o administrador!", , "Login"
    cnaConexao.Execute ("Update Tabela_Usuarios Set Senha = '2003', Bloqueio = '1' WHERE Usuario = " & txtUser)
    bloqueio
 End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Centraliza Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
If nmUsuario = "" Then
   End
Else
    frmLogin.Hide
    mdiMenu.Show
    
End If
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
          If KeyAscii < 48 Or KeyAscii > 57 Then
             KeyAscii = 0
             MsgBox "Favor digitar apenas números!", vbCritical + vbOKOnly, "SIP - Sistema Integrado"
          End If
    End If
End Sub
Private Sub bloqueio()
    If isLogado Then
        mdiMenu.Hide
    End If
    Unload Me
    cnaConexao.Close
    End
End Sub
