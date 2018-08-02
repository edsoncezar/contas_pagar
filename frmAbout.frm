VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre  o SIP"
   ClientHeight    =   2730
   ClientLeft      =   4980
   ClientTop       =   3480
   ClientWidth     =   5550
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5550
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turma dos Alunos do 1º F Sala 4 - 2003     ID: 17169-24765-01247-38380"
      Height          =   555
      Left            =   885
      TabIndex        =   2
      Top             =   1335
      Width           =   3075
   End
   Begin VB.Label Label3 
      Caption         =   "Aviso: Este programa está protegido pela política de privacidade."
      Height          =   345
      Left            =   293
      TabIndex        =   1
      Top             =   2025
      Width           =   4965
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0442
      Height          =   870
      Left            =   886
      TabIndex        =   0
      Top             =   360
      Width           =   3585
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Centraliza Me
End Sub

