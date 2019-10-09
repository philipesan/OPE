VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cbLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox tbSenha 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "Senha"
      Top             =   960
      Width           =   3975
   End
   Begin VB.Frame frLogin 
      BackColor       =   &H8000000D&
      Caption         =   "Login"
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.TextBox tbMatricula 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "Matrícula"
         Top             =   360
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbLogin_Click()
    Dim strMatricula As String, strSenha As String
    strMatricula = tbMatricula.text
    strSenha = tbSenha.text
    boolAutenticacao = ChecaAutenticacao(strMatricula, strSenha)
    If boolAutenticacao = True Then
        MsgBox "Usuário autenticado com sucesso!", vbOKOnly, "Login"
    Else
        MsgBox "Usuário ou senha inválidos", vbCritical, "Login"
    End If
    
    Unload Me

End Sub

Private Sub Form_Load()
    frmTelaCadastros.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmTelaCadastros.Enabled = True
End Sub

