VERSION 5.00
Begin VB.Form frmFuncionario 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cbLimpar 
      Caption         =   "Limpar"
      Height          =   735
      Left            =   3120
      TabIndex        =   7
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CommandButton cbCadastrar 
      Caption         =   "Cadastrar"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox tbSenha 
      Height          =   375
      Left            =   120
      MaxLength       =   12
      TabIndex        =   4
      Top             =   2400
      Width           =   5415
   End
   Begin VB.ComboBox coCargo 
      Height          =   315
      ItemData        =   "frmFuncionario.frx":0000
      Left            =   120
      List            =   "frmFuncionario.frx":0002
      TabIndex        =   3
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox tbNome 
      Height          =   375
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   10815
   End
   Begin VB.Label lbSenha 
      BackColor       =   &H8000000D&
      Caption         =   "Senha:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lbCargo 
      BackColor       =   &H8000000D&
      Caption         =   "Cargo"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lbNome 
      BackColor       =   &H8000000D&
      Caption         =   "Nome Completo:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCadastrar_Click()
Call ExportaBancoFuncionario
End Sub

Private Sub cbLimpar_Click()

tbNome.text = ""
coCargo.ListIndex = 0
tbSenha.text = ""

End Sub

Private Sub Form_Load()

con.Open strConn
rs.Open "SELECT nome FROM cargos", con, adOpenForwardOnly, adLockOptimistic
Do Until rs.EOF
    sEntrada = rs("nome")
    coCargo.AddItem sEntrada
    coCargo.ListIndex = 0
    rs.MoveNext
Loop
rs.Close
con.Close
frmTelaCadastros.Enabled = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmTelaCadastros.Enabled = True
End Sub

