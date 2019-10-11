VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFuncionario 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cbDocumentos 
      Caption         =   "Documentos Digitalizados"
      Height          =   855
      Left            =   120
      TabIndex        =   39
      Top             =   6600
      Width           =   6255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4680
      MaxLength       =   11
      TabIndex        =   37
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox tbAgencia 
      Height          =   375
      Left            =   2400
      MaxLength       =   11
      TabIndex        =   35
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox cbBanco 
      Height          =   375
      Left            =   120
      MaxLength       =   11
      TabIndex        =   34
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cbAddFoto 
      Caption         =   "Adicionar Foto"
      Height          =   495
      Left            =   7560
      TabIndex        =   32
      Top             =   7320
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   6960
      ScaleHeight     =   2835
      ScaleWidth      =   2115
      TabIndex        =   31
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox tbEmail 
      Height          =   375
      Left            =   3600
      MaxLength       =   9
      TabIndex        =   29
      Top             =   4680
      Width           =   2775
   End
   Begin VB.TextBox tbTelefone 
      Height          =   375
      Left            =   120
      MaxLength       =   11
      TabIndex        =   28
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CheckBox lbAtivo 
      BackColor       =   &H8000000D&
      Caption         =   "Funcionário Ativo"
      Height          =   195
      Left            =   7560
      TabIndex        =   26
      Top             =   3720
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dpDtNasc 
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   68288513
      CurrentDate     =   43747
   End
   Begin VB.TextBox tbPis 
      Height          =   375
      Left            =   120
      MaxLength       =   11
      TabIndex        =   19
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox tbSerie 
      Height          =   375
      Left            =   7080
      MaxLength       =   11
      TabIndex        =   16
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox cbEstado 
      Height          =   315
      Left            =   8880
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox tbCTPS 
      Height          =   375
      Left            =   5400
      MaxLength       =   11
      TabIndex        =   12
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox tbCPF 
      Height          =   375
      Left            =   2760
      MaxLength       =   11
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox tbRG 
      Height          =   375
      Left            =   120
      MaxLength       =   9
      TabIndex        =   8
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton cbLimpar 
      Caption         =   "Limpar"
      Height          =   735
      Left            =   3600
      TabIndex        =   7
      Top             =   7800
      Width           =   2775
   End
   Begin VB.CommandButton cbCadastrar 
      Caption         =   "Cadastrar"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   7800
      Width           =   2775
   End
   Begin VB.TextBox tbSenha 
      Height          =   375
      Left            =   4080
      MaxLength       =   12
      TabIndex        =   4
      Top             =   1560
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
      Width           =   9375
   End
   Begin MSComCtl2.DTPicker DtContra 
      Height          =   375
      Left            =   3840
      TabIndex        =   22
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   35848193
      CurrentDate     =   43747
   End
   Begin MSComCtl2.DTPicker DtDemissao 
      Height          =   375
      Left            =   5760
      TabIndex        =   24
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   35848193
      CurrentDate     =   43747
   End
   Begin VB.Label lbConta 
      BackColor       =   &H8000000D&
      Caption         =   "Conta:"
      Height          =   375
      Left            =   4800
      TabIndex        =   38
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lbAgencia 
      BackColor       =   &H8000000D&
      Caption         =   "Agência:"
      Height          =   375
      Left            =   2520
      TabIndex        =   36
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lbBanco 
      BackColor       =   &H8000000D&
      Caption         =   "Banco:"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lbEmail 
      BackColor       =   &H8000000D&
      Caption         =   "Telefone:"
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lbTelefone 
      BackColor       =   &H8000000D&
      Caption         =   "Telefone:"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Data de Demissão:"
      Height          =   255
      Left            =   5760
      TabIndex        =   25
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lbContra 
      BackColor       =   &H8000000D&
      Caption         =   "Data de Contratação:"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lbDtNasc 
      BackColor       =   &H8000000D&
      Caption         =   "Data de Nascimento:"
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lbUF 
      BackColor       =   &H8000000D&
      Caption         =   "UF:"
      Height          =   255
      Left            =   8880
      TabIndex        =   15
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lbSerie 
      BackColor       =   &H8000000D&
      Caption         =   "Série:"
      Height          =   255
      Left            =   7080
      TabIndex        =   17
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lbCTPS 
      BackColor       =   &H8000000D&
      Caption         =   "CTPS:"
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lbCPF 
      BackColor       =   &H8000000D&
      Caption         =   "CPF:"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lbRG 
      BackColor       =   &H8000000D&
      Caption         =   "RG:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lbPIS 
      BackColor       =   &H8000000D&
      Caption         =   "PIS:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lbSenha 
      BackColor       =   &H8000000D&
      Caption         =   "Senha:"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
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

Private Sub Label2_Click()

End Sub

Private Sub lbCert_Click()

End Sub

Private Sub Text3_Change()

End Sub

