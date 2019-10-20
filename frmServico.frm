VERSION 5.00
Begin VB.Form frmServico 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Serviço"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cbCadastrar 
      Caption         =   "Cadastrar"
      Height          =   735
      Left            =   4560
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cbLimpar 
      Caption         =   "Limpar"
      Height          =   735
      Left            =   6360
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox tbValor 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "0000,00"
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox tbDescricao 
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   7695
   End
   Begin VB.TextBox tbNome 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   7695
   End
   Begin VB.Label lbValor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Valor Base:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lbDescricao 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Descricao:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lbNome 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Nome:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCadastrar_Click()
'Checa se os campos estão preenchidos com os dados corretos

If IsNumeric(tbValor.text) Then
    Call ExportaBancoServico
Else
    MsgBox "O campo Valor é apenas numérico", vbCritical, "Erro"
    tbValor.text = ""
End If

End Sub
Private Sub cbLimpar_Click()
tbNome.text = ""
tbDescricao.text = ""
tbValor.text = ""

End Sub

Private Sub Form_Load()
    frmTelaCadastros.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmTelaCadastros.Enabled = True
End Sub


Private Sub tbValor_LostFocus()
tbValor.text = Format(tbValor.text, "#####0.00")
End Sub
