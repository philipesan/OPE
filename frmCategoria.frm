VERSION 5.00
Begin VB.Form frmCategoria 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Categoria"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7335
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cbCadastrar 
      Caption         =   "Cadastrar"
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cbLimpar 
      Caption         =   "Limpar"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox tbAdicional 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Text            =   "0000,00"
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox tbNome 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label lbAdicional 
      BackColor       =   &H8000000D&
      Caption         =   "Adicional:"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbnome 
      BackColor       =   &H8000000D&
      Caption         =   "Nome:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCadastrar_Click()
'Checa se os campos estão preenchidos com os dados corretos

If IsNumeric(tbAdicional.text) Then
    Call ExportaBancoCategoria
Else
    MsgBox "O campo Adicional é apenas numérico", vbCritical, "Erro"
    tbAdicional.text = ""
End If

End Sub

Private Sub cbLimpar_Click()
tbAdicional.text = ""
tbNome.text = ""

End Sub

Private Sub Form_Load()
    frmTelaCadastros.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmTelaCadastros.Enabled = True
End Sub



Private Sub tbAdicional_LostFocus()
    tbAdicional.text = Format(tbAdicional.text, "#####0.00")
End Sub



