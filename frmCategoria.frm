VERSION 5.00
Begin VB.Form frmCategoria 
   Caption         =   "Cadastro de Categoria"
   ClientHeight    =   1245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
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
      Caption         =   "Adicional:"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbnome 
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
Call ExportaBancoCategoria
End Sub

Private Sub cbLimpar_Click()
tbAdicional.text = ""
tbNome.text = ""

End Sub

Private Sub tbAdicional_LostFocus()
    tbAdicional.text = Format(tbAdicional.text, "#####0.00")
End Sub



