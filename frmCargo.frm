VERSION 5.00
Begin VB.Form frmCargo 
   Caption         =   "Cadastro de Cargos"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cbCadastrar 
      Caption         =   "Cadastrar"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cbLimpar 
      Caption         =   "Limpar"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CheckBox ckRh 
      Caption         =   "Acesso RH"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CheckBox ckAdmin 
      Caption         =   "Acesso Administrador"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox tbSalario 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox tbNome 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label lbSalario 
      Caption         =   "Salário:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label lbNome 
      Caption         =   "Nome: "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCadastrar_Click()
Call ExportaBancoCargo
End Sub

Private Sub cbLimpar_Click()
tbSalario.text = ""
tbNome.text = ""
ckAdmin.Value = 0
ckRh.Value = 0

End Sub

Private Sub tbSalario_LostFocus()
    tbSalario.text = Format(tbSalario.text, "#####0.00")
End Sub




