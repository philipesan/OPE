VERSION 5.00
Begin VB.Form frmTelaCadastros 
   Caption         =   "EAC - Tela de Cadastros"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton tbCargo 
      Appearance      =   0  'Flat
      Caption         =   "Cargo"
      Height          =   735
      Left            =   8400
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cbBanco 
      Caption         =   "Carregar Banco"
      Height          =   735
      Left            =   8880
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Frame frBanco 
      Caption         =   "Banco de dados:"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   8535
      Begin VB.Label lbBanco 
         Caption         =   "Sem Banco de dados Carregado..."
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   7935
      End
   End
   Begin VB.CommandButton cbCategoris 
      Appearance      =   0  'Flat
      Caption         =   "Categoria"
      Height          =   735
      Left            =   5880
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton cbStatus 
      Appearance      =   0  'Flat
      Caption         =   "Status"
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Frame frCadastros 
      Caption         =   "Cadastros:"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   10575
      Begin VB.CommandButton cbPonto 
         Appearance      =   0  'Flat
         Caption         =   "Ponto"
         Height          =   735
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmTelaCadastros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCategoris_Click()
frmCategoria.Show
End Sub

Private Sub cbPonto_Click()
frmPonto.Show
End Sub



Private Sub Command1_Click()

End Sub

Private Sub cbStatus_Click()
frmStatus.Show
End Sub

Private Sub Form_Load()
lbBanco.Caption = App.Path & "\" & "database" & "\" & "database.mdb"
sFilePath = lbBanco.Caption
sLogPath = App.Path & "\" & "log.txt"
Call CriaBanco(sFilePath)
End Sub

Private Sub tbCargo_Click()
frmCargo.Show

End Sub
