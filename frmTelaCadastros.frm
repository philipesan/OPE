VERSION 5.00
Begin VB.Form frmTelaCadastros 
   BackColor       =   &H8000000D&
   Caption         =   "EAC - Estética Automotiva"
   ClientHeight    =   9360
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   13290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton addUser 
      Appearance      =   0  'Flat
      Caption         =   "Adicionar Usuario"
      Height          =   735
      Left            =   5760
      TabIndex        =   16
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cbVenda 
      Appearance      =   0  'Flat
      Caption         =   "Criar Venda"
      Height          =   735
      Left            =   840
      TabIndex        =   15
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Frame frVendas 
      BackColor       =   &H8000000D&
      Caption         =   "Vendas:"
      Height          =   1455
      Left            =   240
      TabIndex        =   14
      Top             =   4080
      Width           =   12735
   End
   Begin VB.CommandButton cbLogin 
      Caption         =   "Login / Logoff"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   8280
      Width           =   2895
   End
   Begin VB.CommandButton cbFuncionario 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "Funcionario"
      Height          =   735
      Left            =   840
      TabIndex        =   9
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cbServico 
      Appearance      =   0  'Flat
      Caption         =   "Serviço"
      Height          =   735
      Left            =   8280
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton tbCargo 
      Appearance      =   0  'Flat
      Caption         =   "Cargo"
      Height          =   735
      Left            =   3360
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton cbBanco 
      Caption         =   "Carregar Banco"
      Height          =   735
      Left            =   11400
      TabIndex        =   5
      Top             =   8400
      Width           =   1695
   End
   Begin VB.CommandButton cbCategoria 
      Appearance      =   0  'Flat
      Caption         =   "Categoria"
      Height          =   735
      Left            =   5880
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cbStatus 
      Appearance      =   0  'Flat
      Caption         =   "Status"
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Frame frCadastros 
      BackColor       =   &H8000000D&
      Caption         =   "Cadastros:"
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   12735
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
   Begin VB.Frame frBanco 
      BackColor       =   &H8000000D&
      Caption         =   "Banco de dados:"
      Height          =   855
      Left            =   3240
      TabIndex        =   4
      Top             =   8280
      Width           =   8055
      Begin VB.Label lbBanco 
         BackColor       =   &H8000000D&
         Caption         =   "Sem Banco de dados Carregado..."
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   7575
      End
   End
   Begin VB.Label lbRh 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   12120
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lbAdmin 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   11160
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lbSessao 
      Alignment       =   2  'Center
      Caption         =   "Nenhum Usuário"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   5055
   End
   Begin VB.Menu mbPrograma 
      Caption         =   "Programa"
      Index           =   0
      Begin VB.Menu mbLogin 
         Caption         =   "Login / Logoff"
      End
      Begin VB.Menu mbConecta 
         Caption         =   "Conectar Banco"
      End
   End
   Begin VB.Menu mbCadastros 
      Caption         =   "Cadastros"
      Index           =   1
      Begin VB.Menu mbFuncionario 
         Caption         =   "Funcionario"
      End
      Begin VB.Menu mbPonto 
         Caption         =   "Ponto"
      End
      Begin VB.Menu mbServico 
         Caption         =   "Serviço"
      End
      Begin VB.Menu mbCargo 
         Caption         =   "Cargo"
      End
      Begin VB.Menu mbCategoria 
         Caption         =   "Categoria"
      End
      Begin VB.Menu mbStatus 
         Caption         =   "Status"
      End
   End
End
Attribute VB_Name = "frmTelaCadastros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCategoria_Click()
If boolAutenticacao = True And lbAdmin = "Admin" Then
    frmCategoria.Show
Else
    MsgBox "Função não disponível para o usuário logado", vbCritical, "Erro"
End If
End Sub

Private Sub cbFuncionario_Click()
If boolAutenticacao = True And lbRh = "RH" Then
    frmFuncionario.Show
Else
    MsgBox "Função não disponível para o usuário logado", vbCritical, "Erro"
End If
End Sub

Private Sub cbLogin_Click()
    If boolAutenticacao <> True Then
        frmLogin.Show
    Else
        If MsgBox("Deseja Fazer Logoff", vbYesNo, "Confirmação") = vbYes Then
            Call RealizaLogoff
            MsgBox "Usuário deslogado com sucesso", vbOKOnly, "Logoff"
        End If
    End If
End Sub

Private Sub cbPonto_Click()
If boolAutenticacao = True And lbAdmin = "Admin" Then
    frmPonto.Show
Else
    MsgBox "Função não disponível para o usuário logado", vbCritical, "Erro"
End If
End Sub


Private Sub cbServico_Click()
If boolAutenticacao = True And lbAdmin = "Admin" Then
    frmServico.Show
Else
    MsgBox "Função não disponível para o usuário logado", vbCritical, "Erro"
End If
End Sub

Private Sub cbStatus_Click()
If boolAutenticacao = True And lbAdmin = "Admin" Then
    frmStatus.Show
Else
    MsgBox "Função não disponível para o usuário logado", vbCritical, "Erro"
End If
End Sub

Private Sub cbVenda_Click()
If boolAutenticacao = True Then
    frmCriarVenda.Show
Else
    MsgBox "Função não disponível para o usuário logado", vbCritical, "Erro"
End If

End Sub

Private Sub Form_Load()
lbBanco.Caption = App.Path & "\" & "database" & "\" & "database.mdb"
sFilePath = lbBanco.Caption
sLogPath = App.Path & "\" & "log.txt"
Call CriaBanco(sFilePath)
End Sub



Private Sub mbLogin_Click()
    If boolAutenticacao <> True Then
        frmLogin.Show
    Else
        If MsgBox("Deseja Fazer Logoff", vbYesNo, "Confirmação") = vbYes Then
            boolAutenticacao = False
            frmTelaCadastros.lbSessao.Caption = "Nenhum Usuário"
            frmTelaCadastros.lbAdmin.Caption = ""
            frmTelaCadastros.lbRh.Caption = ""
            MsgBox "Usuário deslogado com sucesso", vbOKOnly, "Logoff"
        End If
    End If
    End Sub

Private Sub mbPonto_Click()
If boolAutenticacao = True And frmSessionControl.lbAdmin.Caption = "Admin" Then
    frmPonto.Show
End If
End Sub

Private Sub tbCargo_Click()
If boolAutenticacao = True And lbRh = "RH" Then
    frmCargo.Show
Else
    MsgBox "Função não disponível para o usuário logado", vbCritical, "Erro"
End If
End Sub
