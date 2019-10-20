VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRelatorios 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Relatórios"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opFuncionarios 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Relatórios de Funcionários"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.OptionButton opVendas 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Relatórios de Vendas"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtTeto 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   143982593
      CurrentDate     =   43749
   End
   Begin VB.Frame frRelatorios 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Relatórios"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   7095
      Begin VB.CommandButton cbGerar 
         Caption         =   "Gerar"
         Height          =   615
         Left            =   5280
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtPiso 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   143982595
         CurrentDate     =   43749
      End
   End
End
Attribute VB_Name = "frmRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbGerar_Click()

    Dim varPiso, varTeto As Date

    varPiso = CDate(dtPiso.value)
    varTeto = CDate(dtTeto.value)


    'Checa qual Opção está marcada
    If opVendas.value = True Then
        Call RelatorioVendas(varPiso, varTeto)
    Else
        Call RelatorioFuncionarios
    End If

End Sub

