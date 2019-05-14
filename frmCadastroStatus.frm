VERSION 5.00
Begin VB.Form frmStatus 
   Caption         =   "Cadastro Status"
   ClientHeight    =   780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cdCadastrar 
      Caption         =   "Cadastrar"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox tbStatus 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cdCadastrar_Click()
Call ExportaBancoStatus
End Sub
