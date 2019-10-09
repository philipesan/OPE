VERSION 5.00
Begin VB.Form frmStatus 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Status"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cdCadastrar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cadastrar"
      Height          =   495
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
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

Private Sub Form_Load()
    frmTelaCadastros.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmTelaCadastros.Enabled = True
End Sub
