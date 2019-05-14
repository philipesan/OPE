VERSION 5.00
Begin VB.Form frmPonto 
   Caption         =   "Cadastro de Ponto"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbFechamentoMins 
      Height          =   285
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   16
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox tbAberturaMins 
      Height          =   285
      Left            =   960
      MaxLength       =   2
      TabIndex        =   15
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox tbCep 
      Height          =   375
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cbLimpar 
      Caption         =   "Limpar"
      Height          =   735
      Left            =   4920
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cbCadastrar 
      Caption         =   "Cadastrar"
      Height          =   735
      Left            =   6120
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox tbFechamentoHrs 
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2520
      Width           =   375
   End
   Begin VB.Frame frFuncionamento 
      Caption         =   "Horário de funcionamento:"
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   4575
      Begin VB.ComboBox coFimDeSemana 
         Height          =   315
         ItemData        =   "frmPonto.frx":0000
         Left            =   2400
         List            =   "frmPonto.frx":0010
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox tbAberturaHrs 
         Height          =   285
         Left            =   240
         MaxLength       =   2
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lbSinal 
         Caption         =   "/"
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.ComboBox coGerente 
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox tbTelefone 
      Height          =   375
      Left            =   240
      MaxLength       =   11
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox tbNumero 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox tbEndereco 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label lbCep 
      Caption         =   "CEP:"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lbHorario 
      Caption         =   "Gerente:"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lbTelefone 
      Caption         =   "Telefone:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lbNumero 
      Caption         =   "Número:"
      Height          =   255
      Left            =   4920
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lbEndereco 
      Caption         =   "Endereço:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmPonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub cbCadastrar_Click()
Call ExportaBancoPonto
End Sub

Private Sub cbLimpar_Click()
tbEndereco.text = ""
tbNumero.text = ""
tbCep.text = ""
tbTelefone.text = ""
tbAberturaHrs.text = ""
tbAberturaMins.text = ""
tbFechamentoHrs.text = ""
tbFechamentoMins.text = ""

End Sub

Private Sub Form_Load()

con.Open strConn
rs.Open "SELECT nome, matricula FROM funcionarios", con, adOpenForwardOnly, adLockOptimistic
Do Until rs.EOF
    sEntrada = rs("nome") & " - " & rs("matricula")
    frmPonto.coGerente.AddItem sEntrada
    rs.MoveNext
Loop
rs.Close
con.Close

End Sub
Private Sub tbAberturaMins_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
        MsgBox "Apenas valores numéricos!", vbCritical
    End If
End Sub
Private Sub tbAberturaHrs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
        MsgBox "Apenas valores numéricos!", vbCritical
    End If
End Sub

Private Sub tbCep_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
        MsgBox "Apenas valores numéricos!", vbCritical
    End If
End Sub

Private Sub tbFechamentoHrs_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
        MsgBox "Apenas valores numéricos!", vbCritical
    End If
End Sub

Private Sub tbFechamentoMins_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
        MsgBox "Apenas valores numéricos!", vbCritical
    End If
End Sub
Private Sub tbNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
        MsgBox "Apenas valores numéricos!", vbCritical
    End If
End Sub
Private Sub tbTelefone_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
        MsgBox "Apenas valores numéricos!", vbCritical
    End If
End Sub
