VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmCriarVenda 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Criar Venda"
   ClientHeight    =   10605
   ClientLeft      =   5550
   ClientTop       =   180
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   11475
   Begin VB.CommandButton cbConcluir 
      Caption         =   "Concluir Ordem"
      Height          =   615
      Left            =   240
      TabIndex        =   34
      Top             =   9720
      Width           =   11055
   End
   Begin VB.CommandButton cbLimpar 
      Caption         =   "Limpar Campos"
      Height          =   375
      Left            =   2040
      TabIndex        =   31
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cbRemover 
      Caption         =   "Remover"
      Height          =   735
      Left            =   7200
      TabIndex        =   30
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cbLimparGrid 
      Caption         =   "Limpar Grid"
      Height          =   615
      Left            =   5280
      TabIndex        =   29
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton cbAdicionar 
      Caption         =   "Adicionar"
      Height          =   735
      Left            =   9240
      TabIndex        =   28
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cbAtualizaPreco 
      Caption         =   "Atualizar Preco"
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   3840
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid fgVendaItens 
      Height          =   3255
      Left            =   360
      TabIndex        =   24
      Top             =   5040
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      GridColor       =   -2147483645
      Appearance      =   0
   End
   Begin VB.CheckBox ckAdicional1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Adicional: R$"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox tbDesconto1 
      Height          =   285
      Left            =   7440
      TabIndex        =   21
      Top             =   3000
      Width           =   1695
   End
   Begin VB.ComboBox coServico1 
      Height          =   315
      Left            =   1080
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox tbEmail 
      Height          =   285
      Left            =   8160
      TabIndex        =   17
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox tbTelefone 
      Height          =   285
      Left            =   5640
      MaxLength       =   11
      TabIndex        =   15
      Text            =   "11999999999"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox tbDocumento 
      Height          =   285
      Left            =   1320
      MaxLength       =   9
      TabIndex        =   13
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Frame frCliente 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dados do Cliente:"
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   11055
      Begin VB.TextBox tbNome 
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   10095
      End
      Begin VB.Label lbEmail 
         BackColor       =   &H00FFC0C0&
         Caption         =   "E-mail:"
         Height          =   255
         Left            =   7440
         TabIndex        =   16
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbTelefone 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Telefone:"
         Height          =   255
         Left            =   4680
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbDocumento 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Documento: "
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lbNome 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nome:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.TextBox tbModelo 
      Height          =   285
      Left            =   5880
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox tbMarca 
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.Frame frAutomovel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dados do Automóvel:"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11055
      Begin VB.ComboBox coCategoria 
         Height          =   315
         Left            =   8400
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox tbPlaca 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbCategoria 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Categoria: "
         Height          =   255
         Left            =   7560
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lbModelo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modelo:"
         Height          =   255
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbMarca 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Marca:"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbPlaca 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Placa:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label lbTotalValor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8400
      TabIndex        =   33
      Top             =   8400
      Width           =   2655
   End
   Begin VB.Label lbTotal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Total: R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      TabIndex        =   32
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   120
      X2              =   11280
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Label lbSubvalor 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   10320
      TabIndex        =   26
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lbSubtotal 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Subtotal: R$"
      Height          =   255
      Left            =   9240
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.Line lnLine 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      X1              =   120
      X2              =   11280
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label lbAdicional1 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lbDesconto1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Desconto: "
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lbServico1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Serviço:"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   3000
      Width           =   615
   End
End
Attribute VB_Name = "frmCriarVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbAdicionar_Click()
    Dim SubTotal, desconto, adicional As Double
    Dim Servico As Integer
    Dim ServicoNome As String
    
    
    If coServico1.ListIndex = 0 Then
        MsgBox "Você deve selecionar um serviço antes", vbCritical, "Erro"
    Else
        If coCategoria.ListIndex = 0 Then
            MsgBox "Você deve selecionar uma categoria antes", vbCritical, "Erro"
        Else
            SubTotal = CDbl(Mid(coServico1.text, InStrRev(coServico1.text, "$") + 1))
            Servico = Val(Left(coServico1.text, InStr(coServico1.text, "-") - 1))
            ServicoNome = Mid(coServico1.text, InStr(coServico1.text, "-") + 1, InStrRev(coServico1.text, "R") - 4)
            If ckAdicional1.value = 1 Then
                SubTotal = SubTotal + CDbl(lbAdicional1.Caption)
            End If
            If IsNumeric(tbDesconto1.text) Then
                SubTotal = SubTotal - CDbl(tbDesconto1.text)
                desconto = tbDesconto1.text
            End If
            lbSubvalor.Caption = SubTotal
            fgVendaItens.AddItem Servico & Chr(9) & ServicoNome & Chr(9) & adicional & Chr(9) & desconto & Chr(9) & SubTotal & Chr(9)
            lbTotalValor.Caption = Format(CDbl(lbTotalValor.Caption) + SubTotal, "#####0.00")
        End If
    End If
End Sub

Private Sub cbAtualizaPreco_Click()
    Dim SubTotal As Double
    If coServico1.ListIndex = 0 Then
        MsgBox "Você deve selecionar um serviço antes", vbCritical, "Erro"
    Else
        If coCategoria.ListIndex = 0 Then
            MsgBox "Você deve selecionar uma categoria antes", vbCritical, "Erro"
        Else
            SubTotal = CDbl(Mid(coServico1.text, InStrRev(coServico1.text, "$") + 1))
            If ckAdicional1.value = 1 Then
                SubTotal = Format(SubTotal + CDbl(lbAdicional1.Caption), "#####0.00")
            End If
            If IsNumeric(tbDesconto1.text) Then
                SubTotal = Format(SubTotal - CDbl(tbDesconto1.text), "#####0.00")
            End If
            lbSubvalor.Caption = SubTotal
        End If
    End If
End Sub

Private Sub cbConcluir_Click()
    
    Dim id_ordem As Integer 'Variavel que recebe o ID da ordem de serviço
    Dim max_rows As Integer 'Variavel que recebe o número máximo de colunas
    Dim id_linha As Integer, id_servico As Integer
    Dim valor As Double, desconto As Double, adicional As Double
    
    'Pega a quantidade de colunas
    max_rows = fgVendaItens.Rows
    
    'Fazendo a inserção do cabeçalho da ordem
    id_ordem = ExportaBancoOrdem
    
    'Percorre o FlexGrid
    For id_linha = 1 To max_rows - 1
        
        'Pega os valores
        id_servico = CInt(fgVendaItens.TextMatrix(id_linha, 0))
        adicional = CDbl(fgVendaItens.TextMatrix(id_linha, 2))
        desconto = CDbl(IIf(fgVendaItens.TextMatrix(id_linha, 3) <> "", fgVendaItens.TextMatrix(id_linha, 3), 0))
        valor = CDbl(fgVendaItens.TextMatrix(id_linha, 4))
        
        'Insere a linha no banco
        Call ExportaBancoOrdemServicos(id_ordem, id_linha, id_servico, valor, desconto, adicional)
        
    Next
    
    MsgBox "Venda cadastrada com sucesso!"
    Call cbLimpar_Click
    Call cbLimparGrid_Click
End Sub

Private Sub cbLimpar_Click()
    tbPlaca.text = ""
    tbMarca.text = ""
    tbModelo.text = ""
    tbNome.text = ""
    tbDocumento.text = ""
    tbTelefone.text = ""
    tbEmail.text = ""
End Sub

Private Sub cbLimparGrid_Click()
    fgVendaItens.Clear
    fgVendaItens.Rows = 1
    fgVendaItens.Row = 0
    fgVendaItens.Col = 0
    fgVendaItens.text = "Serviço"
    fgVendaItens.Col = 1
    fgVendaItens.text = "Nome do Serviço"
    fgVendaItens.Col = 2
    fgVendaItens.text = "Adicional"
    fgVendaItens.Col = 3
    fgVendaItens.text = "Desconto"
    fgVendaItens.Col = 4
    fgVendaItens.text = "Valor total"
    lbTotalValor.Caption = "000,00"
    

End Sub

Private Sub cbRemover_Click()
    Dim Subtrai As Double
    
    If fgVendaItens.Rows = 2 Then
        Call cbLimparGrid_Click
    Else
        Subtrai = fgVendaItens.TextMatrix(fgVendaItens.Row, 4)
        lbTotalValor.Caption = Format(CDbl(lbTotalValor.Caption - Subtrai), "#####0.00")

        fgVendaItens.RemoveItem fgVendaItens.Row
    End If
End Sub

Private Sub coCategoria_LostFocus()
    If coCategoria.ListIndex <> 0 Then
        ckAdicional1.Enabled = True
        lbAdicional1.Caption = Mid(coCategoria.text, InStrRev(coCategoria.text, "$") + 1)
    Else
        ckAdicional1.Enabled = False
        lbAdicional1.Caption = None
        
    End If
    

End Sub




Private Sub Form_Load()

coCategoria.AddItem " "
coServico1.AddItem " "

'Mapeia todas as categorias
con.Open strConn
rs.Open "SELECT * FROM categorias  WHERE flag = 0", con, adOpenForwardOnly, adLockOptimistic
Do Until rs.EOF
    sEntrada = rs("id_categoria") & "-" & rs("nome") & "R$" & rs("adicional")
    coCategoria.AddItem sEntrada
    coCategoria.ListIndex = 0
    rs.MoveNext
Loop
rs.Close

'Mapeia todas os servicos
rs.Open "SELECT * FROM servicos WHERE flag = 0", con, adOpenForwardOnly, adLockOptimistic
Do Until rs.EOF
    sEntrada = rs("id_servico") & "-" & rs("nome") & "R$" & rs("preco")
    coServico1.AddItem sEntrada
    coServico1.ListIndex = 0
    rs.MoveNext
Loop
rs.Close

' configura largura das colunas
fgVendaItens.ColWidth(0) = 1000
fgVendaItens.ColWidth(1) = 2250
fgVendaItens.ColWidth(2) = 2200
fgVendaItens.ColWidth(3) = 2500
fgVendaItens.ColWidth(4) = 2500

' define a altura da linha 0
fgVendaItens.RowHeight(0) = 250

' Define o titulo das colunas fixas
fgVendaItens.Row = 0
fgVendaItens.Col = 0
fgVendaItens.text = "Serviço"
fgVendaItens.Col = 1
fgVendaItens.text = "Nome do Serviço"
fgVendaItens.Col = 2
fgVendaItens.text = "Adicional"
fgVendaItens.Col = 3
fgVendaItens.text = "Desconto"
fgVendaItens.Col = 4
fgVendaItens.text = "Valor total"

' Define o alinhamento das colunas fixas
fgVendaItens.FixedAlignment(0) = 2
fgVendaItens.FixedAlignment(1) = 2
fgVendaItens.FixedAlignment(2) = 2
fgVendaItens.FixedAlignment(3) = 2

' define o alinhamento das demais colunas
fgVendaItens.ColAlignment(0) = 0
fgVendaItens.ColAlignment(1) = 0
fgVendaItens.ColAlignment(2) = 1
fgVendaItens.ColAlignment(3) = 1

con.Close


End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmTelaCadastros.Enabled = True
End Sub

Private Sub tbDesconto_LostFocus()
    tbDesconto.text = Format(tbDesconto.text, "#####0.00")
End Sub

