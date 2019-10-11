Attribute VB_Name = "mdRelatoriosCSV"
Public Sub RelatorioVendas(ByVal dtPiso As Date, dtTeto As Date)
    Dim rSQL, freport As String
    Dim timestamp As Date
    Dim localReportPath As String
    
    timestamp = Now
    
    localReportPath = sReportPath & "vendas.csv"
    
    ' Get a free file number
    freport = FreeFile
    ' Create Report
    Open localReportPath For Output As freport
    
    
    rSQL = "SELECT * FROM ordens WHERE hora BETWEEN #" & dtPiso & "# AND #" & dtTeto & "#"
    con.Open strConn
    rs.Open rSQL, con, adOpenForwardOnly, adLockOptimistic
    
    sLinha = "id_ordem;placa;categoria;valor_total;usuario;hora"
    Print #freport, sLinha
    
    Do Until rs.EOF
        sLinha = rs("ordens.id_ordem") & ";" & rs("ordens.placa") & ";" & rs("categorias.id_categoria") & ";" & rs("ordens.valor_total") & ";" & rs("funcionarios.nome") & ";" & rs("hora")
        Print #freport, sLinha
        rs.MoveNext
    Loop
    rs.Close
    Close freport
    Call Log("Gerando relatórios de vendas...")
    MsgBox "Relatório gerado com sucesso!"
End Sub
Public Sub RelatorioFuncionarios()
    Dim rSQL, freport As String
    Dim timestamp As Date
    Dim localReportPath As String
    
    timestamp = Now
    
    localReportPath = sReportPath & "funcionarios.csv"
    
    ' Get a free file number
    freport = FreeFile
    ' Create Report
    Open localReportPath For Output As freport
    
    
    rSQL = "SELECT * FROM funcionarios"
    con.Open strConn
    rs.Open rSQL, con, adOpenForwardOnly, adLockOptimistic
    
    sLinha = "matricula;nome;cargo"
    Print #freport, sLinha
    
    Do Until rs.EOF
        sLinha = rs("matricula") & ";" & rs("nome") & ";" & rs("cargo")
        Print #freport, sLinha
        rs.MoveNext
    Loop
    rs.Close
    Close freport
    Call Log("Gerando relatórios de funcionarios...")
    MsgBox "Relatório gerado com sucesso!"
End Sub
