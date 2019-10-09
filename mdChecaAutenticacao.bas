Attribute VB_Name = "mdChecaAutenticacao"
Public Function ChecaAutenticacao(ByVal strMatricula, strSenha As String) As Boolean
    Dim CSQL As String
    If IsNumeric(strMatricula) Then
        CSQL = "SELECT * FROM funcionarios LEFT JOIN cargos ON funcionarios.cargo = cargos.id_cargo WHERE funcionarios.flag = 0 and matricula = " & strMatricula
    
        Call Log("Tentativa de login")
        
        'Converte a senha para Hash
        strSenha = UCase(MD5.DigestStrToHexStr(strSenha))
        
        'Pesquisa no banco para verificar se tem o usuário
        
        con.Open strConn
        rs.Open CSQL, con, adOpenForwardOnly, adLockReadOnly
        Do Until rs.EOF
            If strSenha = rs("senha") And rs("funcionarios.flag") = 0 Then
                frmTelaCadastros.lbSessao = rs("matricula") & " - " & rs("funcionarios.nome")
                If rs("acesso_admin") = 1 Then frmTelaCadastros.lbAdmin = "Admin"
                If rs("acesso_rh") = 1 Then frmTelaCadastros.lbRh = "RH"
                ChecaAutenticacao = True
                rs.Close
                con.Close
                Exit Function
            End If
        Loop
    Else
        Exit Function
    End If
End Function
Public Sub RealizaLogoff()

        boolAutenticacao = False
        frmTelaCadastros.lbSessao.Caption = "Nenhum Usuário"
        frmTelaCadastros.lbAdmin.Caption = ""
        frmTelaCadastros.lbRh.Caption = ""

End Sub
