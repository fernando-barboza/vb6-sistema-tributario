VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFuncoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim pstrNomeCampo As String

Public Function bolFuncGravaDados(vetCampos, strModoOperacao As String, strTabela As String, _
                                  adoConn As ADODB.Connection) As String
'=============================================================================================
'Função monta a query para manutenção na base de dados
'Se houver algum erro na execução da query, retorna a mensagem de erro tratada
'Parâmetros:
    'vetCampos =>       O vetor que contém os nomes dos campos na tabela e os valores do
    '                   formulário
    'strModoOperacao => (I)nclusão, (A)lteração, (E)xclusão
    'strTabela       => A tabela que recebera manutenção
    'AdoConn         => A conexão aberta com o banco de dados
'=============================================================================================


'******************************************************************************************
' Data: 06/03/2003
' Alteração: - Substituição dos comandos nativos do SQL Server pelas variáveis de comandos
'            nativos.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 07/03/2003
' Alteração: - Qdo o campo começar com "dtm" e o Banco de Dados corrente for Oracle não
'            devem ser incluídas as aspas.
' Responsável: Everton Bianchini
'******************************************************************************************
    
'Auxiliares para o loop no formulário
Dim i As Integer
Dim j As Integer
Dim k As Integer

'Auxliar para montagem da query
Dim strSql As String
Dim strAux As String

'Auxiliar para montagem da condição da query (WHERE)
Dim strCondicao As String
Dim strAspas    As String

'Auxiliares para recuperação do nome do campo onde o banco retornou erro
Dim intPos As Integer
Dim strCampo As String
Dim strValor As String

'Auxiliares para montagem da mensagem de erro (quando houver)
Dim vetX() As String
Dim vetChaves() As String

On Error GoTo err_bolFuncGravaDados

Select Case strModoOperacao
    Case "A" 'Alteração
        'Monta query
        strSql = ""
        strSql = strSql & " UPDATE " & strTabela & " SET "
        'Percorre o vetor
        For i = 0 To UBound(vetCampos, 2)
            'Se PKID sempre será a condição para alteração
            If UCase(vetCampos(0, i)) = "PKID" Then
                'By power strCondicao = " WHERE " & vetCampos(0, i) & " = '" & vetCampos(1, i) & "'"
                strCondicao = " WHERE " & vetCampos(0, i) & " = " & vetCampos(1, i)
            Else
                
                'Vamos tratar se o campo e do tipo string
'                If Left(vetCampos(0, i), 3) = "str" Or Left(vetCampos(0, i), 3) = "dtm" Then strAspas = "'" Else strAspas = ""
                If Left(vetCampos(0, i), 3) = "str" Or _
                    (Left(vetCampos(0, i), 3) = "dtm" And bytDBType = EDatabases.SQLServer) Then
                    strAspas = "'"
                Else
                    strAspas = ""
                End If
        
                'Monta (campo da tabela = valor)
'                If vetCampos(1, i) <> "NULL" Then
                If (vetCampos(1, i) <> "NULL") And _
                    ((Left(vetCampos(0, i), 3) <> "dtm" And bytDBType = EDatabases.Oracle) Or _
                     (bytDBType = EDatabases.SQLServer)) Then
                
                    'Verifica se há divisória entre o valor a ser atribuído a tabela
                    'Obs: Isso ocorre quando o objeto é DataCombo que possui Codigo e Descricao
                    If UCase(vetCampos(3, i)) = "DBC" Or UCase(vetCampos(3, i)) = "CMB" Or _
                       UCase(vetCampos(3, i)) = "CBO" Then
                        vetX = Split(vetCampos(1, i), "_")
                        'Se encontrou divisória
                        If UBound(vetX) > 0 Then
                            'Atribui o código ao banco
                            strSql = strSql & vetCampos(0, i) & " = " & strAspas & Replace(vetX(0), "'", Chr(207)) & strAspas & ","
                            'Atribui ao vetor a Descrição para mensagem caso ocorra erro.
                            vetCampos(1, i) = vetX(1)
                        Else
                            strSql = strSql & vetCampos(0, i) & " = " & strAspas & Replace(vetCampos(1, i), "'", Chr(207)) & strAspas & ","
                        End If
                    Else
                        strSql = strSql & vetCampos(0, i) & " = " & strAspas & Replace(vetCampos(1, i), "'", Chr(207)) & strAspas & ","
                    End If
                Else
                    'Atribuição de nulo.
                    strSql = strSql & vetCampos(0, i) & " = " & vetCampos(1, i) & ","
                End If
            End If
        Next i
        'Conclui com os campos fixos (data de atualizacao e código do usuário
'        strSql = strSql & " dtmDtAtualizacao = GETDATE(), "
        strSql = strSql & " dtmDtAtualizacao = " & strGETDATE & ", "
        strSql = strSql & "lngCodUsr = " & glngCodUsr
        strSql = strSql & strCondicao
    Case "E" 'Exclusão
        strSql = ""
        strSql = strSql & "DELETE FROM " & strTabela
        
        'Vamos tratar se o campo e do tipo string
        If Left(vetCampos(0, i), 3) = "str" Or Left(vetCampos(0, i), 3) = "dtm" Then strAspas = "'" Else strAspas = ""
        
        'For i = 0 To UBound(vetCampos) - 1
            'Se PKID sempre será a condição para exclusão
            'If UCase(vetCampos(0, i)) = "PKID" Then
                strCondicao = " WHERE " & vetCampos(0, i) & " = " & strAspas & vetCampos(1, i) & strAspas
             '   Exit For
            'End If
            
            
        'Next i
        strSql = strSql & strCondicao
    Case "I" 'Inlcuir
        'Monta query
        strSql = ""
        strSql = strSql & " INSERT INTO " & strTabela & " ("
        
        strCampo = ""
        strValor = ""
        'Percorre o vetor
        For i = 0 To UBound(vetCampos, 2)
            'Se PKID sempre será a condição para alteração
            If UCase(vetCampos(0, i)) = "PKID" Then
                strCondicao = ""
            Else
                strCampo = strCampo & vetCampos(0, i) & ","
                
                'Vamos tratar se o campo e do tipo string
'                If Left(vetCampos(0, i), 3) = "str" Or Left(vetCampos(0, i), 3) = "dtm" Then strAspas = "'" Else strAspas = ""
                If Left(vetCampos(0, i), 3) = "str" Or _
                    (Left(vetCampos(0, i), 3) = "dtm" And bytDBType = EDatabases.SQLServer) Then
                    strAspas = "'"
                Else
                    strAspas = ""
                End If
                
'                If vetCampos(1, i) <> "NULL" Then
                If (vetCampos(1, i) <> "NULL") And _
                    ((Left(vetCampos(0, i), 3) <> "dtm" And bytDBType = EDatabases.Oracle) Or _
                     (bytDBType = EDatabases.SQLServer)) Then
                    'Verifica se há divisória entre o valor a ser atribuído a tabela
                    'Obs: Isso ocorre quando o objeto é DataCombo que possui Codigo e Descricao
                    If UCase(vetCampos(3, i)) = "DBC" Or UCase(vetCampos(3, i)) = "CMB" Or _
                       UCase(vetCampos(3, i)) = "CBO" Then
                        vetX = Split(vetCampos(1, i), "_")
                        'Se encontrou divisória
                        If UBound(vetX) > 0 Then
                            'Atribui o código ao banco
                            strValor = strValor & strAspas & Replace(vetX(0), "'", "X") & strAspas & ","
                            'Atribui ao vetor a Descrição para mensagem caso ocorra erro.
                            vetCampos(1, i) = vetX(1)
                        Else
                            strValor = strValor & strAspas & Replace(vetCampos(1, i), "'", Chr(207)) & strAspas & ","
                        End If
                    Else
                        strValor = strValor & strAspas & Replace(vetCampos(1, i), "'", Chr(207)) & strAspas & ","
                    End If
                Else
                    strValor = strValor & vetCampos(1, i) & ","
                End If
            End If
        Next i
        'Conclui com os campos fixos (data de atualizacao e código do usuário
        strSql = strSql & strCampo & "dtmDtAtualizacao, lngcodusr)"
        strSql = strSql & " VALUES ("
'        strSql = strSql & strValor & " GETDATE(), "
        strSql = strSql & strValor & strGETDATE & ", "
        strSql = strSql & glngCodUsr & ")"
End Select
'Registra a movimentação do usuário
strAux = MontaRegistroAlteracao(vetCampos, strTabela, strModoOperacao)

'Executa a query montada acima (óbvio)
adoConn.Execute strSql

err_bolFuncGravaDados:

Select Case Err.Number
    Case 0
        GravaHistoricoOperacao strTabela, strModoOperacao, strAux
    Case -2147217900 'Campo duplicado
        If InStr(1, UCase(Err.Description), "CANNOT INSERT THE VALUE NULL") > 0 Then
            vetX() = Split(Err.Description, "'")
            strCampo = vetX(1)
            For i = 0 To UBound(vetCampos, 2)
            'Verifica se o campo da tabela é booleano (0/1)
                If UCase(Left(strCampo, 3)) = "BLN" Or UCase(Left(strCampo, 3)) = "BIT" Or UCase(Left(strCampo, 3)) = "BYT" Then
                    bolFuncGravaDados = "NOME } fra_" & strCampo
                    Exit For
                Else
                    If UCase(strCampo) = UCase(vetCampos(0, i)) Then
                        strNomeCampo = vetCampos(3, i) & strCampo
                        bolFuncGravaDados = RetornaMensagem(CStr(vetCampos(2, i)), CStr(vetCampos(1, i)), "CANNOT INSERT THE VALUE NULL")
                        Exit For
                    End If
                End If
            Next i
        ElseIf InStr(1, UCase(Err.Description), UCase("Cannot insert duplicate key row in object")) > 0 Then
            vetX() = Split(Err.Description, "'")
            
            strCampo = Replace(vetX(3), "IX_", "")
            
            vetX() = Split(strCampo, "_")
            
            If UBound(vetX) = 0 Then
                For i = 0 To UBound(vetCampos, 2)
                    If UCase(strCampo) = UCase(vetCampos(0, i)) Then
                        strNomeCampo = vetCampos(3, i) & strCampo
                        bolFuncGravaDados = RetornaMensagem(CStr(vetCampos(2, i)), CStr(vetCampos(1, i)), UCase("Cannot insert duplicate key row in object"))
                        Exit For
                    End If
                Next i
            Else
                For i = 0 To UBound(vetCampos, 2)
                    If UCase(vetX(UBound(vetX))) = UCase(vetCampos(0, i)) Then
                        strNomeCampo = vetCampos(3, i) & vetX(UBound(vetX))
                        For j = 0 To UBound(vetX)
                            For k = 0 To UBound(vetCampos, 2)
                                If UCase(vetX(j)) = UCase(vetCampos(0, k)) Then
                                    ReDim Preserve vetChaves(1, j)
                                    If UCase(vetCampos(3, k)) = "OPT" Then
                                        vetChaves(0, j) = vetCampos(2, k)
                                    Else
                                        strNomeCampo = vetCampos(3, k) & vetX(j)
                                        vetChaves(0, j) = vetCampos(2, k)
                                        vetChaves(1, j) = vetCampos(1, k)
                                    End If
                                End If
                            Next k
                        Next j
                        Exit For
                    End If
                Next i
                bolFuncGravaDados = RetornaMensagem("", "", UCase("Cannot insert duplicate key row in object"), vetChaves)
            End If
        ElseIf InStr(1, UCase(Err.Description), "DELETE STATEMENT CONFLICTED") > 0 Then
            vetX = Split(Err.Description, "'")
            vetX = Split(vetX(1), "_")
            bolFuncGravaDados = RetornaMensagem("", "", "DELETE STATEMENT CONFLICTED", vetX, adoConn)
        Else
            ExibeDetalheErro "", , vbCritical
        End If
    Case -2147217913
        If InStr(1, UCase(Err.Description), "SYNTAX ERROR CONVERTING THE VARCHAR VALUE") > 0 Then
            vetX = Split(Err.Description, "'")
                        
            'Localiza o valor inválido
            For i = 0 To UBound(vetCampos, 2)
                If UCase(vetCampos(1, i)) = UCase(vetX(1)) Then
                    strCampo = vetCampos(2, i)
                    strValor = vetCampos(1, i)
                    strNomeCampo = vetCampos(3, i) & vetCampos(0, i)
                    bolFuncGravaDados = RetornaMensagem(strCampo, strValor, "SYNTAX ERROR CONVERTING THE VARCHAR VALUE", vetX, adoConn)
                End If
            Next i
        Else
            If Err.Number <> 0 Then
                ExibeDetalheErro "", , vbCritical
            End If
        End If
    Case 3021
        ExibeMensagem "Não foi possível completar a operação." & vbCr & _
                      "Este registro foi excluído por outro usuário."
    Case -2147217873 And strModoOperacao = "E"
        ExibeMensagem gstrMsgExclusao(Err, "")
    Case (-2147217873 And strModoOperacao = "I") Or (-2147217873 And strModoOperacao = "A")
        ExibeMensagem "Não foi possível completar a operação." & Chr(13) & " Registro já se encontra cadastrado. "
    Case Else
        If Err.Number <> 0 Then
            ExibeDetalheErro "", , vbCritical
            bolFuncGravaDados = Err.Number & " = " & Err.Description
        End If
End Select
End Function

Private Function RetornaMensagem(strCampo As String, strValor As String, varErro, Optional vetC, _
                                 Optional adoConn As ADODB.Connection) As String
Dim strMsg As String
Dim i As Integer
Dim j As Integer
Dim STRNOME As String

strMsg = ""
Select Case varErro
    Case UCase("Cannot insert duplicate key row in object")
        If IsMissing(vetC) Then
            strMsg = "O campo " & Trim(strCampo) & " = " & Trim(strValor) & " já está cadastrado!"
        Else
            strMsg = ""
            strMsg = strMsg & "Os Campos "
            For i = 0 To UBound(vetC, 2)
                If STRNOME <> vetC(0, i) Then
                    If Trim(vetC(1, i)) = "" Then
                        strMsg = strMsg & vetC(0, i)
                    Else
                        strMsg = strMsg & Trim(vetC(0, i)) & " = " & Trim(vetC(1, i))
                    End If
                    If i + 1 <= UBound(vetC, 2) Then
                        strMsg = strMsg & " e "
                    End If
                End If
                STRNOME = vetC(0, i)
            Next i
            strMsg = strMsg & " já estão cadastrados!"
        End If
    Case "CANNOT INSERT THE VALUE NULL"
        strMsg = "O campo " & Trim(strCampo) & " não pode ser nulo!"
    Case "DELETE STATEMENT CONFLICTED"
        strMsg = RetornaNomeTabela(vetC(1), adoConn)
        strMsg = "Este registro não pode ser excluído porque ele " & Chr(10) & "está sendo utilizado em " & strMsg
    Case "SYNTAX ERROR CONVERTING THE VARCHAR VALUE"
        strMsg = strValor & " não é um valor válido para o campo " & strCampo
End Select



RetornaMensagem = strMsg
End Function

Public Property Get strNomeCampo() As String
strNomeCampo = pstrNomeCampo
End Property

Public Property Let strNomeCampo(ByVal vNovoValor As String)
pstrNomeCampo = vNovoValor
End Property

Private Function RetornaNomeTabela(strNomeTabela, adoConn As ADODB.Connection) As String
Dim strSql As String
Dim adoRec As ADODB.Recordset

strSql = ""
strSql = strSql & " SELECT strDescricao FROM " & gstrCatalogoTabela
strSql = strSql & " WHERE UPPER(strTabela) = '" & UCase(strNomeTabela) & "'"

Set adoRec = New ADODB.Recordset
adoRec.Open strSql, adoConn

If Not adoRec.EOF Then
    RetornaNomeTabela = UCase(adoRec!strDescricao)
End If
End Function


Public Function strTrataErro(strErrNumber As Long, strErrDescription As String, strTabela As String, adoConn As ADODB.Connection) As String
Dim i As Integer
Dim j As Integer
Dim k As Integer

'Auxiliares para recuperação do nome do campo onde o banco retornou erro
Dim intPos As Integer
Dim strCampo As String
Dim strValor As String

'Auxiliares para montagem da mensagem de erro (quando houver)
Dim vetX() As String
Dim vetChaves() As String


Select Case Err.Number
    Case -2147217900 '
        'Campo Nulo
        If InStr(1, UCase(Err.Description), "CANNOT INSERT THE VALUE NULL") > 0 Then
            vetX() = Split(Err.Description, "'")
            
            strCampo = vetX(1)
            strNomeCampo = strCampo
            strTrataErro = "" 'RetornaMensagem(CStr(vetCampos(2, i)), CStr(vetCampos(1, i)), "CANNOT INSERT THE VALUE NULL")
        'Campo Duplicado
        ElseIf InStr(1, UCase(Err.Description), UCase("Cannot insert duplicate key row in object")) > 0 Then
            vetX() = Split(Err.Description, "'")
            
            strCampo = Replace(vetX(3), "IX_", "")
            
            vetX() = Split(strCampo, "_")
            
            If UBound(vetX) = 0 Then
                
                strNomeCampo = vetX(1)
                strTrataErro = "" 'RetornaMensagem(CStr(vetCampos(2, i)), CStr(vetCampos(1, i)), UCase("Cannot insert duplicate key row in object"))
            Else
    
                For j = 0 To UBound(vetX)
                    ReDim Preserve vetChaves(1, j)
                    strNomeCampo = vetX(j)
                Next j
                strTrataErro = "" 'RetornaMensagem("", "", UCase("Cannot insert duplicate key row in object"), vetChaves)
            End If
        'Integridade Referencial Inclusão
        ElseIf InStr(1, UCase(Err.Description), "DELETE STATEMENT CONFLICTED") > 0 Then
            vetX = Split(Err.Description, "'")
            vetX = Split(vetX(1), "_")
            strTrataErro = "" 'RetornaMensagem("", "", "DELETE STATEMENT CONFLICTED", vetX, adoConn)
        Else
            ExibeDetalheErro ""
        End If
    Case -2147217913
        'Tipo de Dados Incorreto
        If InStr(1, UCase(Err.Description), "SYNTAX ERROR CONVERTING THE VARCHAR VALUE") > 0 Then
            vetX = Split(Err.Description, "'")
            
            'Localiza o valor inválido
            
            strCampo = UCase(vetX(1))
            strNomeCampo = strCampo
            strTrataErro = "" 'RetornaMensagem(strCampo, strValor, "SYNTAX ERROR CONVERTING THE VARCHAR VALUE", vetX, adoConn)
        Else
            If Err.Number <> 0 Then
                ExibeDetalheErro ""
            End If
        End If
    Case Else
        'Erro imprevisto
        If Err.Number <> 0 Then
            ExibeDetalheErro ""
        End If
End Select

End Function

Public Function MontaRegistroAlteracao(vetCampos, strTabela As String, strModoOperacao As String) As String
Dim strSql As String
Dim adoRec As ADODB.Recordset
Dim i As Integer
Dim lngPkid As Long
Dim objCampo As ADODB.Field
Dim vetAtual() As String
Dim strAux As String
Dim strCampo As String

On Error Resume Next

Set gobjBanco = New clsBanco

Select Case UCase(strModoOperacao)
    Case "A" 'Alteração
        For i = 0 To UBound(vetCampos, 2)
            If UCase(vetCampos(0, i)) = "PKID" Then
                lngPkid = vetCampos(1, i)
                Exit For
            End If
        Next i
        
        strSql = ""
        strSql = strSql & "SELECT * FROM " & strTabela & " WHERE PKID = " & lngPkid
        
        strAux = ""
        'Alterada no 2º Retorno da ficha orc1584
        If gobjBanco.CriaADO(strSql, 10, adoRec) Then
            For Each objCampo In adoRec.Fields
                If UCase(objCampo.Name) <> "DTMDTATUALIZACAO" And UCase(objCampo.Name) <> "LNGCODUSR" And _
                   UCase(objCampo.Name) <> "PKID" Then
                    For i = 0 To UBound(vetCampos, 2)
                        If UCase(objCampo.Name) = UCase(vetCampos(0, i)) Then
                            If Not IsNull(objCampo.Value) Then
                                'Campos Option Button
                                If InStr(UCase(objCampo.Value), "TRUE") > 0 Then
                                    strCampo = "1"
                                ElseIf InStr(UCase(objCampo.Value), "FALSE") > 0 Then
                                    strCampo = "0"
                                Else
                                    'Outros campos
                                    strCampo = objCampo.Value
                                End If
                            Else
                                'Se for Nullo atribuo NULL para
                                'igualar com a rotina SalvarGeral
                                strCampo = "NULL"
                            End If
                            If strCampo <> vetCampos(1, i) Then
                                strAux = strAux & vetCampos(2, i) & " => DE: " & Trim(objCampo.Value) & " Para: " & Replace(Trim(vetCampos(1, i)), "'", Chr(207)) & Chr(13)
                            Else
                                Exit For
                            End If
                            
                        End If
                    Next i
                End If
            Next
        End If
        
    Case "E" 'Exclusão
        lngPkid = vetCampos(1, 0)
        
        strSql = ""
        strSql = strSql & "SELECT * FROM " & strTabela & " WHERE PKID = " & lngPkid
        
        strAux = ""
        If gobjBanco.CriaADO(strSql, 10, adoRec) Then
            For Each objCampo In adoRec.Fields
                If InStr(UCase(objCampo.Name), "DESCRICAO") Or InStr(UCase(objCampo.Name), "NOME") Then
                    strAux = "Descrição = " & objCampo.Value
                    Exit For
                End If
            Next
            If strAux = "" Then
                strAux = "ID = " & lngPkid
            End If
        End If
    Case "I" 'Inclusão
        strAux = ""
        For i = 1 To UBound(vetCampos, 2) - 1
            strAux = strAux & vetCampos(2, i) & " = " & Replace(Trim(vetCampos(1, i)), "'", Chr(207)) & Chr(13)
        Next i
End Select

MontaRegistroAlteracao = strAux
End Function


Public Sub GravaHistoricoOperacao(strTabela As String, strModoOperacao As String, strAux As String)
    Dim strSql As String
    Dim adoRec As ADODB.Recordset
    Dim lngPKIdTabela As Long
    
Retorno:

    strSql = ""
    strSql = strSql & " SELECT PKID FROM " & gstrCatalogoTabela & " WHERE strTabela = '" & strTabela & "'"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 10, adoRec) Then
        If Not (adoRec.BOF And adoRec.EOF) Then
            lngPKIdTabela = adoRec!Pkid
        Else
            strSql = ""
            strSql = strSql & "INSERT INTO " & gstrCatalogoTabela
            strSql = strSql & "(strtabela, strdescricao, dtmdtatualizacao, lngcodusr)"
            strSql = strSql & "Values "
            strSql = strSql & "('" & strTabela & "', '" & strTabela & "', " & gstrConvDtParaSql(gstrDataDoSistema(True)) & ", " & glngCodUsr & ")"
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSql) Then
                GoTo Retorno:
            Else
                Exit Sub
            End If
        End If
    End If
    
    strSql = ""
    strSql = strSql & " INSERT INTO " & gstrHistoricoOperacao
    strSql = strSql & " (intUsuario, intCatalogoTabela, bytModulo, strOperacao, dtmData, strValor) VALUES ("
    strSql = strSql & glngCodUsr
    strSql = strSql & ", " & lngPKIdTabela
    strSql = strSql & ", " & bytRetornaCodigoModulo(App.ProductName)
    strSql = strSql & ", '" & strModoOperacao & "'"
    strSql = strSql & ", " & gstrConvDtParaSql(gstrDataDoSistema(True))
    strSql = strSql & ", '" & strAux & "'"
    strSql = strSql & ")"
    
    Set gobjBanco = New clsBanco
    gobjBanco.Execute strSql
    
End Sub

