VERSION 5.00
Begin VB.Form frmImportarDados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importação de Dados"
   ClientHeight    =   1785
   ClientLeft      =   4290
   ClientTop       =   4380
   ClientWidth     =   4995
   Icon            =   "frmImportarDados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImportacao 
      Caption         =   "Iniciar Importação"
      Height          =   450
      Left            =   2910
      TabIndex        =   5
      Top             =   1200
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      Caption         =   " Parâmetros de Importação "
      Height          =   900
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   4695
      Begin VB.TextBox txt_Inicial 
         Height          =   285
         Left            =   1110
         MaxLength       =   10
         TabIndex        =   2
         Top             =   330
         Width           =   1065
      End
      Begin VB.TextBox txt_Final 
         Height          =   285
         Left            =   3390
         MaxLength       =   10
         TabIndex        =   4
         Top             =   345
         Width           =   1065
      End
      Begin VB.Label lbl_Inicial 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   225
         TabIndex        =   1
         Top             =   390
         Width           =   795
      End
      Begin VB.Label lbl_Final 
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
         Height          =   195
         Left            =   2565
         TabIndex        =   3
         Top             =   390
         Width           =   720
      End
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   1305
      Width           =   45
   End
End
Attribute VB_Name = "frmImportarDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aCriticas()             As Variant

Private Sub cmdImportacao_Click()
Dim cncADO                  As ADODB.Connection
Dim adoBanco                As ADODB.Recordset
Dim adoReceita              As ADODB.Recordset
Dim adoAux                  As ADODB.Recordset
Dim adoResultado            As ADODB.Recordset
Dim strSql                  As String
Dim bExtra                  As Boolean
Dim strDataBase             As String
Dim strCodigoOrcamentario   As String
Dim strConeccao             As String
Dim aBanco                  As New XArrayDB
Dim aReceita                As New XArrayDB
Dim aArrecadacao            As New XArrayDB
Dim aImportacao()           As String
Dim aTipoMovimento()       As Variant
Dim aContaExtra()          As Variant
Dim aValor()               As Variant
Dim ncount                  As Integer
Dim dtmDataEmOperacao       As Date
Dim dblPorcentagem          As Double
Dim intTotalBanco           As Integer
Dim intTotalReceita         As Integer
Dim dblValorTotalBanco      As Double
Dim dblValorTotalReceita    As Double
Dim dblValorPorReceita      As Double
Dim dblValorPorBanco        As Double
Dim dblValorDiferenca       As Double
Dim intForBanco             As Integer
Dim intForReceita           As Integer
Dim intForImportacao        As Integer
Dim intForColunas           As Integer
Dim varAux                  As Variant
Dim blnErroServidor         As Boolean

On Error GoTo ErroAbreBancoDados
    
    If blnDadosOK Then
                 
        lblStatus.Caption = "Conectando ..."
        
        'Abrindo coneccao
        strDataBase = "C:\temp\Arrecadacao.mdb"
        
        Set cncADO = New ADODB.Connection
    
'        strConeccao = "DRIVER={Microsoft Access Driver (*.mdb)};" & _
'                      "DBQ=" & strDataBase & ";" & _
'                      "DefaultDir=" & "C:\Temp\" & ";" & "UID=admin;PWD=;"
                      
         strConeccao = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & strDataBase & ""
                   
                      
              
        cncADO.ConnectionTimeout = 10
        cncADO.CommandTimeout = 10
        
        cncADO.ConnectionString = strConeccao
        cncADO.Open
            
        lblStatus.Caption = "Consultando banco de origem ..."
        
        
        'Capturando quantidade de dados da tabela de banco
        Set adoAux = New ADODB.Recordset
        
        adoAux.CursorLocation = adUseClient
        
        adoAux.Open "SELECT A.cdrcbanco " & _
                    "FROM vrcbco AS A " & _
                    "WHERE A.dtfech BETWEEN #" & Format(txt_Inicial.Text, "mm/dd/yyyy hh:mm:ss") & "# AND #" & Format(txt_Final.Text, "mm/dd/yyyy hh:mm:ss") & "# " & _
                    "GROUP BY A.cdrcbanco ", cncADO, adOpenKeyset, adLockOptimistic
        If adoAux.EOF Then
            
            ExibeMensagem "Não foram encontrados registros neste intervalo de datas."
            
            lblStatus.Caption = ""
                   
            cncADO.Close
            Set cncADO = Nothing
            
            Exit Sub
            
        End If
        
        adoAux.Close: Set adoAux = Nothing
        
        'Capturando dados da tabela de banco
        Set adoBanco = New ADODB.Recordset
        
        adoBanco.CursorLocation = adUseClient
        
        adoBanco.Open "SELECT SUM(A.vlCred) AS vlcred, A.cdrcbanco, A.dtfech, SUM(A.vlcred) / (SELECT SUM(B.vlCred) FROM vrcbco AS B WHERE B.dtfech =  A.dtfech GROUP BY B.dtfech) AS Porcentagem " & _
                      "FROM vrcbco AS A " & _
                      "WHERE A.dtfech BETWEEN #" & Format(txt_Inicial.Text, "mm/dd/yyyy hh:mm:ss") & "# AND #" & Format(txt_Final.Text, "mm/dd/yyyy hh:mm:ss") & "# " & _
                      "GROUP BY A.cdrcbanco, A.dtfech ORDER BY A.dtfech, A.cdrcbanco", cncADO, adOpenKeyset, adLockOptimistic

        
        'Capturando dados da tabela de receita
        Set adoReceita = New ADODB.Recordset
        
        adoReceita.CursorLocation = adUseClient

        adoReceita.Open "SELECT SUM(A.vldistr) AS vldistr, A.cdrecred, A.cdrubrec, A.dtfech " & _
                        "FROM disrec AS A " & _
                        "WHERE A.dtfech BETWEEN #" & Format(txt_Inicial.Text, "mm/dd/yyyy hh:mm:ss") & "# AND #" & Format(txt_Final.Text, "mm/dd/yyyy hh:mm:ss") & "# " & _
                        "GROUP BY A.cdrecred, A.dtfech, A.cdrubrec ORDER BY  A.dtfech, A.cdrubrec DESC, A.cdrecred ", cncADO, adOpenKeyset, adLockOptimistic
        
        lblStatus.Caption = "Preparando dados para importação ..."
       
ProximaData:

        'Vamos preparar o array de receitas
        Set aArrecadacao = New XArrayDB
        aArrecadacao.Clear
        
        aArrecadacao.ReDim 0, adoReceita.RecordCount, 0, adoBanco.RecordCount + 1
        
        'Vamos preparar o array de bancos
        Set aBanco = New XArrayDB
        aBanco.Clear
            
        aBanco.ReDim 0, adoBanco.RecordCount - 1, 0, 4

        'Vamos preparar o array de receitas
        Set aReceita = New XArrayDB
        aReceita.Clear
            
        aReceita.ReDim 0, adoReceita.RecordCount - 1, 0, 3
        
    
        'Preenchendo o array de bancos
        With adoBanco
        
        ReDim aCriticas(1)
        
        dtmDataEmOperacao = adoBanco!dtfech
        
        dblValorTotalBanco = 0
        intTotalBanco = 0
        ncount = 0
        
        Do While Not .EOF
                
            If !dtfech <> dtmDataEmOperacao Then Exit Do
            
            varAux = !cdrcbanco
            aBanco(.AbsolutePosition - 1, 0) = varAux
            
            varAux = !dtfech
            aBanco(.AbsolutePosition - 1, 1) = varAux
            
            varAux = !vlcred
            aBanco(.AbsolutePosition - 1, 2) = varAux
            
            dblValorTotalBanco = dblValorTotalBanco + !vlcred
            
            varAux = !porcentagem
            aBanco(.AbsolutePosition - 1, 3) = varAux
            
            dblPorcentagem = dblPorcentagem + !porcentagem
            
            intTotalBanco = intTotalBanco + 1
            
            .MoveNext
        Loop
        
        End With
        
        'Preenchendo o array de receitas
        With adoReceita
        
        dblValorTotalReceita = 0
        intTotalReceita = 0
        
        Do While Not .EOF
        
            If !dtfech <> dtmDataEmOperacao Then Exit Do
            
            varAux = !cdrecred
            aReceita(.AbsolutePosition - 1, 0) = varAux
            
            varAux = !dtfech
            aReceita(.AbsolutePosition - 1, 1) = varAux
            
            varAux = !vldistr
            aReceita(.AbsolutePosition - 1, 2) = varAux
            
            varAux = !cdrubrec
            aReceita(.AbsolutePosition - 1, 3) = varAux
            
            dblValorTotalReceita = dblValorTotalReceita + !vldistr
            
            intTotalReceita = intTotalReceita + 1
            
            .MoveNext
        Loop
        
        End With
        
        'Vamos redimensionar o array de importacao preservando os dias anteriores
        ReDim Preserve aImportacao(intTotalReceita * intTotalBanco, 4)
        
        'Vamos verificar se o total de bancos e receitas esta batendo
        If CStr(dblValorTotalReceita) <> CStr(dblValorTotalBanco) Then
            ExibeMensagem "O valor total dos Bancos não coincide com o valor total das Receitas no dia " & dtmDataEmOperacao & "."
            ExibeDetalheErro Err.Description
            blnErroServidor = True
            lblStatus.Caption = ""
            cncADO.Close
            Set cncADO = Nothing
            Exit Sub
            
            'Vamos para a proxima data
            If Not adoBanco.EOF Then
                GoTo ProximaData
            Else
                GoTo Importacao
            End If
            
        End If
        
        'Vamos preparar os dados para importacao
        For intForReceita = 0 To intTotalReceita - 1
            
            dblValorPorReceita = 0
            
            For intForColunas = 0 To intTotalBanco
                    
                'Coluna com codigo da Receita
                If intForColunas = 0 Then
                    varAux = aReceita(intForReceita, 0)
                    aArrecadacao(intForReceita, intForColunas) = varAux
                'Colunas por Bancos
                Else
                    
                    varAux = (Round(aReceita(intForReceita, 2) * aBanco(intForColunas - 1, 3), 2))
                    
                    aArrecadacao(intForReceita, intForColunas) = varAux
                        
                    dblValorPorReceita = dblValorPorReceita + varAux
                    
                End If
                    
            Next
            
            'Vamos verificar se a soma das porcentagens das receitas e igual a soma por banco
            If CStr(dblValorPorReceita) <> CStr(aReceita(intForReceita, 2)) Then
                dblValorDiferenca = CCur(aReceita(intForReceita, 2)) - CCur(dblValorPorReceita)
                'Vamos retirar a diferenca da receita atual
                Set aArrecadacao = AcertaDiferenca(aArrecadacao, intForReceita, dblValorDiferenca, True)
            End If
            
        Next
        
        'Vamos verificar se ha diferenca de banco
        For intForColunas = 1 To intTotalBanco
            
            dblValorPorBanco = 0
            
            For intForReceita = 0 To intTotalReceita - 1
                
                dblValorPorBanco = dblValorPorBanco + Round(aArrecadacao(intForReceita, intForColunas), 2)
                    
            Next
            
            'Vamos verificar se a soma das porcentagens dos bancos e igual a soma por receita
            If CStr(dblValorPorBanco) <> CStr(aBanco(intForColunas - 1, 2)) Then
                dblValorDiferenca = CCur(aBanco(intForColunas - 1, 2)) - CCur(dblValorPorBanco)
                'Vamos retirar a diferenca do banco atual
                Set aArrecadacao = AcertaDiferenca(aArrecadacao, intForColunas, dblValorDiferenca, False)
            End If
            
        Next
        
        'Vamos carregar os dados para um array de importacao, ja com dados preparados
        For intForBanco = 1 To intTotalBanco
        
            'For intForReceita = 0 To aArrecadacao.UpperBound(1) - 1
            For intForReceita = 0 To intTotalReceita - 1
            
                'Data em operacao
                varAux = dtmDataEmOperacao
                aImportacao(intForImportacao, 0) = varAux
                
                'Banco da porcentagem em consulta
                varAux = aBanco(intForBanco - 1, 0)
                aImportacao(intForImportacao, 1) = varAux
                
                'Receita em consulta
                varAux = aArrecadacao(intForReceita, 0)
                aImportacao(intForImportacao, 2) = varAux

                'Valor da porcentagem do banco em consulta
                varAux = aArrecadacao(intForReceita, intForBanco)
                aImportacao(intForImportacao, 3) = varAux
                
                'Rubrica da Receita em consulta
                varAux = aReceita(intForReceita, 3)
                aImportacao(intForImportacao, 4) = Space$(0) & varAux
                
                intForImportacao = intForImportacao + 1
                
            Next
            
        Next
        
        'Vamos para a proxima data
        If Not adoBanco.EOF Then
            GoTo ProximaData
        End If
        
Importacao:

        Set adoBanco = Nothing
        
        Set adoReceita = Nothing
        
        'Fechando coneccao
        cncADO.Close
        
        Set cncADO = Nothing
        
        '******* FAZER PARTE DE IMPORTACAO *********
        
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        
        Dim lngBanco              As Long
        Dim blnOrcamentario       As Boolean
        Dim strGrupoConta         As String
        Dim lngEmpenho            As Long
        Dim lngPkidUltArrecadacao As Long
        Dim lngUltNumArrecadacao  As String
        Dim dtmDatabyevento       As String
        Dim blnPrimeiraVez        As Boolean
        Dim intFor                As Integer
        
        blnPrimeiraVez = True
        dtmDataEmOperacao = 0
        
        dblValorTotalReceita = 0
        
        For intFor = 0 To UBound(aImportacao) - 1
         
            If dtmDataEmOperacao <> aImportacao(intFor, 0) Then
            
                dtmDataEmOperacao = aImportacao(intFor, 0)
                
                lblStatus.Caption = "Convertendo dia " & dtmDataEmOperacao
                 
                'Vamos apagar os registros da mesma data
                gobjBanco.Execute ("DELETE FROM " & gstrContaArrecadacaoReceita & " WHERE intArrecadacao IN (SELECT Pkid FROM " & gstrArrecadacaoReceita & " WHERE dtmData = " & gstrConvDtParaSql(dtmDataEmOperacao) & " AND bytImportacao = 1)")
                gobjBanco.Execute ("DELETE FROM " & gstrArrecadacaoReceita & " WHERE dtmData = " & gstrConvDtParaSql(dtmDataEmOperacao) & " AND bytImportacao = 1")
                 
                gobjBanco.Execute ("DELETE FROM " & gstrLancamentoContabil & "   WHERE Intprocesso in( select " & gstrProcessoPagamento & ".pkid From " & gstrProcessoPagamento & " Where " & gstrProcessoPagamento & ".intorigem = 6 and " & gstrProcessoPagamento & ".dtmData =  " & gstrConvDtParaSql(dtmDataEmOperacao) & ")")
                gobjBanco.Execute ("DELETE FROM " & gstrProcessoPagamento & " WHERE intorigem = 6 and  dtmData = " & gstrConvDtParaSql(dtmDataEmOperacao) & "")
            
            End If
            'Vamos inserir na tabela de Arrecadacao de Receita
            If lngBanco <> aImportacao(intFor, 1) Then
                
NovoRegistroDeBanco:
                 ' Vamos chamar a rotina para gerar os movimentos de evento
                If blnPrimeiraVez = False Then
                    'chamada aqui
                     'AGUARDANDO VERIFICACAO DE GRUPOS QUE NAO POSSUEM EVENTOS
                     aTipoMovimento(1) = 1
                     aValor(1) = dblValorTotalReceita
                     
                     If Not GeraMovimentosByEvento(lngEmpenho, dtmDatabyevento, Str(dblValorTotalReceita), "", lngUltNumArrecadacao, "6", aContaExtra, aTipoMovimento, IIf(UBound(aContaExtra) > 1, True, False), aValor, True) Then
                         ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
                         gobjBanco.ExecutaRollbackTrans
                         Exit Sub
                     End If
                    
                     lngUltNumArrecadacao = 0
                     dblValorTotalReceita = 0
                End If
                
                blnPrimeiraVez = False
                blnOrcamentario = False
                
                'Capturando Pkid da tabela de banco
                Set adoAux = New ADODB.Recordset
                
                adoAux.Open "SELECT Pkid FROM " & gstrPlanoConta & " WHERE intContaBancaria IN (SELECT Pkid FROM " & gstrContaBancaria & " WHERE intNumeroConta = " & aImportacao(intFor, 1) & ")", gcncADOMain, adOpenKeyset, adLockOptimistic
                ReDim aContaExtra(1)
                ReDim aValor(1)
                ReDim aTipoMovimento(1)
                aContaExtra(1) = adoAux!Pkid
                ncount = 1
                
                strSql = "INSERT INTO " & gstrArrecadacaoReceita & " ("
                strSql = strSql & "intNumero, intExercicio, dtmData, intContaContabil, bytImportacao, "
                strSql = strSql & "dtmDtAtualizacao, lngCodUsr) "
                strSql = strSql & "SELECT " & gstrISNULL("MAX(intNumero) + 1", "0" + 1) & " , " & Year(aImportacao(intFor, 0)) & ", "
                strSql = strSql & gstrConvDtParaSql(aImportacao(intFor, 0)) & ", " & adoAux!Pkid & ", 1, "
                strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSql = strSql & glngCodUsr & " FROM " & gstrArrecadacaoReceita
               
                If Not gobjBanco.Execute(strSql) Then
                    ExibeMensagem "Ocorreu um erro ao importar dados para Arrecadação de Receita. A operação foi cancelada."
                    gobjBanco.ExecutaRollbackTrans
                    Exit Sub
                End If
                
                adoAux.Close: Set adoAux = New ADODB.Recordset
                
                adoAux.Open "SELECT " & gstrISNULL("MAX(Pkid)", "0") & " Maximo, " & gstrISNULL("MAX(intNumero)", "0") & " MaxNumero  FROM " & gstrArrecadacaoReceita, gcncADOMain, adOpenKeyset, adLockOptimistic
                lngPkidUltArrecadacao = adoAux!Maximo
                lngUltNumArrecadacao = adoAux!MaxNumero
                
                lngBanco = aImportacao(intFor, 1)
                
                adoAux.Close: Set adoAux = New ADODB.Recordset
                
            End If
            
            'Vamos verificar se e conta Orcamentaria ou Extra
            If aImportacao(intFor, 2) < 500 Then
                
                If RetornaReceitaDeReferencia(Trim(aImportacao(intFor, 4))) = "-1" Then
                    ExibeMensagem "A Rubrica de origem " & Trim(aImportacao(intFor, 4)) & " não está cruzada no sistema e o programa de importação não pode continuar." & vbNewLine & "Nenhum registro foi importado."
                    gobjBanco.ExecutaRollbackTrans
                    lblStatus.Caption = ""
                    Exit Sub
                End If
                
                adoAux.Open "SELECT Pkid, strCodigoOrcamentario FROM " & gstrCodigoOrcamentario & " WHERE " & strSUBSTRING & "(strCodigoOrcamentario,1," & Len(Trim(aImportacao(intFor, 4))) & ") = " & RetornaReceitaDeReferencia(Trim(aImportacao(intFor, 4))) & " AND intExercicio = " & Year(aImportacao(intFor, 0)), gcncADOMain, adOpenKeyset, adLockOptimistic
                'adoAux.Open "SELECT Pkid, strCodigoOrcamentario FROM " & gstrCodigoOrcamentario & " WHERE " & strSUBSTRING & "(strCodigoOrcamentario,1," & Len(Trim(aImportacao(intFor, 4))) & ") = '11120201' AND intExercicio = " & Year(aImportacao(intFor, 0)), gcncADOMain, adOpenKeyset, adLockOptimistic
                
                strCodigoOrcamentario = IIf(adoAux.EOF And adoAux.BOF = True, "", adoAux!strCodigoOrcamentario)
            
                If strCodigoOrcamentario = "" Then
                    ExibeMensagem "O Código Orçamentário de Destino " & RetornaReceitaDeReferencia(Trim(aImportacao(intFor, 4))) & " não foi encontrado e a importação não pode continuar." & vbNewLine & "Nenhum registro foi importado."
                    gobjBanco.ExecutaRollbackTrans
                    lblStatus.Caption = ""
                    Exit Sub
                End If
                
                blnOrcamentario = True
                
                'Vamos gravar o Evento Contabil
                If lngPkidUltArrecadacao > 0 Then
                    
                    Set adoBanco = New ADODB.Recordset
                    adoBanco.Open "SELECT EC.intEvento FROM " & gstrEventoContaContabilCredito & " EC, " & gstrPlanoConta & " PC WHERE EC.intEvento IN ( SELECT pkid FROM " & gstrEvento & " WHERE inttipoevento = 1) AND PC.Pkid = EC.intContaContabil AND SUBSTR(PC.strContaContabil,1,3) = '" & gstrDigitoReceita & Mid(strCodigoOrcamentario, 1, 2) & "'", gcncADOMain, adOpenKeyset, adLockOptimistic
                    If Not adoBanco.EOF Then
                        lngEmpenho = adoBanco!INTEVENTO
                    Else
                        lngEmpenho = 0
                    End If
                    adoBanco.Close: Set adoBanco = Nothing
                    
                    strSql = " UPDATE " & gstrArrecadacaoReceita & " SET intEvento = (" & lngEmpenho & ") WHERE Pkid = " & lngPkidUltArrecadacao

                    If Not gobjBanco.Execute(strSql) Then
                        ExibeMensagem "Ocorreu ao gravar o Evento para Arrecadação de Receita. A operação foi cancelada."
                        gobjBanco.ExecutaRollbackTrans
                        Exit Sub
                    End If
                    
                    lngPkidUltArrecadacao = 0
                                        
                End If
                
                'Vamos verificar se o Grupo da conta mudou. Caso tenha mudado vamos criar um novo registro
                If Len(strGrupoConta) > 0 And strGrupoConta <> Mid(strCodigoOrcamentario, 1, 2) Then
                    
                    strGrupoConta = Mid(strCodigoOrcamentario, 1, 2)
                    
'                    'AGUARDANDO VERIFICACAO DE GRUPOS QUE NAO POSSUEM EVENTOS
'                     aryTpMov(0) = 1
'                     aryValor(0) = dblValorTotalReceita
'                     If Not GeraMovimentosByEvento(lngEmpenho, aImportacao(intFor, 0), Str(dblValorTotalReceita), "", lngUltNumArrecadacao, "6", aryPlanoConta, aryTpMov, False, aryValor) Then
'                         ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
'                         gobjBanco.ExecutaRollbackTrans
'                         Exit Sub
'                     End If
'
'                    lngUltNumArrecadacao = 0
                    
                    GoTo NovoRegistroDeBanco
                    
                End If
                
                strGrupoConta = Mid(strCodigoOrcamentario, 1, 2)
                
            Else
                adoAux.Open "SELECT Pkid FROM " & gstrPlanoConta & " WHERE intExtraMaua = " & aImportacao(intFor, 2), gcncADOMain, adOpenKeyset, adLockOptimistic
                'Caso este banco ja tenha receita Orcamentaria, vamos criar um novo registro com as Extras
                If blnOrcamentario Then
                    
'                    If bExtra = True Then
'                        If Not GeraMovimentosByEvento(lngEmpenho, aImportacao(intFor, 0), Str(dblValorTotalReceita), "", lngUltNumArrecadacao, "6", aContaExtra, aTipoMovimento, True, aValor) Then
'                            ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
'                            gobjBanco.ExecutaRollbackTrans
'                            Exit Sub
'                        End If
'                        ncount = 1
'                    End If
                    GoTo NovoRegistroDeBanco
                    bExtra = False
                    
                Else
'                  carrega arrays
                  ncount = ncount + 1
                  bExtra = True
                  ReDim Preserve aContaExtra(ncount)
                  ReDim Preserve aTipoMovimento(ncount)
                  ReDim Preserve aValor(ncount)
                  aContaExtra(ncount) = adoAux!Pkid
                  aTipoMovimento(ncount) = 0
                  aValor(ncount) = aImportacao(intFor, 3)
                End If
            End If
            
            'Vamos inserir na tabela de Contas Arrecadacao de Receita
            strSql = "INSERT INTO " & gstrContaArrecadacaoReceita & " ("
            strSql = strSql & "intArrecadacao, intConta, dblValorOrcamentario, bytCancelado, dtmDataCancelamento, bytTipo, strTempRubRec, "
            strSql = strSql & "dtmDtAtualizacao, lngCodUsr) "
            strSql = strSql & "SELECT MAX(Pkid), " & adoAux!Pkid & ", "
            strSql = strSql & gstrConvVrParaSql(Abs(aImportacao(intFor, 3))) & ", " & IIf(aImportacao(intFor, 3) < 0, 1, 0) & ", " & IIf(aImportacao(intFor, 3) < 0, gstrConvDtParaSql(aImportacao(intFor, 0)), "NULL") & ", " & IIf(aImportacao(intFor, 2) < 500, 0, 1) & ", '" & Trim(aImportacao(intFor, 4)) & "', "
            strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema) & ", "
            strSql = strSql & glngCodUsr & " FROM " & gstrArrecadacaoReceita
           
            If Not gobjBanco.Execute(strSql) Then
                ExibeMensagem "Ocorreu um erro ao importar dados para Contas Arrecadação de Receita. A operação foi cancelada."
                gobjBanco.ExecutaRollbackTrans
                Exit Sub
            End If
            
            dblValorTotalReceita = dblValorTotalReceita + aImportacao(intFor, 3)
            
            adoAux.Close: Set adoAux = New ADODB.Recordset
            'Joga data do movimento do dia para a variavel
            dtmDatabyevento = aImportacao(intFor, 0)
        Next
         
        'Vamos chamar a rotina de geracao de eventos contabeis novamente para gerar o ultimo registro
        ' Vamos chamar a rotina para gerar os movimentos de evento
        If blnPrimeiraVez = False Then
            'chamada aqui
            'AGUARDANDO VERIFICACAO DE GRUPOS QUE NAO POSSUEM EVENTOS
            aTipoMovimento(1) = 1
            aValor(1) = dblValorTotalReceita
            
            If Not GeraMovimentosByEvento(lngEmpenho, dtmDatabyevento, Str(dblValorTotalReceita), "", lngUltNumArrecadacao, "6", aContaExtra, aTipoMovimento, IIf(UBound(aContaExtra) > 1, True, False), aValor, True) Then
               ExibeMensagem "Ocorreram erros durante a gravação dos movimentos do Evento Contabil."
               gobjBanco.ExecutaRollbackTrans
               Exit Sub
            End If
        End If

        gobjBanco.ExecutaCommitTrans
        
    End If
    
    If UBound(aCriticas) > 0 Then

        Open "C:\Desenv\Criticas" & Replace(dtmDataEmOperacao, "/", "") & ".txt" For Output As #1
        For intFor = 1 To UBound(aCriticas)
            Print #1, aCriticas(intFor) & ", "
        Next
        Close #1

    End If
    
    lblStatus.Caption = ""
    MsgBox "Importação concluída com sucesso", vbInformation
    Exit Sub
    
ErroAbreBancoDados:
    
    ExibeDetalheErro Err.Description
    blnErroServidor = True
    Resume Next
    lblStatus.Caption = ""
    cncADO.Close
    Set cncADO = Nothing
End Sub

Private Sub Form_Load()
    txt_Final.Text = gstrDataDoSistema
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
End Sub

Sub LimpaObjetos()
    txt_Inicial = ""
    txt_Final = ""
    txt_Inicial.SetFocus
End Sub

Private Function blnDadosOK() As Boolean
    
    blnDadosOK = False
    
    If txt_Inicial.Text = "" Then
        ExibeMensagem "O campo data inicial deve ser digitado."
        txt_Inicial.SetFocus
        Exit Function
    End If
 
    If txt_Final.Text = "" Then
        ExibeMensagem "O campo data Final deve ser digitado."
        txt_Final.SetFocus
        Exit Function
    End If
    
    If CVDate(txt_Inicial) > CVDate(txt_Final) Then
        ExibeMensagem "A data Inicial tem que ser anterior à data Final."
        txt_Final.SetFocus
        Exit Function
    End If
    
    blnDadosOK = True
    
End Function

Private Sub txt_Final_GotFocus()
    MarcaCampo txt_Final
End Sub

Private Sub txt_Final_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_Final
End Sub

Private Sub txt_Final_LostFocus()
    txt_Final = gstrDataFormatada(txt_Final)
End Sub

Private Sub txt_Inicial_GotFocus()
    MarcaCampo txt_Inicial
End Sub

Private Sub txt_Inicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_Inicial
End Sub

Private Sub txt_Inicial_LostFocus()
    txt_Inicial = gstrDataFormatada(txt_Inicial)
End Sub

Private Function AcertaDiferenca(aArray As XArrayDB, intRowCol As Integer, dblDiferenca As Double, blnAcertoEmReceita As Boolean) As XArrayDB
Dim intFor As Integer
    
    If blnAcertoEmReceita Then
    
        'Vamos varrer as colunas
        For intFor = 1 To aArray.UpperBound(2)
        
            'Vamos verificar se o conteudo e maior que zero
            If CCur(aArray(intRowCol, intFor)) <> 0 Then
                aArray(intRowCol, intFor) = aArray(intRowCol, intFor) + dblDiferenca
                Exit For
            End If
        
        Next
    
    Else
    
        'Vamos varrer as linhas
        For intFor = 0 To aArray.UpperBound(1)
        
            'Vamos verificar se o conteudo e maior que zero
            If CCur(aArray(intFor, intRowCol)) <> 0 Then
                aArray(intFor, intRowCol) = aArray(intFor, intRowCol) + dblDiferenca
                aArray(intFor, intRowCol + 1) = aArray(intFor, intRowCol + 1) - dblDiferenca
                Exit For
            End If
        
        Next
    
    End If
    
    Set AcertaDiferenca = aArray
    
End Function

Private Function RetornaReceitaDeReferencia(strReceitaOrigem As String) As String
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset

    strSql = "SELECT Destino FROM tblTempReferenciasDeDespesa WHERE Origem = '" & strReceitaOrigem & "'"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        
        If Not adoResultado.EOF Then
            RetornaReceitaDeReferencia = adoResultado("Destino").Value
        Else
            'RetornaReceitaDeReferencia = strReceitaOrigem
            
            ReDim Preserve aCriticas(UBound(aCriticas) + 1)
            aCriticas(UBound(aCriticas)) = strReceitaOrigem
            
            RetornaReceitaDeReferencia = "-1"
            
        End If
    
    End If
    
    adoResultado.Close: Set adoResultado = Nothing

End Function
