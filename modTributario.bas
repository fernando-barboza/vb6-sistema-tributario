Attribute VB_Name = "modTributario"
Option Explicit

'CHARLES
    Public gfrmFormularioQueEstaImprimindoGuia As Form

    'HTML Help
    Public hHelp As New clsHTMLHelp
    
    Private Declare Function FindWindowEx Lib "USER32" _
    Alias "FindWindowExA" _
    (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

    Private Declare Function SendTBMessage Lib "USER32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Integer, _
    ByVal lParam As Any) As Long
    
    Private Const WM_USER = &H400
    Private Const TB_SETSTYLE = (WM_USER + 56)
    Private Const TB_GETSTYLE = (WM_USER + 57)

    Private Const TBSTYLE_FLAT = &H800
    Private Const TBSTYLE_LIST = &H1000
    Private Const TBSTYLE_TRANSPARENT = &H8000
    Private Const CCS_NODIVIDER = &H40
        
    Public Const LVS_EX_HEADERDRAGDROP = &H10
    Public Const LVS_EX_FULLROWSELECT = &H20
        
    Type CondicaoRelatorioGerado
        strCondicao    As String
    End Type

    Public typCondicaoDoRelatorio()   As CondicaoRelatorioGerado
    
    Type ClassificacaoRelatorioGerado
        strClassificacao    As String
    End Type
    
    Public typClassificacaoDoRelatorio()   As ClassificacaoRelatorioGerado
    
    Type GeraRelatorio
        strNomeDoCampo      As String
        strNomeDaDescricao  As String
    End Type
    
    Type Campo
        strDescricao     As String
        intPosicao       As Integer
        intTamanho       As Integer
        intTipo          As String
        blnVirgula       As Boolean
        strCasasDecimais As String
    End Type
    
    Public typCampoDoRelatorio()   As GeraRelatorio
        
    'Constantes com os nomes das tabelas
    
    
    Public Const gstrDescontosProvisorios = "tblDesctosProvisorios"
    Public Const gstrCriticaIptu = "Tblcriticaiptu"
    Public Const gstrParametroDividaAtiva = "tblParametroDividaAtiva"
    Public Const gstrParametroAtualizacaoMulta = "tblParametroAtualizacaoMulta"
    Public Const gstrAcordo = "tblAcordo"
    Public Const gstrFormaPagtoCancelamentos = "tblformapagtocancelamentos"
    Public Const gstrCriticaBaixa = "tblCriticaBaixa"
    Public Const gstrTipoCriticaBaixa = "tblTipoCriticaBaixa"
    Public Const gstrMovimentoBancario = "tblMovimentoBancario"
    'Public Const gstrCategoriaConstrucao = "tblCategoriaConstrucao"
    Public Const gstrTestadaIPTU = "tblTestadaIPTU"
    Public Const gstrFormaPagtoVencimentos = "tblFormaPagtoVencimentos"
    Public Const gstrParametroIPTUPagto = "tblParametroIPTUPagto"
    Public Const gstrParametroIPTU = "tblParametroIPTU"
    Public Const gstrCaracBoletimIPTU = "tblCaracBoletimIPTU"
    Public Const gstrLancamentoPredioIPTU = "tblLancamentoPredioIPTU"
    Public Const gstrLancamentoFatores = "tblLancamentoFatores"
    Public Const gstrLancamentoEnvolvidos = "tblLancamentoEnvolvidos"
    Public Const gstrLancamentoIPTU = "tblLancamentoIPTU"
    Public Const gstrImovel = "tblImovel"
    Public Const gstrSecao = "tblSecao"
    Public Const gstrSocioContribuinte = "tblSocioContribuinte"
    Public Const gstrQualificacao = "tblQualificacao"
    Public Const gstrTabelaDocumento = "tblDocumento"
    Public Const gstrOrigem = "tblOrigem"
    Public Const gstrSituacaoDebito = "tblSituacao"
    Public Const gstrTipoContribuinte = "tblTipoDeContribuinte"
    Public Const gstrGrupoContribuinte = "tblGrupoDeContribuinte"
    Public Const gstrRegiao = "tblRegiao"
    Public Const gstrMensagemNaGuia = "tblMensagemNaGuia"
    Public Const gstrPeriodicidade = "tblPeriodicidade"
    Public Const gstrFormacao = "tblFormacao"
    Public Const gstrFormaLancmto = "tblFormaLancamento"

    
    Public Const gstrTipoDeAtividade = "tblTipoDeAtividade"
    Public Const gstrSegmentoDeAtividade = "tblSegmentoDeAtividade"
    Public Const gstrServicoContribuinte = "tblServicoContribuinte"
    Public Const gstrServicoPrestado = "tblServicoPrestado"

    
    Public Const gstrParametroDeLancamento = "tblParametroDeLancamento"
    Public Const gstrServicoUrbano = "tblServicoUrbano"
    Public Const gstrPlantaDeValor = "tblPlantaDeValor"
    'Cl�udio passada esta constante para o modGeral, pois agora ser� usada tamb�m em Patrim�nio
    'Public Const gstrLoteamento = "tblLoteamento"
    Public Const gstrCustoUnitario = "tblCustoUnitarioObra"
    Public Const gstrTipoItemDespesa = "tblTipoItemDespesa"
    Public Const gstrObra = "tblObra"
    Public Const gstrTipoObra = "tblTipoObra"
    Public Const gstrTipoRateio = "tblTipoRateio"
    Public Const gstrLancamentoDoRateio = "tblLancamentoDoRateio"
    Public Const gstrCustoDeDistribuicao = "tblCustoDeDistribuicao"
    Public Const gstrRural = "tblRural"
    Public Const gstrNaturezaDoImovel = "tblNaturezaDoImovel"
    Public Const gstrBenfeitoriaDoImovel = "tblBenfeitoriaDoImovel"
    Public Const gstrCartorio = "tblCartorio"
    Public Const gstrTipoDaTransacao = "tblTipoDaTransacao"
    Public Const gstrQualidadeDeConservacao = "tblQualidadeDeConservacao"
    Public Const gstrAutorizacaoDeLicencas = "tblAutorizacaoDeLicencas"
    Public Const gstrNotasFiscais = "tblNotasFiscais"
    Public Const gstrImoveisForeiros = "tblImoveisForeiros"
    Public Const gstrTransmicaoDeImovel = "tblTransmicaoDeImovel"
    Public Const gstrResponsavel = "tblResponsavel"
    Public Const gstrTramitacao = "tblTramitacao"
    Public Const gstrProcessoEmJulgamento = "tblProcessoEmJulgamento"
    Public Const gstrNotaProdutorRural = "tblNotaProdutorRural"
    Public Const gstrMovimentacaoEconomica = "tblMovimentacaoEconomica"
    Public Const gstrAgenteFiscal = "tblAgenteFiscal"
    Public Const gstrOrdemServico = "tblOrdemServico"
    Public Const gstrOrdemServicoFiscal = "tblOrdemServicoFiscal"
    Public Const gstrAgenteArrecadador = "tblAgenteArrecadador"
    Public Const gstrReceitaNaoTributaria = "tblReceitaNaoTributaria"
    Public Const gstrPlanoDeConta = "tblPlanoDeConta"
    Public Const gstrDispositivoLegal = "tblDispositivoLegal"
    Public Const gstrClassificacaoDoAuto = "tblClassificacaoDoAuto"
    Public Const gstrAutoDaInfracao = "tblAutoDaInfracao"
    Public Const gstrContencioso = "tblContencioso"
    Public Const gstrRedutorDeProfundidade = "tblRedutorDeProfundidade"
    Public Const gstrAcidenteGeografico = "tblAcidenteGeografico"
    Public Const gstrPadraoDeAcabamento = "tblPadraoDeAcabamento"
    Public Const gstrConstrucaoPorUtilizacao = "tblConstrucaoPorUtilizacao"
    Public Const gstrLocalizacaoPorEspecie = "tblLocalizacaoPorEspecie"
    Public Const gstrDepreciacaoPorIdade = "tblDepreciacaoPorIdade"
    Public Const gstrReducaoPorArea = "tblReducaoPorArea"
    Public Const gstrEquipamento = "tblEquipamento"
    Public Const gstrAliquota = "tblAliquota"
    Public Const gstrIncidencia = "tblIncidencia"
    Public Const gstrMemoriaDeCalculoIPTU = "tblMemoriaDeCalculoIPTU"
    Public Const gstrIndicesDiversos = "tblIndicesDiversos"
    Public Const gstrTaxaLimpezaPublica = "tblTaxaLimpezaPublica"
    Public Const gstrBIC = "tblBIC"
    Public Const gstrVencimentos = "tblVencimentos"
    Public Const gstrVencimentosDasParcelas = "tblVencimentosDasParcelas"
    Public Const gstrTipoDeInscricao = "tblTipoDeInscricao"
    Public Const gstrUtilizacaoDaOcorrencia = "tblUtilizacaoDaOcorrencia"
    Public Const gstrPagamentoParcela = "tblPagamentoParcela"
    Public Const gstrDescricaoLayout = "tblDescricaoLayout"
    Public Const gstrLayoutColuna = "tblLayoutColuna"
    Public Const gstrFonteDeRecurso = "tblFonteRecurso"
    
    Public Const gstrIssConstrucaoTipo = "tblIssConstrucaoTipo"
    Public Const gstrIssConstrucaoPadrao = "tblIssConstrucaoPadrao"
    Public Const gstrIssConstrucaoExercicio = "tblIssConstrucaoExercicio"
    Public Const gstrIssConstrucaoVlrM2 = "tblIssConstrucaoVlrM2"
    Public Const gstrLanctoIssConstrucao = "tblLanctoIssConstrucao"
    Public Const gstrLanctoIssConstrucaoPredios = "tblLanctoIssConstrucaoPredios"
    Public Const gstrTipoPadraoExercicio = "tblTipoPadraoExercicio"
    
    'Transferidos para modGeral, Compartilhado com Compras
    'Public Const gstrGrupoDeAtividade = "tblGrupoDeAtividade"
    'Public Const gstrSubGrupoDeAtividade = "tblSubGrupoDeAtividade"
    
    Public Const gstrTextoLivre = "tblTextoLivre"
    Public Const gstrFormaDeComunicacaoContador = "tblFormaDeComunicacaoContador"
    
    Public Const gstrTabelaDeApoio = "tblTabelaDeApoioCaracteristicaImovel"
    
    Public Const gstrDevolucao = "tblDevolucao"
    Public Const gstrEmissaoValidade = "tblEmissaoValidade"
    Public Const gstrFatorDeCorrecao = "tblFatorDeCorrecao"
    Public Const gstrFiscais = "tblFiscal"
    Public Const gstrContNotasFiscais = "tblContNotaFiscal"
    
    
    Public Const gstrReceitaDiversaValor = "tblReceitaDiversaValor"
    
    Public Const gstrMapaAcaoFiscal = "tblMapaAcaoFiscal"
    Public Const gstrMapaAcaoFiscalDocumento = "tblMapaAcaoFiscalDocumento"
    Public Const gstrSuspensaoDeExigencia = "tblSuspensaoDeExigencia"
    Public Const gstrAutoDeInfracao = "tblAutoDeInfracao"
    Public Const gstrFormulaBasica = "tblFormulaBasica"
    
    'Constantes de encargos
    Public Const BIT_HONORARIOS        As Byte = 1
    Public Const BIT_DIRIGENCIAS       As Byte = 2
    Public Const BIT_CUSTAS            As Byte = 3
    
    Global VerificaFormAtivo As Boolean
    
'---- Chave - KEY - dos bot�es comuns da barra de ferramentas dos sistemas

    'Array utilizados para impressao do termo de acordo
    Dim XParcelas                   As XArrayDB
    Dim XArrayAlinhaColunas         As XArrayDB

    Type Pais
        NomePai        As String
        OcupacaoPai    As String
        NomeMae        As String
        OcupacaoMae    As String
    End Type
    
    Public vetPais() As Pais

    Type TabelasGerador
        NomeTable As String
    End Type
    
    Public gstrTabelaAtual      As String
    
    
    Public Const gstrLerArquivo = "LERARQUIVO"
    Dim Documentos()    As cWordWrapper
    
    Public aAnaliseReceita             As XArrayDB
    
Function gstrConteudoOuDescricao(vntCampo As Variant, vntConteudo As Variant) As String
    Dim adoResultado As ADODB.Recordset
    Dim strSQL       As String
    
    Select Case UCase(Trim(vntCampo))
        Case "SEXO", "SEXORESPONSAVEL"
            gstrConteudoOuDescricao = gstrMOuF(Not CBool(vntConteudo), True)
        Case "ATINGIUMAIORIDADE", "ESTUDA", "TRAJETORIA", "USUARIODROGA", "RESPONSAVEL"
            gstrConteudoOuDescricao = gstrSimOuNao(vntConteudo)
        Case "SITUACAO"
            gstrConteudoOuDescricao = gstrSituacao(CByte(vntConteudo))
        Case "ENVOLVIDO"
            gstrConteudoOuDescricao = gstrEnvolvido(CByte(vntConteudo))
        Case "CONVENIO"
            strSQL = "Select Descricao From Convenio Where PKId = " & vntConteudo
        Case "OLHOS"
            strSQL = "Select Descricao From TipoOlho Where PKId = " & vntConteudo
        Case "ENERGIAELETRICA"
            strSQL = "Select Descricao From RedeEletrica Where PKId = " & vntConteudo
        Case "REDEDEESGOTO"
            strSQL = "Select Descricao From RedeEsgoto Where PKId = " & vntConteudo
        Case "REDEDAGUA"
            strSQL = "Select Descricao From RedeDagua Where PKId = " & vntConteudo
        Case "COLETADELIXO"
            strSQL = "Select Descricao From ColetaLixo Where PKId = " & vntConteudo
        Case "TIPODECONSTRUCAO"
            strSQL = "Select Descricao From TipoConstrucao Where PKId = " & vntConteudo
        Case "IMOVEL"
            strSQL = "Select Descricao From Imovel Where PKId = " & vntConteudo
        Case "AUXILIORENDA"
            strSQL = "Select Descricao From AuxilioRenda Where PKId = " & vntConteudo
        Case "MOTIVOENVOLVIDO"
            strSQL = "Select Descricao From Motivo Where PKId = " & vntConteudo
        Case "TECNICORESPONSAVEL"
            strSQL = "Select Nome As Descricao From Tecnico Where PKId = " & vntConteudo
        Case "FUNCAOCARGO"
            strSQL = "Select Descricao From Ocupacao Where PKId = " & vntConteudo
        Case "LOCALTRABALHO"
            strSQL = "Select Descricao From LocalTrabalho Where PKId = " & vntConteudo
        Case "CODREGIONAL"
            strSQL = "Select Descricao From Regional Where PKId = " & vntConteudo
        Case "RESPONSAVELPGTR"
            strSQL = "Select Nome As Descricao From ResponsavelPGTR Where PKId = " & vntConteudo
        Case "ENCAMINHADO"
            strSQL = "Select Descricao From Encaminhador Where PKId = " & vntConteudo
        Case "RENDAFAMILIAR"
            gstrConteudoOuDescricao = gvntConvVrDoSql(vntConteudo)
        Case "REGIAOPGTR"
            strSQL = "Select Descricao From Regiao Where PKId = " & vntConteudo
        Case "FAZTRATAMENTO"
            gstrConteudoOuDescricao = gstrSimOuNao(vntConteudo)
        Case "DEPENDENTEQUIMICO"
            gstrConteudoOuDescricao = gstrSimOuNao(vntConteudo)
        Case "DOENCACLONICA"
            gstrConteudoOuDescricao = gstrSimOuNao(vntConteudo)
        Case "PROCEDENCIA"
            strSQL = "Select Descricao From Procedencia Where PKId = " & vntConteudo
        Case "TURNO"
            strSQL = "Select Descricao From Turno Where PKId = " & vntConteudo
        Case "SERIE"
            strSQL = "Select Descricao From Serie Where PKId = " & vntConteudo
        Case "REGIAO"
            strSQL = "Select Descricao From Regiao Where PKId = " & vntConteudo
        Case "CASO"
            strSQL = "Select Descricao From Caso Where PKId = " & vntConteudo
        Case "ORIGEM"
            strSQL = "Select Descricao From Origem Where PKId = " & vntConteudo
        Case "CODCIDADE"
            strSQL = "Select Descricao From Cidade Where PKId = " & vntConteudo
        Case "CODBAIRRO"
            strSQL = "Select Descricao From Bairro Where PKId = " & vntConteudo
        Case "LOGRADOURO"
            strSQL = "Select Descricao From Logradouro Where PKId = " & vntConteudo
        Case "MARCAS"
            strSQL = "Select Descricao From Marca Where PKId = " & vntConteudo
        Case "CUTIS"
            strSQL = "Select Descricao From Cutis Where PKId = " & vntConteudo
        Case "ESTADOCIVIL"
            strSQL = "Select Descricao From EstadoCivil Where PKId = " & vntConteudo
        Case "NATURALIDADE"
            strSQL = "Select Descricao From Cidade Where PKId = " & vntConteudo
        Case "ESCOLARIDADE"
            strSQL = "Select Descricao From Escolaridade Where PKId = " & vntConteudo
        Case "OLHOS"
            strSQL = "Select Descricao From TipoOlho Where PKId = " & vntConteudo
        Case "RENDA"
            gstrConteudoOuDescricao = gvntConvVrDoSql(vntConteudo)
        Case "OCUPACAO"
            strSQL = "Select Descricao From Ocupacao Where PKId = " & vntConteudo
        Case "PARENTESCO"
            strSQL = "Select Descricao From Parentesco Where PKId = " & vntConteudo
        Case ""
            gstrConteudoOuDescricao = ""
        Case ""
            gstrConteudoOuDescricao = ""
        Case ""
            gstrConteudoOuDescricao = ""
        Case ""
            gstrConteudoOuDescricao = ""
        Case ""
            gstrConteudoOuDescricao = ""
        Case ""
            gstrConteudoOuDescricao = ""
        Case ""
            gstrConteudoOuDescricao = ""
            
        Case Else
            gstrConteudoOuDescricao = gstrVerificaCampoNulo(vntConteudo)
    End Select
    
    If strSQL <> "" Then
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                If Not .EOF Then
                    gstrConteudoOuDescricao = !descricao
                End If
            End With
            Set adoResultado = Nothing
        End If
    End If
End Function

Public Function gstrSituacao(bytInd As Byte) As String
    Select Case bytInd
        Case 0
            gstrSituacao = "Desativado"
        Case 1
            gstrSituacao = "Ativo"
    End Select
End Function

Public Function gstrEnvolvido(bytInd As Byte) As String
    Select Case bytInd
        Case 0
            gstrEnvolvido = "Juizado"
        Case 1
            gstrEnvolvido = "Conselho Tutelar"
    End Select
End Function

Public Function gstrNomeCampoNaTabela(vntCampo As Variant) As String
    Select Case UCase(Trim(vntCampo))
        Case "DESCRICAO"
            gstrNomeCampoNaTabela = "Descri��o"
        Case "NOME"
            gstrNomeCampoNaTabela = "Nome"
        Case "TRAJETORIA"
            gstrNomeCampoNaTabela = "Trajet�ria de Rua"
        Case "USUARIODROGA"
            gstrNomeCampoNaTabela = "Usu�rio de Drogas"
        Case "ATINGIUMAIORIDADE"
            gstrNomeCampoNaTabela = "Atingiu Maioridade"
        Case "REDEDAGUA"
            gstrNomeCampoNaTabela = "Rede de �gua"
        Case "CONVENIO"
            gstrNomeCampoNaTabela = "Conv�nio"
        Case "ENERGIAELETRICA"
            gstrNomeCampoNaTabela = "Energia el�trica"
        Case "IMOVEL"
            gstrNomeCampoNaTabela = "Im�vel"
        Case "TIPODECONSTRUCAO"
            gstrNomeCampoNaTabela = "Tipo de Constru��o"
        Case "REDEDEESGOTO"
            gstrNomeCampoNaTabela = "Rede de Esgoto"
        Case "COLETADELIXO"
            gstrNomeCampoNaTabela = "Coleta de Lixo"
        Case "AUXILIORENDA"
            gstrNomeCampoNaTabela = "Aux�lio Renda"
        Case "PROCEDENCIA"
            gstrNomeCampoNaTabela = "Proced�ncia"
        Case "DOENCACLONICA"
            gstrNomeCampoNaTabela = "Tem doen�a cr�nica"
        Case "DEPENDENTEQUIMICO"
            gstrNomeCampoNaTabela = "Dependente qu�mico"
        Case "FAZTRATAMENTO"
            gstrNomeCampoNaTabela = "Faz tratamento"
        Case "DESCRICAODADOENCA"
            gstrNomeCampoNaTabela = "Descri�ao da doen�a cr�nica"
        Case "TIPODROGA"
            gstrNomeCampoNaTabela = "Descri��o dos tipos de drogas"
        Case "TRATAMENTO"
            gstrNomeCampoNaTabela = "Descri��o do tratamento"
        Case "POSICAOFORUM"
            gstrNomeCampoNaTabela = "Posi��o no F�rum"
        Case "DATAREGISTRO"
            gstrNomeCampoNaTabela = "Data de Registro"
        Case "NUMEROCASO"
            gstrNomeCampoNaTabela = "N� do Caso"
        Case "CASO"
            gstrNomeCampoNaTabela = "Descri��o do Caso"
        Case "COMENTARIOSFINAIS"
            gstrNomeCampoNaTabela = "Coment�rios Finais"
        Case "SITUACAO"
            gstrNomeCampoNaTabela = "Situa��o"
        Case "CODTIPOLOGRADOURO"
            gstrNomeCampoNaTabela = "Tipo de Logradouro"
        Case "LOGRADOURO"
            gstrNomeCampoNaTabela = "Logradouro"
        Case "CODBAIRRO"
            gstrNomeCampoNaTabela = "Bairro"
        Case "CODCIDADE"
            gstrNomeCampoNaTabela = "Cidade"
        Case "NUMERO"
            gstrNomeCampoNaTabela = "N�"
        Case "TELEFONE"
            gstrNomeCampoNaTabela = "Fone"
        Case "REGIAO"
            gstrNomeCampoNaTabela = "Regi�o (Endere�o)"
        Case "DATADOREGISTRO"
            gstrNomeCampoNaTabela = "Data de Cadastramento"
        Case "DATANACIMENTO"
            gstrNomeCampoNaTabela = "Data de Nascimento"
        Case "ESTADOCIVIL"
            gstrNomeCampoNaTabela = "Estado Civil"
        Case "CUTIS"
            gstrNomeCampoNaTabela = "Cor"
        Case "CODFOTO"
            gstrNomeCampoNaTabela = "Foto"
        Case "NOMEESCOLA"
            gstrNomeCampoNaTabela = "Escola"
        Case "SERIE"
            gstrNomeCampoNaTabela = "S�rie"
        Case "LOCALTRABALHO"
            gstrNomeCampoNaTabela = "Lota��o"
        Case "FUNCAOCARGO"
            gstrNomeCampoNaTabela = "Cargo/Fun��o"
        Case "TECNICORESPONSAVEL"
            gstrNomeCampoNaTabela = "T�cnico Respons�vel"
        Case "ENCAMINHADO"
            gstrNomeCampoNaTabela = "Encaminhado Por"
        Case "RESPONSAVELPGTR"
            gstrNomeCampoNaTabela = "Respons�vel"
        Case "OBSTRABALHISTA"
            gstrNomeCampoNaTabela = "Motivo do Desligamento"
        Case "DATADESLIGAMENTO"
            gstrNomeCampoNaTabela = "Data do Desligamento"
        Case "REGIAOPGTR"
            gstrNomeCampoNaTabela = "Regi�o (PGTR)"
        Case "NUMEROFILHOS"
            gstrNomeCampoNaTabela = "N� de Filhos"
        Case "NUMEROIRMAOS"
            gstrNomeCampoNaTabela = "N� de Irm�os"
        Case "MOTIVOENVOLVIDO"
            gstrNomeCampoNaTabela = "Motivo do Envolvimento"
        Case "CODREGIONAL"
            gstrNomeCampoNaTabela = "Regional"
        Case "RENDAFAMILIAR"
            gstrNomeCampoNaTabela = "Renda Familiar"
        Case "OCUPACAO"
            gstrNomeCampoNaTabela = "Ocupa��o (Parente)"
        Case "SEXORESPONSAVEL"
            gstrNomeCampoNaTabela = "Sexo (Parente)"
        Case "RESPONSAVEL"
            gstrNomeCampoNaTabela = "Respons�vel (Parente)"
        Case "DATANACIMENTOPARENTE"
            gstrNomeCampoNaTabela = "Data Nascimento (Parente)"
        Case "RENDA"
            gstrNomeCampoNaTabela = "Renda (Parente)"
        Case "NOMEPARENTE"
            gstrNomeCampoNaTabela = "Nome (Parente)"
        Case ""
            gstrNomeCampoNaTabela = ""
        Case ""
            gstrNomeCampoNaTabela = ""
        Case ""
            gstrNomeCampoNaTabela = ""

        Case Else
            gstrNomeCampoNaTabela = Trim(vntCampo)
    End Select
End Function

Private Sub MontaBotoes(chd, intIndice As Integer, strMenu As String, strTitulo As String, _
                       Optional blnSubBand As Boolean = False, _
                       Optional strNomeSubBand As String, _
                       Optional strToolTip As String, _
                       Optional strTagVariant As String)

    Dim Tool As ActiveBar2LibraryCtl.Tool
    
    Set Tool = chd.Tools.Add(intIndice, strMenu)
    
    With Tool
        .Alignment = ddALeftTop
        .CaptionPosition = ddCPStandard
        If blnSubBand Then
            .SubBand = strNomeSubBand
            .ControlType = ddTTButton
            .MenuVisibility = ddMVVisibleIfRecentlyUsed
            .SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("PASTAFECHADA").Picture
        Else
            .ControlType = ddTTButton
            .SetPicture ddITNormal, MDIMenu.img_ListaIconesGeral.ListImages("PASTAABERTA").Picture
        End If
    
        .Style = ddSIconText
        
        .Caption = strTitulo
    
        If strToolTip <> "" Then
            .ToolTipText = Trim(strToolTip)
        Else
            .ToolTipText = Trim(strTitulo)
        End If
        
        If strTagVariant <> "" Then
            .TagVariant = Trim(strTagVariant)
        End If
        
        .Category = Right(strMenu, Len(strMenu) - 3)
                
        'Ou para o caso de Documentos do Word
        If Trim(strNomeSubBand) = "" Or Trim(strNomeSubBand) = "chdArqSubWordModelosGravados" Then
            If Not gblnVerificaPermissoes(intIndice, "bndFormulario", True) Then
                .Enabled = False
            End If
        End If

    End With
    
End Sub



Private Sub MontaSubBandImobiliariasUrbanas(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubTabelasImobiliariasUrbanas"
    .DockingArea = ddDAPopup
    .Caption = "Tabelas"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 611, "mnuSubTabelasImobili�riasUrbanas", "Melhoramentos P�blicos"
MontaBotoes chd, 609, "mnuSubTabelasImobili�riasUrbanas", "Tipo de �rea"
MontaBotoes chd, 610, "mnuSubTabelasImobili�riasUrbanas", "Tipos de Testada      "
MontaBotoes chd, 1025, "mnuSubTabelasImobili�riasUrbanas", "Valor Metro Terreno"

'MontaBotoes chd, 612, "mnuSubTabelasImobili�riasUrbanas", "Se��es de Logradouro"
'MontaBotoes chd, 613, "mnuSubTabelasImobili�riasUrbanas", "Fatores de Corre��o" '- RETIRADO 26/07/04 Rafael

End Sub

Private Sub MontaSubBandEconomicas(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubTabelasEconomicas"
    .DockingArea = ddDAPopup
    .Caption = "Tabelas"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 1068, "mnuSubTabelasEconomicas", "Atividades B�sicas   "
MontaBotoes chd, 615, "mnuSubTabelasEconomicas", "Atividades Econ�micas "
MontaBotoes chd, 1160, "mnuSubTabelasEconomicas", "Feiras               "
MontaBotoes chd, 1158, "mnuSubTabelasEconomicas", "Servi�os             "
MontaBotoes chd, 1159, "mnuSubTabelasEconomicas", "Tipos de Feira       "
MontaBotoes chd, 1173, "mnuSubTabelasEconomicas", "Tipos de Tributos    "
MontaBotoes chd, 1157, "mnuSubTabelasEconomicas", "Tributos             "
'MontaBotoes chd, 1294, "mnuSubTabelasEconomicas", "Ocorr�ncia - Processo"

End Sub

Private Sub MontaSubBandContribuicaoDeMelhorias(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubTabelasContribuicaoDeMelhorias"
    .DockingArea = ddDAPopup
    .Caption = "Tabelas"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 617, "mnuSubTabelasContribuicaoDeMelhorias", "Editais "

End Sub

Private Sub MontaSubBandGerais(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubTabelasGerais"
    .DockingArea = ddDAPopup
    .Caption = "Tabelas"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 1393, "mnuSubTabelasGerais", "Advogados                    "
MontaSubBandAgentesArrecadadores ab
MontaBotoes chd, 589, "mnuSubTabelasGerais", "Agentes Arrecadadores...   ", True, "chdSubTabelaGeraisAgentesArrecadadores"
MontaBotoes chd, 598, "mnuSubTabelasGerais", "Caracter�sticas de Boletins"
MontaBotoes chd, 1120, "mnuSubTabelasGerais", "C�digos de Baixa         "
MontaBotoes chd, 1299, "mnuSubTabelasGerais", "Descontos Provis�rios    "
MontaBotoes chd, 758, "mnuSubTabelasGerais", "Documentos Emitidos       "
MontaBotoes chd, 607, "mnuSubTabelasGerais", "Fiscais                   "
MontaBotoes chd, 578, "mnuSubTabelasGerais", "F�rmulas de C�lculo        "
MontaBotoes chd, 1117, "mnuSubTabelasGerais", "Indexador Econ�mico      "
MontaBotoes chd, 1119, "mnuSubTabelasGerais", "Moedas                   "
MontaBotoes chd, 588, "mnuSubTabelasGerais", "Ocorr�ncias                "
MontaBotoes chd, 602, "mnuSubTabelasGerais", "Par�metros                "
MontaSubBandPlantasDeValores ab
MontaBotoes chd, 592, "mnuSubTabelasGerais", "Planta de Valores...       ", True, "chdSubTabelaGeraisPlantasDeValores"
MontaSubBandReceitaDoMunicipio ab
MontaBotoes chd, 595, "mnuSubTabelasGerais", "Receita do Munic�pio...    ", True, "chdSubTabelaGeraisReceitaDoMunicipio"
MontaSubBandTextos ab
MontaBotoes chd, 604, "mnuSubTabelasGerais", "Textos...                 ", True, "chdSubTabelaGeraisTextos"
MontaBotoes chd, 601, "mnuSubTabelasGerais", "Tipos de Comunica��o      "
MontaBotoes chd, 6, "mnuSubTabelasGerais", "Tipos de Documento        "
MontaBotoes chd, 1148, "mnuSubTabelasGerais", "Tipos de Isen��o e Imunidade"
MontaBotoes chd, 8, "mnuSubTabelasGerais", "Unidades de Medida        "


MontaBotoes chd, 599, "mnuSubTabelasGerais", "Dias n�o �teis             "

'MontaBotoes chd, 397, "mnuSubTabelasGerais", "�ndices Econ�micos         "
'MontaBotoes chd, 600, "mnuSubTabelasGerais", "Vencimento de Parcela     "
'ALtera�ao feita por hugo
End Sub
Private Sub MontaSubBandLogradouros(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubTabelasLogradouros"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 581, "mnuSubTabelasLogradouros", "Bairros"
MontaBotoes chd, 585, "mnuSubTabelasLogradouros", "Distritos Fiscais"
MontaBotoes chd, 1026, "mnuSubTabelasLogradouros", "Face de Quadra"
MontaBotoes chd, 584, "mnuSubTabelasLogradouros", "Logradouros"
MontaBotoes chd, 587, "mnuSubTabelasLogradouros", "Loteamentos"
MontaBotoes chd, 53, "mnuSubTabelasLogradouros", "Munic�pios"
MontaBotoes chd, 586, "mnuSubTabelasLogradouros", "Setores Fiscais"
MontaBotoes chd, 582, "mnuSubTabelasLogradouros", "Tipos de Logradouro"
MontaBotoes chd, 1058, "mnuSubTabelasLogradouros", "Tipos de Vias"
MontaBotoes chd, 583, "mnuSubTabelasLogradouros", "T�tulos de Logradouro"

End Sub
Private Sub MontaSubBandPlantasDeValores(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubTabelaGeraisPlantasDeValores"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 594, "mnuSubTabelaGeraisPlantasDeValores", "Faixa de Valores"
MontaBotoes chd, 593, "mnuSubTabelaGeraisPlantasDeValores", "Valores"


End Sub
Private Sub MontaSubBandAgentesArrecadadores(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubTabelaGeraisAgentesArrecadadores"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 591, "mnuSubTabelaGeraisAgentesArrecadadores", "Ag�ncias"
MontaBotoes chd, 590, "mnuSubTabelaGeraisAgentesArrecadadores", "Bancos"

End Sub
Private Sub MontaSubBandReceitaDoMunicipio(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubTabelaGeraisReceitaDoMunicipio"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 445, "mnuSubTabelaGeraisReceitaDoMunicipio", "Composi��o Das Receitas"
MontaBotoes chd, 444, "mnuSubTabelaGeraisReceitaDoMunicipio", "Receitas"

End Sub
Private Sub MontaSubBandTextos(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubTabelaGeraisTextos"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 605, "mnuSubTabelaGeraisTextos", "Mensagens"
MontaBotoes chd, 606, "mnuSubTabelaGeraisTextos", "Textos Livres"

End Sub

Private Sub MontaMenuPrincipal(ab As ActiveBar2, chd)

Set chd = ab.Bands("bndFormulario").ChildBands.Add("ChildBand")
ab.Bands("bndFormulario").ChildBands.ChildBandCaptionAlignment = ddCACenter
ab.Bands("bndFormulario").ChildBands.BackColor = &HE0E0E0
With chd
    .Name = "chdPrincipal"
    .Caption = "Principal"
    .GrabHandleStyle = ddGSCaption
    .flags = 127
    .Width = 12000
End With

'MontaBotoes chd, 573, "mnuPrincipal", "Entidade                "
MontaBotoes chd, 3000, "mnuPrincipal", "Efetuar Logoff de " & Trim(gstrNomeUsuario) & "..."
MontaBotoes chd, 3001, "mnuPrincipal", "Sair...                "
End Sub


Private Sub MontaMenuTabelas(ab As ActiveBar2, chd)

Set chd = ab.Bands("bndFormulario").ChildBands.Add("ChildBand")
ab.Bands("bndFormulario").ChildBands.ChildBandCaptionAlignment = ddCACenter
ab.Bands("bndFormulario").ChildBands.BackColor = &HE0E0E0
With chd
    .Name = "chdTabelas"
    .Caption = "Tabelas"
    .GrabHandleStyle = ddGSCaption
    .flags = 127
    .Width = 12000
End With

MontaSubBandGerais ab
'MontaSubBandContribuicaoDeMelhorias ab
'MontaBotoes chd, 616, "mnuTabelas", "Contribuic�o de Melhorias...", True, "chdSubTabelasContribuicaoDeMelhorias"
MontaSubBandEconomicas ab
MontaBotoes chd, 614, "mnuTabelas", "Econ�micas...               ", True, "chdSubTabelasEconomicas"
MontaBotoes chd, 577, "mnuTabelas", "Gerais...                   ", True, "chdSubTabelasGerais"
MontaSubBandImobiliariasUrbanas ab
MontaBotoes chd, 608, "mnuTabelas", "Imobili�rias Urbanas...     ", True, "chdSubTabelasImobiliariasUrbanas"
MontaSubBandLogradouros ab
MontaBotoes chd, 579, "mnuTabelas", "Logradouros...           ", True, "chdSubTabelasLogradouros"

'MontaBotoes chd, 618, "mnuTabelas", "Sair...                     "

End Sub

Private Sub MontaMenuCadastros(ab As ActiveBar2, chd)
    Set chd = ab.Bands("bndFormulario").ChildBands.Add("ChildBand")
    ab.Bands("bndFormulario").ChildBands.ChildBandCaptionAlignment = ddCACenter
    'ab.Bands("bndFormulario").ChildBands.BackColor = &H8000000A
    With chd
        .Name = "chdCadastros"
        .Caption = "Cadastros"
        .GrabHandleStyle = ddGSCaption
        .flags = 127
        .Width = 12000
    End With
    MontaSubBandAcordo ab
    MontaBotoes chd, 1147, "mnuCadastros", "Acordos                    ", True, "chdSubAcordo"
    MontaBotoes chd, 452, "mnuCadastros", "Cadastramento de Processos"
    MontaBotoes chd, 450, "mnuCadastros", "Cat�logo de Assuntos"
    MontaBotoes chd, 626, "mnuCadastros", "Contadores               "
    MontaBotoes chd, 628, "mnuCadastros", "Contas Banc�rias         "
    MontaBotoes chd, 738, "mnuCadastros", "Contribui��o de Melhorias"
    MontaBotoes chd, 15, "mnuCadastros", "Contribuintes            "
    MontaBotoes chd, 1207, "mnuCadastros", "D�vida Ativa             "
    MontaBotoes chd, 737, "mnuCadastros", "Econ�mico                "
    MontaBotoes chd, 735, "mnuCadastros", "Imobili�rio Urbano       "
    MontaBotoes chd, 629, "mnuCadastros", "Isen��es e Imunidades   "
    MontaSubBandFiscalizacaoiss ab
    MontaBotoes chd, 1424, "mnuCadastros", "ISS - Fiscaliza��o         ", True, "chdSubFiscalizacaoiss"
    MontaSubBandLancamentos ab
    MontaBotoes chd, 1122, "mnuLivroCaixa", "Lan�amentos             ", True, "chdSubLancamentos"
    MontaSubBandParametros ab
    MontaBotoes chd, 1075, "mnuCadastros", "Par�metros                 ", True, "chdSubParametros"
    MontaBotoes chd, 627, "mnuCadastros", "S�cios                   "
    'MontaBotoes chd, 736, "mnuCadastros", "Imobili�rio Rural        "
    
End Sub
    

Private Sub MontaSubBandLancamentos(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubLancamentos"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
        '.Left = 15000
    End With
    
    MontaBotoes chd, 1115, "mnuSubCadastroLancamentos", "Guias                    "
    MontaBotoes chd, 1076, "mnuSubCadastroLancamentos", "IPTU                     "
    MontaBotoes chd, 1242, "mnuSubCadastroLancamentos", "ISS Constru��o           "
    MontaBotoes chd, 1190, "mnuSubCadastroLancamentos", "ISS e Taxas de Licen�as  "
    MontaBotoes chd, 1241, "mnuSubCadastroLancamentos", "Pre�o P�blico            "
    MontaBotoes chd, 1387, "mnuSubCadastroLancamentos", "Executivos Fiscais       "
    
End Sub

Private Sub MontaSubBandParametros(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubParametros"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
        '.Left = 15000
    End With

    MontaBotoes chd, 1151, "mnuSubCadastroLancamentos", "Atualiza��o de Valores          "
    MontaBotoes chd, 1247, "mnuSubCadastroLancamentos", "Divida Ativa                    "
    MontaBotoes chd, 1246, "mnuSubCadastroLancamentos", "Lan�amentos                     "
End Sub

Private Sub MontaSubBandFiscalizacaoiss(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubFiscalizacaoiss"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
        '.Left = 15000
    End With

    MontaBotoes chd, 1425, "mnuSubCadastroFiscalizacaoiss", "Notas Fiscais                   "
End Sub

Private Sub MontaSubBandAcordo(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubAcordo"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
        '.Left = 15000
    End With

    MontaBotoes chd, 1147, "mnuSubCadastroAcordo", "Acordo                                                   ", True
    MontaBotoes chd, 1406, "mnuSubCadastroAcordo", "Cancelamento de Acordo por Inadimpl�ncia", True
    
End Sub


Private Sub MontaMenuExpediente(ab As ActiveBar2, chd)
Set chd = ab.Bands("bndFormulario").ChildBands.Add("ChildBand")
ab.Bands("bndFormulario").ChildBands.ChildBandCaptionAlignment = ddCACenter
ab.Bands("bndFormulario").ChildBands.BackColor = &HE0E0E0
With chd
    .Name = "chdExpediente"
    .Caption = "Expediente"
    .GrabHandleStyle = ddGSCaption
    .flags = 127
    .Width = 12000
End With
MontaSubBandAdministracao ab
MontaBotoes chd, 631, "mnuExpediente", "Administra��o...                         ", True, "chdSubExpedienteAdministracao"
MontaSubBandAlteracaoEndNotificacao ab
MontaBotoes chd, 1260, "mnuExpediente", "Altera��o de Endere�o de Notifica��o      ", True, "chdSubAlteracaoEndNotificacao"
MontaBotoes chd, 1161, "mnuExpediente", "Atualiza��o de D�bitos                    "
MontaSubBandBaixas ab
MontaBotoes chd, 675, "mnuExpediente", "Controle de Arrecada��o...               ", True, "chdSubExpedienteBaixas"
MontaSubBandInscricaoDA ab
MontaBotoes chd, 1223, "mnuExpediente", "Inscri��o de D�vida                     ", True, "chdSubExpedienteInscricaoDA"
MontaSubBandCalculos ab
MontaBotoes chd, 663, "mnuExpediente", "Lan�amentos...                           ", True, "chdSubExpedienteCalculos"

MontaBotoes chd, 1376, "mnuExpediente", "Executivos Fiscais                         ", True, "chdSubExecutivos"
MontaSubBandExecutivos ab

End Sub

Private Sub MontaSubBandExecutivos(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubExecutivos"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
        '.Left = 15000
    End With
    
    MontaBotoes chd, 1377, "mnuSubCadastroLancamentos", "Gera��o Arquivo Distribuidor                        "
End Sub


Private Sub MontaSubBandAlteracaoEndNotificacao(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubAlteracaoEndNotificacao"
        .DockingArea = ddDAPopup
        .Caption = "Expediente"
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    
    MontaBotoes chd, 1262, "mnuSubAlteracaoEndNotificacao", "Contribuinte     "
    MontaBotoes chd, 1261, "mnuSubAlteracaoEndNotificacao", "Imobili�rio      "
        
End Sub


Private Sub MontaSubBandAdministracao(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteAdministracao"
    .DockingArea = ddDAPopup
    .Caption = "Expediente"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 633, "mnuSubExpedienteAdministracao", "Controle de Devolu��o de Documentos   "
MontaBotoes chd, 632, "mnuSubExpedienteAdministracao", "Emiss�o e Validade de Documentos      "
MontaBotoes chd, 634, "mnuSubExpedienteAdministracao", "Reavalia��o de Valores                "

End Sub
Private Sub MontaSubBandFinanceiro(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteFinanceiro"
    .DockingArea = ddDAPopup
    .Caption = "Expediente"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 636, "mnuSubExpedienteFinanceiro", "Lan�amentos em Conta Corrente                                 " '- RETIRADO 26/07/04 Rafael
'MontaSubBandTransferenciasParaDividaAtiva ab '- RETIRADO 27/07/04 Rafael
'MontaBotoes chd, 639, "mnuSubExpedienteFinanceiro", "Transfer�ncias para Divida Ativa...                           ", True, "chdSubExpedienteFinanceiroTransferenciasParaDividaAtiva" '- RETIRADO 27/07/04 Rafael
'MontaSubBandParcelamentos ab '- RETIRADO 27/07/04 Rafael
'MontaBotoes chd, 642, "mnuSubExpedienteFinanceiro", "Parcelamentos...                                              ", True, "chdSubExpedienteFinanceiroParcelamentos" '- RETIRADO 27/07/04 Rafael
'MontaBotoes chd, 646, "mnuSubExpedienteFinanceiro", "C�lculo de Acr�scimos Legais                                  " '- RETIRADO 27/07/04 Rafael
End Sub
Private Sub MontaSubBandTransferenciasParaDividaAtiva(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteFinanceiroTransferenciasParaDividaAtiva"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 640, "mnuSubExpedienteFinanceiroTransferenciasParaDividaAtiva", "D�bitos Gerados Pelo Sistema " '- RETIRADO 26/07/04 Rafael
'MontaBotoes chd, 641, "mnuSubExpedienteFinanceiroTransferenciasParaDividaAtiva", "D�bitos Gerados Manualmente" '- RETIRADO 26/07/04 Rafael

End Sub
Private Sub MontaSubBandParcelamentos(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteFinanceiroParcelamentos"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 643, "mnuSubExpedienteFinanceiroParcelamentos", "Parcelamentos de D�bitos  " '- RETIRADO 27/07/04 Rafael

End Sub

Private Sub MontaSubBandISSQNVariavel(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteCalculosISSQNVariavel"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 667, "chdSubExpedienteCalculosISSQNVariavel", "ISSQN Mensal ou Homologado" '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 668, "chdSubExpedienteCalculosISSQNVariavel", "ISSQN Arbitrado           " '- RETIRADO 28/07/04 Rafael
'MontaBotoes chd, 669, "chdSubExpedienteCalculosISSQNVariavel", "ISSQN Estimado            " '- RETIRADO 28/07/04 Rafael

End Sub

Private Sub MontaSubBandITBIUrbanoeRural(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteCalculosITBIUrbanoeRural"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 672, "chdSubExpedienteCalculosITBIUrbanoeRural", "C�lculo" '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 673, "chdSubExpedienteCalculosITBIUrbanoeRural", "Troca Propriet�rio de Im�veis" '- RETIRADO 29/07/04 Rafael
End Sub

Private Sub MontaSubBandFiscalizacao(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteFiscalizacao"
    .DockingArea = ddDAPopup
    .Caption = "Expediente"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 648, "mnuSubExpedienteFiscalizacao", "Controle de Notas Fiscais               " '- RETIRADO 27/07/04 Rafael
'MontaBotoes chd, 649, "mnuSubExpedienteFiscalizacao", "Ordens de Servi�o                       " '- RETIRADO 27/07/04 Rafael
'MontaBotoes chd, 650, "mnuSubExpedienteFiscalizacao", "Mapa de A��o Fiscal                     " '- RETIRADO 27/07/04 Rafael
'MontaBotoes chd, 651, "mnuSubExpedienteFiscalizacao", "Autos de Infra��o                       " '- RETIRADO 27/07/04 Rafael
'MontaBotoes chd, 653, "mnuSubExpedienteFiscalizacao", "Controle de Declara��o de ISSQN Vari�vel" '- RETIRADO 27/07/04 Rafael
End Sub

Private Sub MontaSubBandContencioso(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteContencioso"
    .DockingArea = ddDAPopup
    .Caption = "Expediente"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
   ' .Left = 15000
End With

'MontaSubBandAdministrativo ab '- RETIRADO 28/07/04 Rafael
'MontaBotoes chd, 655, "mnuSubExpedienteContencioso", "Administrativo...               ", True, "chdSubExpedienteContenciosoAdministrativo" '- RETIRADO 28/07/04 Rafael
'MontaSubBandJudicial ab '- RETIRADO 28/07/04 Rafael
'MontaBotoes chd, 661, "mnuSubExpedienteContencioso", "Judicial...                     ", True, "chdSubExpedienteContenciosoJudicial" '- RETIRADO 28/07/04 Rafael
'MontaBotoes chd, 3, "mnuSubExpedienteContencioso", "Rela��o de Documentos Devolvidos"
End Sub
Private Sub MontaSubBandAdministrativo(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteContenciosoAdministrativo"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 656, "mnuSubExpedienteContenciosoAdministrativo", "Suspens�o De Exig�ncias" '- RETIRADO 28/07/04 Rafael
'MontaBotoes chd, 657, "mnuSubExpedienteContenciosoAdministrativo", "Prescri��o De D�bitos" '- RETIRADO 28/07/04 Rafael
'MontaBotoes chd, 658, "mnuSubExpedienteContenciosoAdministrativo", "Cancelamento De D�bitos" '- RETIRADO 28/07/04 Rafael
'MontaBotoes chd, 659, "mnuSubExpedienteContenciosoAdministrativo", "Remiss�o De D�bitos" '- RETIRADO 28/07/04 Rafael
'MontaBotoes chd, 660, "mnuSubExpedienteContenciosoAdministrativo", "Cobran�a Extra-Judicial" '- RETIRADO 28/07/04 Rafael
End Sub
Private Sub MontaSubBandJudicial(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteContenciosoJudicial"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 662, "mnuSubExpedienteContenciosoJudicial", "Execu��o Fiscal" '- RETIRADO 28/07/04 Rafael
End Sub
Private Sub MontaSubBandCalculos(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteCalculos"
    .DockingArea = ddDAPopup
    .Caption = "Expediente"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 1204, "mnuSubExpedienteCalculos", "Pre�o P�blico - Guias             "
MontaBotoes chd, 664, "mnuSubExpedienteCalculos", "Tributos                           "
MontaBotoes chd, 1370, "mnuSubExpedienteCalculos", "Executivos Fiscais                "
MontaBotoes chd, 1437, "mnuSubExpedienteCalculos", "Cobran�a Amig�vel                  "
'MontaBotoes chd, 665, "mnuSubExpedienteCalculos", "ISSQN Fixo ou Anual               " '- RETIRADO 28/07/04 Rafael
'MontaSubBandISSQNVariavel ab '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 666, "mnuSubExpedienteCalculos", "ISSQN Vari�vel...                 ", True, "chdSubExpedienteCalculosISSQNVariavel" '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 670, "mnuSubExpedienteCalculos", "Contribui��o de Melhorias         " '- RETIRADO 29/07/04 Rafael
'MontaSubBandITBIUrbanoeRural ab '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 671, "mnuSubExpedienteCalculos", "ITBI Urbano e Rural...            ", True, "chdSubExpedienteCalculosITBIUrbanoeRural" '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 674, "mnuSubExpedienteCalculos", "Receita Diversas                  " '- RETIRADO 29/07/04 Rafael
End Sub

Private Sub MontaSubBandInscricaoDA(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteInscricaoDA"
    .DockingArea = ddDAPopup
    .Caption = "Expediente"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

MontaBotoes chd, 1295, "mnuSubExpedienteInscricaoDA", "Manual                            "
MontaBotoes chd, 1296, "mnuSubExpedienteInscricaoDA", "Por Composi��o e Exerc�cio        "

End Sub

Private Sub MontaSubBandBaixas(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubExpedienteBaixas"
    .DockingArea = ddDAPopup
    .Caption = "Expediente"
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaSubBandCobrancaBancaria ab '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 676, "mnuSubExpedienteBaixas", "Cobran�a Banc�ria", True, "chdSubConbrancaBancaria" '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 680, "mnuSubExpedienteBaixas", "Arrecada��o Manual " '- RETIRADO 29/07/04 Rafael
MontaBotoes chd, 285, "mnuSubExpedienteBaixas", "Arrecada��o da Receita "
MontaBotoes chd, 1144, "mnuSubExpedienteBaixas", "Baixa Manual"
MontaBotoes chd, 1419, "mnuSubExpedienteBaixas", "Gerar D�bito Autom�tico"
MontaBotoes chd, 1110, "mnuSubExpedienteBaixas", "Movimento Banc�rio"
MontaBotoes chd, 1132, "mnuSubExpedienteBaixas", "Processamento De Baixa"
MontaBotoes chd, 1154, "mnuSubExpedienteBaixas", "Receber Movimento Banc�rio"
MontaBotoes chd, 1108, "mnuSubExpedienteBaixas", "Resumo Banc�rio "


End Sub

Private Sub MontaSubBandCobrancaBancaria(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubConbrancaBancaria"
    .DockingArea = ddDAPopup
    .Caption = ""
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 677, "mnuConbrancaBancaria", "Confec��o de Layout" '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 678, "mnuConbrancaBancaria", "Baixa Autom�tica     " '- RETIRADO 29/07/04 Rafael
End Sub

Private Sub MontaMenuGraficos(ab As ActiveBar2, chd)
Set chd = ab.Bands("bndFormulario").ChildBands.Add("ChildBand")
ab.Bands("bndFormulario").ChildBands.ChildBandCaptionAlignment = ddCACenter
'ab.Bands("bndFormulario").ChildBands.BackColor = &H8000000A
With chd
    .Name = "chdGraficos"
    .Caption = "Gr�ficos"
    .GrabHandleStyle = ddGSCaption
    .flags = 127
    .Width = 12000
End With


End Sub

Private Sub MontaMenuRelatorios(ab As ActiveBar2, chd)
Set chd = ab.Bands("bndFormulario").ChildBands.Add("ChildBand")
ab.Bands("bndFormulario").ChildBands.ChildBandCaptionAlignment = ddCACenter
'ab.Bands("bndFormulario").ChildBands.BackColor = &H8000000A
With chd
    .Name = "chdRelatorios"
    .Caption = "Relat�rios"
    .GrabHandleStyle = ddGSCaption
    .flags = 127
    .Width = 12000
End With

MontaSubBandSegundasVias ab
MontaBotoes chd, 1368, "mnuRelatorios", "2� Vias                               ", True, "chdSubSegundasVias"
MontaBotoes chd, 1361, "mnuRelatorios", "Altera��es Cadastrais Econ�micas"
MontaBotoes chd, 1181, "mnuRelatorios", "Atividades e Contribuintes por Logradouro"
MontaSubBandBaixaAnaliseReceita ab
MontaBotoes chd, 1131, "mnuRelatorios", "Baixa e An�lise da Receita            ", True, "chdSubBaixaAnaliseReceita"
MontaBotoes chd, 1196, "mnuRelatorios", "Contribuintes por Atividades"
MontaBotoes chd, 1441, "mnuRelatorios", "Devedores por Faixa de Valores"

MontaSubBandExecutivosFiscais ab
MontaBotoes chd, 1388, "mnuRelatorios", "Executivos Fiscais                    ", True, "chdSubExecutivosFiscais"

MontaSubBandFichasCadastrais ab
MontaBotoes chd, 1378, "mnuRelatorios", "Fichas Cadastrais", True, "chdSubBandFichasCadastrais"

MontaSubBandFichasLancamentos ab
MontaBotoes chd, 1379, "mnuRelatorios", "Fichas de Lan�amentos", True, "chdSubBandFichasLAncamentos"

MontaBotoes chd, 1169, "mnuRelatorios", "Inscri��o Setor e Quadra"

MontaSubBandISS ab
MontaBotoes chd, 1381, "mnuRelatorios", "ISS", True, "chdSubBandISS"

MontaBotoes chd, 1274, "mnuRelatorios", "Isen��o Imunidade"

MontaSubBandLancamentosIPTU ab
MontaBotoes chd, 1265, "mnuRelatorios", "Lan�amentos", True, "chdSubLancamentosIPTU"
MontaBotoes chd, 1228, "mnuRelatorios", "Livro de D�vida Ativa"
MontaBotoes chd, 1359, "mnuRelatorios", "Ocorr�ncias do Econ�mico"

MontaSubBandPagamentos ab
MontaBotoes chd, 1380, "mnuRelatorios", "Pagamentos", True, "chdSubBandPagamentos"
MontaBotoes chd, 1392, "mnuRelatorios", "Posi��o de Lan�amentos - Pagamentos"
MontaBotoes chd, 1178, "mnuRelatorios", "Quantidade de Contribuintes por Atividade"
MontaBotoes chd, 1395, "mnuRelatorios", "Receita de Composi��o - Lan�ado Pago      "
MontaBotoes chd, 1177, "mnuRelatorios", "Rol de Atividades"
MontaBotoes chd, 1174, "mnuRelatorios", "Rol de Logradouros"

MontaSubBandSaldoDividaAtiva ab
MontaBotoes chd, 1239, "mnuRelatorios", "Saldo de D�vida Ativa", True, "chdSubSaldoDividaAtiva"
MontaBotoes chd, 1195, "mnuRelatorios", "Taxas de Licen�a"

'MontaSubBandRelCadastroTecnico ab '- RETIRADO 30/07/04 Rafael
'MontaBotoes chd, 4, "mnuRelatorios", "Cadastro T�cnico Municipal...          ", True, "chdSubBandRelCadastroTecnico" '- (intCodigo = 685) RETIRADO 30/07/04 Rafael

'MontaSubBandRelControleArrecadacao ab '- RETIRADO 02/08/04 Rafael
'MontaBotoes chd, 5, "mnuRelatorios", "Controle da Arrecada��o...             ", True, "chdSubBandRelControleArrecadacao" '- (intCodigo = 701) RETIRADO 02/08/04 Rafael
'MontaSubBandRelContaCorrenteFiscal ab '- RETIRADO 03/08/04 Rafael
'MontaBotoes chd, 6, "mnuRelatorios", "Conta Corrente Fiscal...               ", True, "chdSubBandRelContaCorrenteFiscal" '- (intCodigo = 766) RETIRADO 03/08/04 Rafael
'MontaSubBandRelCobranca ab '- RETIRADO 03/08/04 Rafael
'MontaBotoes chd, 7, "mnuRelatorios", "Cobran�a...                            ", True, "chdSubBandRelCobranca" '- (intCodigo = 773) RETIRADO 03/08/04 Rafael
'MontaSubBandRelFiscalizacaoContencioso ab '- RETIRADO 03/08/04 Rafael
'MontaBotoes chd, 8, "mnuRelatorios", "Fiscaliza��o / Contencioso...          ", True, "chdSubBandRelFiscalizacaoContencioso" '- (intCodigo = 776) RETIRADO 03/08/04 Rafael
'MontaSubBandRelDividaAtiva ab '- RETIRADO 04/08/04 Rafael
'MontaBotoes chd, 9, "mnuRelatorios", "D�vida Ativa...                       ", True, "chdSubBandRelDividaAtiva" '- (intCodigo = 781) RETIRADO 04/08/04 Rafael
'MontaBotoes chd, 10, "mnuRelatorios", "Contribuintes Duplicados", False '- (N�o tem no banco) RETIRADO 04/08/04 Rafael

End Sub

Private Sub MontaSubBandISS(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubBandISS"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    
    MontaBotoes chd, 1382, "mnuSubBandISS", "Notas Fiscais - Emiss�o"
    
End Sub

Private Sub MontaSubBandPagamentos(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubBandPagamentos"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    
    MontaBotoes chd, 1267, "mnuSubBandPagamentos", "Pagamentos por Aviso"
    MontaBotoes chd, 1259, "mnuSubBandPagamentos", "Relat�rio de Pagamentos"
End Sub


Private Sub MontaSubBandFichasLancamentos(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubBandFichasLAncamentos"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    
    MontaBotoes chd, 1222, "mnuSubBandFichasLAncamentos", "Ficha de Lan�amento Imobili�rio"
    MontaBotoes chd, 1302, "mnuSubBandFichasLAncamentos", "Ficha de Lan�amento ISS Constru��o"
   
End Sub

Private Sub MontaSubBandFichasCadastrais(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubBandFichasCadastrais"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    
    MontaBotoes chd, 1152, "mnuSubBandFichasCadastrais", "Ficha de Cadastro Imobili�rio"
    MontaBotoes chd, 1358, "mnuSubBandFichasCadastrais", "Ficha de Cadastro Econ�mico"
    
End Sub

Private Sub MontaSubBandLancamentosIPTU(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubLancamentosIPTU"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    MontaBotoes chd, 1284, "mnuSubRelLancamentos", "Comparativo IPTU"
    MontaBotoes chd, 1266, "mnuSubRelLancamentos", "Totaliza��o do IPTU"
    
    MontaSubBandRolLancamentos ab
    MontaBotoes chd, 1430, "mnuRelatorios", "Rol de Lan�amentos", True, "chdSubRolLancamentos"
    
End Sub

Private Sub MontaSubBandRolLancamentos(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubRolLancamentos"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    MontaBotoes chd, 1431, "mnuSubRolLancamentos", "Econ�mico"
End Sub

Private Sub MontaSubBandArquivosGrafica(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubArquivosGrafica"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    MontaBotoes chd, 1439, "mnuSubArquivosGrafica", "IPTU"
    MontaBotoes chd, 1440, "mnuSubArquivosGrafica", "Cobran�a Amig�vel"
End Sub

Private Sub MontaSubBandSegundasVias(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubSegundasVias"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    
    MontaBotoes chd, 1203, "mnuSubSegundasVias", "Acordo               "
    MontaBotoes chd, 1408, "mnuSubSegundasVias", "Acordo (Parcelas Atualizadas)"
    MontaBotoes chd, 1197, "mnuSubSegundasVias", "IPTU                 "
    MontaBotoes chd, 1240, "mnuSubSegundasVias", "ISS                  "
    MontaBotoes chd, 1243, "mnuSubSegundasVias", "ISS Constru��o       "
    MontaBotoes chd, 1369, "mnuSubSegundasVias", "ISS Vari�vel         "
    MontaBotoes chd, 1268, "mnuSubSegundasVias", "Pre�o P�blico        "
    
    
End Sub

Private Sub MontaSubBandExecutivosFiscais(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubExecutivosFiscais"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    
    MontaBotoes chd, 1389, "mnuSubExecutivosFiscais", "Peti��o                 "
    MontaBotoes chd, 1394, "mnuSubExecutivosFiscais", "Certid�o de D�vida Ativa"
    
End Sub

Private Sub MontaSubBandBaixaAnaliseReceita(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubBaixaAnaliseReceita"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
        '.Left = 15000
    End With
    
    MontaBotoes chd, 1133, "mnuSubRelBaixaAnaliseReceita", "Baixas              "
    MontaBotoes chd, 1135, "mnuSubRelBaixaAnaliseReceita", "Cr�ticas            "
    MontaBotoes chd, 1131, "mnuSubRelBaixaAnaliseReceita", "Diverg�ncias        "
    MontaBotoes chd, 304, "mnuSubRelBaixaAnaliseReceita", "Receita Arrecadada  "
    MontaBotoes chd, 1418, "mnuSubRelBaixaAnaliseReceita", "Movimento Banc�rio  "

End Sub


Private Sub MontaSubBandRelDividaAtiva(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubBandRelDividaAtiva"
    .DockingArea = ddDAPopup
    .Caption = ""
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 782, "mnuRelDividaAtiva", "Peti��es de Ajuizamento"
'MontaBotoes chd, 783, "mnuRelDividaAtiva", "Livro da D�vida Ativa" '- RETIRADO 03/08/04 Rafael
'MontaBotoes chd, 784, "mnuRelDividaAtiva", "Rela��o de Adimpl�ncia em D�vida Ativa " '- RETIRADO 04/08/04 Rafael
'MontaBotoes chd, 785, "mnuRelDividaAtiva", "Rela��o de Inadimpl�ncia em D�vida Ativa" '- RETIRADO 04/08/04 Rafael

End Sub

Private Sub MontaSubBandRelFiscalizacaoContencioso(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubBandRelFiscalizacaoContencioso"
    .DockingArea = ddDAPopup
    .Caption = ""
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 777, "mnuRelFiscalizacaoContencioso", "Posi��o de Alvar�s" '- RETIRADO 03/08/04 Rafael
'MontaBotoes chd, 778, "mnuRelFiscalizacaoContencioso", "Contesta��es Apresentadas"
'MontaBotoes chd, 779, "mnuRelFiscalizacaoContencioso", "Decis�es/Pareceres/Despachos"
'MontaBotoes chd, 780, "mnuRelFiscalizacaoContencioso", "Documentos Diversos" '- RETIRADO 03/08/04 Rafael

End Sub

Private Sub MontaSubBandRelCobranca(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubBandRelCobranca"
    .DockingArea = ddDAPopup
    .Caption = ""
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 774, "mnuRelCobranca", "Rela��o de Documentos Devolvidos " '- RETIRADO 03/08/04 Rafael
'MontaBotoes chd, 775, "mnuRelCobranca", "Documentos Diversos              " '- RETIRADO 03/08/04 Rafael

End Sub

Private Sub MontaSubBandRelContaCorrenteFiscal(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubBandRelContaCorrenteFiscal"
    .DockingArea = ddDAPopup
    .Caption = ""
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 768, "mnuRelContaCorrenteFiscal", "Rela��o de Inadimpl�ncia Anal�tico" '- RETIRADO 03/08/04 Rafael
'MontaBotoes chd, 769, "mnuRelContaCorrenteFiscal", "Rela��o de Inadimpl�ncia Sint�tico" '- RETIRADO 03/08/04 Rafael
'MontaBotoes chd, 768, "mnuRelContaCorrenteFiscal", "Anal�tico dos Maiores Devedores"
'MontaBotoes chd, 769, "mnuRelContaCorrenteFiscal", "D�bitos Baixados e Pagamentos Registrados "
'MontaBotoes chd, 770, "mnuRelContaCorrenteFiscal", "Maiores D�bitos com Suspens�o de Exig�ncia"
'MontaBotoes chd, 771, "mnuRelContaCorrenteFiscal", "D�bitos n�o Inscritos em Divida Ativa " '- RETIRADO 03/08/04 Rafael
'MontaBotoes chd, 772, "mnuRelContaCorrenteFiscal", "Documentos Diversos" '- RETIRADO 03/08/04 Rafael


End Sub

Private Sub MontaSubBandRelControleArrecadacao(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubBandRelControleArrecadacao"
    .DockingArea = ddDAPopup
    .Caption = ""
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 763, "mnuRelControleArrecadacao", "Registro de Documentos n�o conciliados "
'MontaBotoes chd, 764, "mnuRelControleArrecadacao", "Receita di�ria por Tipo, Valor e Agente"
'MontaBotoes chd, 765, "mnuRelControleArrecadacao", "Documentos Diversos                    " '- RETIRADO 30/07/04 Rafael

End Sub


Private Sub MontaSubBandRelCadastroTecnico(ab As ActiveBar2)
Dim chd

Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdSubBandRelCadastroTecnico"
    .DockingArea = ddDAPopup
    .Caption = ""
    .GrabHandleStyle = ddGSNone
    .Type = ddBTNormal
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

'MontaBotoes chd, 686, "mnuRelCadastroTecnico", "Conformidade com Inclus�o/Altera��o/Exclus�o                  "
'MontaBotoes chd, 687, "mnuRelCadastroTecnico", "Beneficiados com Imunidade/Isen��o/N�o Incid�ncia             " '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 688, "mnuRelCadastroTecnico", "Inconsist�ncias Imobili�rias                                  " '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 689, "mnuRelCadastroTecnico", "Rela��o de Contadores por Empresa                             " '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 690, "mnuRelCadastroTecnico", "Contadores e Arrecada��o no Per�odo                           " '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 691, "mnuRelCadastroTecnico", "Inscritos Ativos/Inativos/Baixados                            " '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 692, "mnuRelCadastroTecnico", "Contribuintes em Contencioso Administrativo                   " '- RETIRADO 29/07/04 Rafael
'MontaBotoes chd, 693, "mnuRelCadastroTecnico", "Editais / Notifica��es de Lan�amento                          " '- RETIRADO 30/07/04 Rafael
'MontaBotoes chd, 695, "mnuRelCadastroTecnico", "Demonstrativo de Arrecada��o de ISSQN por Atividade Econ�mica " '- RETIRADO 30/07/04 Rafael
'MontaBotoes chd, 696, "mnuRelCadastroTecnico", "Quantidade de Lan�amento, Valor e Tipo                        " '- RETIRADO 30/07/04 Rafael
'MontaBotoes chd, 698, "mnuRelCadastroTecnico", "Extrato Individualizado de Lan�amento                         " '- RETIRADO 30/07/04 Rafael
'MontaBotoes chd, 699, "mnuRelCadastroTecnico", "Rela��o das Parcelas Lan�adas                                 " '- RETIRADO 30/07/04 Rafael
'MontaBotoes chd, 697, "mnuRelCadastroTecnico", "Rela��o das Parcelas Arrecadadas                              " '- RETIRADO 30/07/04 Rafael
'MontaBotoes chd, 682, "mnuRelCadastroTecnico", "Rela��o de Diferen�a nos Pagamentos                           " '- RETIRADO 30/07/04 Rafael
'MontaBotoes chd, 700, "mnuRelCadastroTecnico", "Documentos Diversos                                           " '- RETIRADO 30/07/04 Rafael

End Sub

Private Sub MontaMenuFerramentas(ab As ActiveBar2, chd)
Set chd = ab.Bands("bndFormulario").ChildBands.Add("ChildBand")
ab.Bands("bndFormulario").ChildBands.ChildBandCaptionAlignment = ddCACenter
'ab.Bands("bndFormulario").ChildBands.BackColor = &H8000000A
With chd
    .Name = "chdFerramentas"
    .Caption = "Ferramentas"
    .GrabHandleStyle = ddGSCaption
    .flags = 127
    .Width = 12000
End With
    
    
    MontaBotoes chd, 2, "mnuSubFerramentas", "Alterar senha"
    MontaBotoes chd, 3, "mnuSubFerramentas", "Auto Numera��o"
    
    'Nino
    MontaSubWordModelosGravados ab
    'Nino
    MontaBotoes chd, 1107, "mnuSubFerramentas", "Editor de Modelos de Documentos", True, "chdArqSubWordModelosGravados"
    MontaBotoes chd, 1, "mnuSubFerramentas", "Op��es"
    MontaBotoes chd, 1291, "mnuSubFerramentas", "Receber Lan�amento Externo"
    MontaBotoes chd, 1310, "mnuSubFerramentas", "Taxas de Licen�a - Arquivo Gr�fica"
    
    MontaSubBandArquivosGrafica ab
    MontaBotoes chd, 1407, "mnuSubFerramentas", "Gerar Arquivo para Gr�fica", True, "chdSubArquivosGrafica"
    
    MontaBotoes chd, 1417, "mnuSubFerramentas", "Gerar Arquivo para Internet - IPTU"

End Sub
Private Sub MontaMenuAjuda(ab As ActiveBar2, chd)
Set chd = ab.Bands("bndFormulario").ChildBands.Add("ChildBand")
ab.Bands("bndFormulario").ChildBands.ChildBandCaptionAlignment = ddCACenter
'ab.Bands("bndFormulario").ChildBands.BackColor = &H8000000A
With chd
    .Name = "chdAjuda"
    .Caption = "Ajuda"
    .GrabHandleStyle = ddGSCaption
    .flags = 127
    .Width = 12000
End With
End Sub

Public Sub CreateToolsLateral(ab As ActiveBar2)
    Dim bnd As ActiveBar2LibraryCtl.Band
    Dim chd
    Dim iCat As Integer
    Dim keys(0) As New ShortCut
    
    Set bnd = ab.Bands.Add("bndFormulario"): bnd.Type = ddBTNormal
    bnd.Caption = ""
    With bnd
        .DockingArea = ddDALeft
        .GrabHandleStyle = ddGSCaption
        .Type = ddBTNormal
        .ChildBandStyle = ddCBSSlidingTabs
        .AutoSizeForms = True
        .DisplayMoreToolsButton = False
        .Caption = "Menu Principal"
        .DockedHorzHeight = 1995
        .DockedHorzMinWidth = 1200
        .DockedHorzWidth = 7740
        .DockedVertHeight = 3270
        .DockedVertMinWidth = 1440
        .DockedVertWidth = 1995
        .Height = 4320
        .Width = 2100
        .flags = 14462
    End With
    
    'Monta o menu Principal
    MontaMenuPrincipal ab, chd
    
    'Monta o menu Tabelas
    MontaMenuTabelas ab, chd
    
    'Monta o menu Cadastros
    MontaMenuCadastros ab, chd
    
    'Monta o menu Expediente
    MontaMenuExpediente ab, chd
    
    'Nino
    MontaMenuDocumentos ab, chd
    
   'Monta o menu Relat�rios
    MontaMenuRelatorios ab, chd
    
    'Monta o menu Ferramentas
    MontaMenuFerramentas ab, chd
    
    'Monta o menu Ajuda
    'MontaMenuAjuda ab, chd
    
    ab.RecalcLayout
    ab.Refresh
End Sub

Public Sub CarregaIconeEspecial(actBarra As ActiveBar2, img_ListaIcones As ImageList)
    Dim Tool As ActiveBar2LibraryCtl.Tool
    Dim acbBandeira As ActiveBar2LibraryCtl.Band
    Set Tool = MDIMenu.actBarra.Tools.Add(207, "miESeparador")
    Tool.ControlType = ddTTSeparator
    Tool.Category = gstrBtnArquivo
    
    Set Tool = actBarra.Tools.Add(208, gstrCalcularReajuste)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(gstrCalcularReajuste).Picture
    Tool.ToolTipText = "Calcular reajuste"
    
'    Set Tool = actBarra.Tools.Add(209, gstrBrasao)
'    Tool.Category = gstrBtnArquivo
'    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(20).Picture
'    Tool.ToolTipText = "Cadastra ou altera o Bras�o"
'
'    Set Tool = actBarra.Tools.Add(210, gstrLogotipo)
'    Tool.Category = gstrBtnArquivo
'    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(21).Picture
'    Tool.ToolTipText = "Cadastra ou altera o Logotipo"
    
    Set Tool = actBarra.Tools.Add(211, gstrLerArquivo)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(gstrLerArquivo).Picture
    Tool.ToolTipText = "Ler arquivo"

    Set Tool = actBarra.Tools.Add(2112, gstrIncluirItem)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrIncluirItem)).Picture
    Tool.ToolTipText = "Incluir ou alterar item na lista"

    Set Tool = actBarra.Tools.Add(2113, gstrExcluirItem)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrExcluirItem)).Picture
    Tool.ToolTipText = "Excluir item da lista"
    
    Set Tool = actBarra.Tools.Add(2114, gstrProcessamentoBaixa)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrProcessamentoBaixa)).Picture
    Tool.ToolTipText = "Baixar Movimento"

    Set Tool = actBarra.Tools.Add(2115, gstrImprimirGuia)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrImprimirGuia)).Picture
    Tool.ToolTipText = "Imprimir Guia"
    
    Set Tool = actBarra.Tools.Add(2116, gstrParcelamentoDebitoAtualizado)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrParcelamentoDebitoAtualizado)).Picture
    Tool.ToolTipText = "Parcelamento do D�bito Atualizado"
    
    Set Tool = actBarra.Tools.Add(2117, gstrGuiaDeAcordo)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrGuiaDeAcordo)).Picture
    Tool.ToolTipText = "Imprimir Guia de Acordo"
    
    Set Tool = actBarra.Tools.Add(2118, gstrGuiaCertidaoNegativa)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrGuiaCertidaoNegativa)).Picture
    Tool.ToolTipText = "Imprimir Certid�o Negativa"
    
    Set Tool = actBarra.Tools.Add(2119, gstrGuiaCertidaoPositiva)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrGuiaCertidaoPositiva)).Picture
    Tool.ToolTipText = "Imprimir Certid�o Positiva"
    
    Set Tool = actBarra.Tools.Add(2120, gstrGuiaRelacaoDeDebitos)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrGuiaRelacaoDeDebitos)).Picture
    Tool.ToolTipText = "Imprimir Rela��o de D�bitos"
    
    Set Tool = actBarra.Tools.Add(2121, gstrCancelarReativar)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrCancelarReativar)).Picture
    Tool.ToolTipText = ""
    
    Set Tool = actBarra.Tools.Add(2122, gstrGuiaCertidaoDividaAtiva)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrGuiaCertidaoDividaAtiva)).Picture
    Tool.ToolTipText = "Certid�o de D�vida Ativa"
    
    Set Tool = actBarra.Tools.Add(2123, gstrGuiaCertidaoPositivaEfeitoNegativa)
    Tool.Category = gstrBtnArquivo
    Tool.SetPicture ddITNormal, img_ListaIcones.ListImages(UCase(gstrGuiaCertidaoPositivaEfeitoNegativa)).Picture
    Tool.ToolTipText = "Imprimir Certid�o Positiva com Efeito Negativo"
    
    Set acbBandeira = actBarra.Bands(gstrBtnArquivo)
    With acbBandeira.Tools
        .Insert .Count, actBarra.Tools("miESeparador")
        actBarra.Tools("miESeparador").ControlType = ddTTSeparator
                
        .Insert .Count, actBarra.Tools(gstrCalcularReajuste)
        actBarra.Tools(gstrCalcularReajuste).ControlType = ddTTButton
        actBarra.Tools(gstrCalcularReajuste).Enabled = False
        
        .Insert .Count, actBarra.Tools(gstrLerArquivo)
        actBarra.Tools(gstrLerArquivo).ControlType = ddTTButton
        actBarra.Tools(gstrLerArquivo).Enabled = False
        
        .Insert .Count, actBarra.Tools(gstrIncluirItem)
        actBarra.Tools(gstrIncluirItem).ControlType = ddTTButton
        actBarra.Tools(gstrIncluirItem).Enabled = False
    
        .Insert .Count, actBarra.Tools(gstrExcluirItem)
        actBarra.Tools(gstrExcluirItem).ControlType = ddTTButton
        actBarra.Tools(gstrExcluirItem).Enabled = False
        
        .Insert .Count, actBarra.Tools(gstrProcessamentoBaixa)
        actBarra.Tools(gstrProcessamentoBaixa).ControlType = ddTTButton
        actBarra.Tools(gstrProcessamentoBaixa).Enabled = False
        
        .Insert .Count, actBarra.Tools(gstrImprimirGuia)
        actBarra.Tools(gstrImprimirGuia).ControlType = ddTTButton
        actBarra.Tools(gstrImprimirGuia).Enabled = False

        .Insert .Count, actBarra.Tools(gstrParcelamentoDebitoAtualizado)
        actBarra.Tools(gstrParcelamentoDebitoAtualizado).ControlType = ddTTButton
        actBarra.Tools(gstrParcelamentoDebitoAtualizado).Enabled = False

        .Insert .Count, actBarra.Tools(gstrGuiaDeAcordo)
        actBarra.Tools(gstrGuiaDeAcordo).ControlType = ddTTButton
        actBarra.Tools(gstrGuiaDeAcordo).Enabled = False

        .Insert .Count, actBarra.Tools(gstrGuiaCertidaoNegativa)
        actBarra.Tools(gstrGuiaCertidaoNegativa).ControlType = ddTTButton
        actBarra.Tools(gstrGuiaCertidaoNegativa).Enabled = False

        .Insert .Count, actBarra.Tools(gstrGuiaCertidaoPositiva)
        actBarra.Tools(gstrGuiaCertidaoPositiva).ControlType = ddTTButton
        actBarra.Tools(gstrGuiaCertidaoPositiva).Enabled = False

        .Insert .Count, actBarra.Tools(gstrGuiaRelacaoDeDebitos)
        actBarra.Tools(gstrGuiaRelacaoDeDebitos).ControlType = ddTTButton
        actBarra.Tools(gstrGuiaRelacaoDeDebitos).Enabled = False
        
        .Insert .Count, actBarra.Tools(gstrGuiaCertidaoDividaAtiva)
        actBarra.Tools(gstrGuiaCertidaoDividaAtiva).ControlType = ddTTButton
        actBarra.Tools(gstrGuiaCertidaoDividaAtiva).Enabled = False

        .Insert .Count, actBarra.Tools(gstrCancelarReativar)
        actBarra.Tools(gstrCancelarReativar).ControlType = ddTTButton
        actBarra.Tools(gstrCancelarReativar).Enabled = False
        
        .Insert .Count, actBarra.Tools(gstrGuiaCertidaoPositivaEfeitoNegativa)
        actBarra.Tools(gstrGuiaCertidaoPositivaEfeitoNegativa).ControlType = ddTTButton
        actBarra.Tools(gstrGuiaCertidaoPositivaEfeitoNegativa).Enabled = False
        
    End With
    MDIMenu.actBarra.RecalcLayout
    MDIMenu.actBarra.Refresh
End Sub

Public Sub Call_HtmlHelp(lngContextID)
  On Error GoTo Exit_HtmlHelp
  
  Dim strHelp As String
  ' Rotina que chama um t�pico espec�fico no HTML Help
  
  Select Case lngContextID
    Case 1
      strHelp = "Tributario.htm"
     Case 16
      strHelp = "Tabelas\gerais\agentes arrecadadores\Ag�ncias.htm"
     Case 17
      strHelp = "Tabelas\gerais\agentes arrecadadores\Bancos.htm"
     Case 41
      strHelp = "Tabelas\gerais\F�rmulas de c�lculos.htm"
     Case 581
      strHelp = "Tabelas\logradouros\Bairros.htm"
     Case 585
      strHelp = "Tabelas\logradouros\Distritos fiscais.htm"
     Case 584
      strHelp = "Tabelas\logradouros\Logradouros.htm"
     Case 53
      strHelp = "Tabelas\logradouros\Municipio.htm"
     Case 586
      strHelp = "Tabelas\logradouros\Setores fiscais.htm"
     Case 582
      strHelp = "Tabelas\logradouros\Tipos de logradouros.htm"
     Case 583
      strHelp = "Tabelas\logradouros\T�tulos de logradouros.htm"
     Case 26
      strHelp = "Tabelas\gerais\�ndices Econ�micos\�ndices econ�micos.htm"
     Case 40
      strHelp = "Tabelas\gerais\Ocorr�ncias.htm"
     Case 27
      strHelp = "Tabelas\gerais\planta de valores\Faixa de valores.htm"
     Case 28
      strHelp = "Tabelas\gerais\planta de valores\valores.htm"
     Case 23
      strHelp = "Tabelas\gerais\receita do municipio\Receitas.htm"
     Case 22
      strHelp = "Tabelas\gerais\receita do municipio\Composi��o das receitas.htm"
     Case 42
      strHelp = "Tabelas\gerais\Caracter�sticas de boletins.htm"
     Case 43
      strHelp = "Tabelas\gerais\Dias n�o �teis.htm"
     Case 48
      strHelp = "Tabelas\gerais\Vencimento de parcelas.htm"
     Case 49
      strHelp = "Tabelas\Gerais\Tipos de comunica��o.htm"
     Case 46
      strHelp = "Tabelas\gerais\Par�metros espec�ficos.htm"
     Case 31
      strHelp = "Tabelas\gerais\textos\Mensagens.htm"
     Case 32
      strHelp = "Tabelas\gerais\textos\textos1.htm"
     Case 8
      strHelp = "Tabelas\gerais\Unidades de medid1.htm"
     Case 44
      strHelp = "Tabelas\gerais\Documentos emitidos.htm"
     Case 45
      strHelp = "Tabelas\gerais\Fiscais.htm"
     Case 39
      strHelp = "Tabelas\Imobiliarias Urbanas\Tipos de �rea.htm"
     Case 38
      strHelp = "Tabelas\Imobiliarias Urbanas\Tipos de testada.htm"
     Case 36
      strHelp = "Tabelas\Imobiliarias Urbanas\Melhoramentos p�blicos.htm"
     Case 37
      strHelp = "Tabelas\Imobiliarias Urbanas\Se��es de logradouros.htm"
     Case 35
      strHelp = "Tabelas\Imobiliarias Urbanas\Fatores de corre��es.htm"
     Case 137
      strHelp = "Tabelas\Economicas\Atividades economicas.htm"
     Case 33
      strHelp = "Tabelas\Contribuicao Melhorias\editais.htm"
    Case 15
      strHelp = "cadastros\Contribuintes.htm"
    Case 735
      strHelp = "cadastros\Imobili�rio  urbano.htm"
    Case 117
      strHelp = "cadastros\Imobili�rio  rural.htm"
    Case 6
      strHelp = "cadastros\Econ�mico.htm"
    Case 108
      strHelp = "cadastros\Contribui��es de melhoria.htm"
    Case 5
      strHelp = "cadastros\D�vida ativa.htm"
    Case 50
      strHelp = "cadastros\Contadores.htm"
    Case 10
      strHelp = "cadastros\S�cios.htm"
    Case 107
      strHelp = "cadastros\Contas banc�rias.htm"
    Case 9
      strHelp = "cadastros\Isen��es e imunidade.htm"
    Case 633
      strHelp = "expediente\administracao\Controle de devolu��o de documentos.htm"
    Case 632
      strHelp = "expediente\administracao\Emiss�o e validade de documentos.htm"
    Case 634
      strHelp = "expediente\administracao\Reavaliacao de valores.htm"
    'Case 636 '- RETIRADO 26/07/04 Rafael
    '  strHelp = "expediente\conta\Lan�amento em conta corrente.htm"
    'Case 646 '- RETIRADO 27/07/04 Rafael
    '  strHelp = "expediente\conta\Calculo de Acrescimos Legais.htm"
    'Case 640 '- RETIRADO 26/07/04 Rafael
    '  strHelp = "Expediente\Conta\Debitos Gerados pelo Sistema.htm"
    'Case 641 '- RETIRADO 26/07/04 Rafael
    '  strHelp = "Expediente\Conta\Debitos Gerados Manualmente.htm"
    'Case 643 '- RETIRADO 27/07/04 Rafael
    '  strHelp = "Expediente\Conta\Parcelamentos de Debitos.htm"
    'Case 648 '- RETIRADO 27/07/04 Rafael
    '    strHelp = "Expediente\Fiscalizacao\Controle de notas fiscais.htm"
    'Case 649 '- RETIRADO 27/07/04 Rafael
    '     strHelp = "Expediente\Fiscalizacao\Ordens de Servicos.htm"
    'Case 650 '- RETIRADO 27/07/04 Rafael
    '    strHelp = "Expediente\Fiscalizacao\Mapa de Acao Fiscal.htm"
    'Case 651 '- RETIRADO 27/07/04 Rafael
    '     strHelp = "Expediente\Fiscalizacao\Autos de Infracao.htm"
    'Case 653 '- RETIRADO 27/07/04 Rafael
    '     strHelp = "Expediente\Fiscalizacao\Controle de Declaracao.htm"
    'Case 656 '- RETIRADO 28/07/04 Rafael
    '    strHelp = "Expediente\Cobranca\Menu Administrativo\Suspensao de Exigencias.htm"
    'Case 657 '- RETIRADO 28/07/04 Rafael
    '    strHelp = "Expediente\Cobranca\Menu Administrativo\Prescricao de Debitos.htm"
    'Case 658 '- RETIRADO 28/07/04 Rafael
    '    strHelp = "Expediente\Cobranca\Menu Administrativo\Cancelamento de Debitos.htm"
    'Case 659 '- RETIRADO 28/07/04 Rafael
    '    strHelp = "Expediente\Cobranca\Menu Administrativo\Remissao de Debitos.htm"
    'Case 660 '- RETIRADO 28/07/04 Rafael
    '    strHelp = "Expediente\Cobranca\Menu Administrativo\Cobranca extra judicial.htm"
    'Case 662 '- RETIRADO 28/07/04 Rafael
    '    strHelp = "Expediente\Cobranca\Execucao Fiscal.htm"
    Case 664
        strHelp = "Expediente\Lancamentos\IPTU.htm"
    'Case 665 '- RETIRADO 28/07/04 Rafael
    '    strHelp = "Expediente\Lancamentos\ISSQN fixo ou anual.htm"
    'Case 670 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Expediente\Lancamentos\contribuicao de melhorias.htm"
    'Case 674 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Expediente\Lancamentos\Receitas Diversas.htm"
    'Case 667 '- RETIRADO 29/07/04 Rafael
    '     strHelp = "Expediente\Lancamentos\ISSQN mensal  ou homologado.htm"
    'Case 668 '- RETIRADO 28/07/04 Rafael
    '     strHelp = "Expediente\Lancamentos\issqn arbitrado.htm"
    'Case 669 '- RETIRADO 28/07/04 Rafael
    '     strHelp = "Expediente\Lancamentos\issqn estimado.htm"
    'Case 672 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Expediente\Lancamentos\Calculo.htm"
    'Case 673 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Expediente\Lancamentos\Troca de Proprietario do imovel.htm"
    'Case 677 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Expediente\Arrecadacao\Confeccao de Layouts.htm"
    'Case 678 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Expediente\Arrecadacao\Baixa Automatica.htm"
    'Case 680 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Expediente\Arrecadacao\Arrecadacao Manual.htm"
    'Case 687 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Beneficiados com Imunidade.htm"
    'Case 688 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Inconsistencia Imobiliaria.htm"
    'Case 689 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Relacao de contadores.htm"
    'Case 690 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Contadores e Arrecadacao.htm"
    'Case 691 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Inscritos Ativos.htm"
    'Case 692 '- RETIRADO 29/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Contribuintes em contencioso.htm"
    'Case 693 '- RETIRADO 30/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Editais notificacao.htm"
    'Case 695 '- RETIRADO 30/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Demonstrativos de arrecadacao.htm"
    'Case 696 '- RETIRADO 30/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Quantidade de Lancamento.htm"
    'Case 698 '- RETIRADO 30/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Extrato Individualizado.htm"
    'Case 699 '- RETIRADO 30/07/04 Rafael
    '    strHelp = ""
    'Case 697 '- RETIRADO 30/07/04 Rafael
    '    strHelp = ""
    'Case 682 '- RETIRADO 30/07/04 Rafael
    '    strHelp = ""
    'Case 700 '- RETIRADO 30/07/04 Rafael
    '    strHelp = "Relatorios\Cadastro Tecnico\Documentos Diversos.htm"
    'Case 765 '- RETIRADO 30/07/04 Rafael
    '    strHelp = "Relatorios\Controle de Arrecadacao\Documentos Diversos.htm"
    'Case 768 '- RETIRADO 03/08/04 Rafael
    '    strHelp = ""
    'Case 769 '- RETIRADO 03/08/04 Rafael
    '    strHelp = ""
    'Case 771 '- RETIRADO 03/08/04 Rafael
    '    strHelp = "Relatorios\Conta Corrente\Debitos Nao Inscritos.htm"
    'Case 772 '- RETIRADO 03/08/04 Rafael
    '    strHelp = "Relatorios\Conta Corrente\Documentos Diversos.htm"
    'Case 774 '- RETIRADO 03/08/04 Rafael
    '    strHelp = "Relatorios\Cobranca\Relacao de Documentos.htm"
    'Case 775 '- RETIRADO 03/08/04 Rafael
    '    strHelp = "Relatorios\Cobranca\Documentos Diversos.htm"
    'Case 780 '- RETIRADO 03/08/04 Rafael
    '    strHelp = "Relatorios\Fiscalizacao\Documentos Diversos.htm"
    'Case 777 '- RETIRADO 03/08/04 Rafael
    '    strHelp = "Relatorios\Fiscalizacao\Posicao de Alvaras.htm"
    'Case 783 '- RETIRADO 03/08/04 Rafael
    '    strHelp = "Relatorios\Divida Ativa\Livro da divida ativa.htm"
    'Case 784 '- RETIRADO 04/08/04 Rafael
    '    strHelp = "Relatorios\Divida Ativa\Relacaode Adimplencia em Divida Ativa.htm"
    'Case 785 '- RETIRADO 04/08/04 Rafael
    '    strHelp = "Relatorios\Divida Ativa\Relacao de Inadimplencia em Divida Ativa.htm"
    Case 379
        strHelp = "Ferramentas\Opcoes.htm"
    Case 382
        strHelp = "Alterar Senha.htm"
  End Select
  
  With hHelp
    .CHMFile = .HHSetHelpFile(1)
    .HHWindow = ""
    .HHTopicURL = strHelp
    .HHDisplayTopicURL
  End With
  
Exit_HtmlHelp:
End Sub

Public Sub FlatLook(ControleToolBar As Object, Optional EstiloLista As Variant)
    '--------------------------------------------------------------'
    ' FUN��O USADA PARA MUDAR O ESTILO DA TOOLBAR NOS FORMUL�RIOS. '
    '--------------------------------------------------------------'
    ' PAR�METROS:                                                  '
    '                                                              '
    ' 1 - ControleToolBar(Objeto ToolBar - Tipo Object)            '
    ' 2 - EstiloLista(Estilo do Toolbar - Tipo Variant)            '
    '--------------------------------------------------------------'

    On Error Resume Next

    Dim Estilo As Long, resultado As Long, Id_Toolbar As Long

    Id_Toolbar = FindWindowEx(ControleToolBar.hWnd, 0&, "ToolbarWindow32", vbNullString)
    Estilo = SendTBMessage(Id_Toolbar, TB_GETSTYLE, 0&, 0&)
    
    If EstiloLista = True Then
        Estilo = Estilo Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT Or CCS_NODIVIDER Or TBSTYLE_LIST
    End If
    If EstiloLista = False Then
        Estilo = TBSTYLE_FLAT
    End If
    resultado = SendTBMessage(Id_Toolbar, TB_SETSTYLE, 0, Estilo)
    ControleToolBar.Refresh
End Sub


Public Sub Formata_ListView(frmForm As Form)
    '--------------------------------------------'
    ' ESTA SUB FORMATA OS LISTVIEW DO FORMULARIO '
    '--------------------------------------------'
    ' PAR�METROS:                                '
    '                                            '
    ' 1 - frmForm(Formul�rio - Tipo Form)        '
    '--------------------------------------------'
    Dim intCountCtr As Integer
    Dim r           As Long
    Dim rStyle      As Long
    
    For intCountCtr = 0 To frmForm.Controls.Count - 1
        If TypeOf frmForm.Controls(intCountCtr) Is ListView Then
            rStyle = SendMessageLong(frmForm.Controls(intCountCtr).hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
            rStyle = rStyle Or LVS_EX_HEADERDRAGDROP
            r = SendMessageLong(frmForm.Controls(intCountCtr).hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
            rStyle = SendMessageLong(frmForm.Controls(intCountCtr).hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
            rStyle = rStyle Xor LVS_EX_FULLROWSELECT
            r = SendMessageLong(frmForm.Controls(intCountCtr).hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
        End If
    Next
End Sub


Public Function gstrColocaAspaSimples(strTexto As String) As String
    '---------------------------------------------------------------'
    ' FUN��O USADA PARA TROCAR ACENTO GRAVE PELA ASPAS SIMPLES      '
    '---------------------------------------------------------------'
    ' PAR�METRO:                                                    '
    '                                                               '
    ' 1 - Texto(Texto - Tipo String)                                '
    '---------------------------------------------------------------'
    Do While InStr(strTexto, "`") <> 0
        Mid$(strTexto, InStr(strTexto, "`"), 1) = "'"
    Loop
    gstrColocaAspaSimples = Trim(strTexto)
End Function

Function blnEmpresaBaixada(intPKIdEconomico As Integer, Optional blnMostrarMensagem As Boolean) As Boolean
Dim strSQL As String
Dim adoRec As ADODB.Recordset

    '---------------------------------------------------------------'
    ' FUN��O USADA PARA AVERIGUAR SE A EMPRESA ESTA BAIXADA.      '
    '---------------------------------------------------------------'
    ' PAR�METRO:                                                    '
    '                                                               '
    ' 1 - intPKIdEconomico - PKId tabela econ�mico        '
    ' 2 - blnMostrarMensagem - Para mostra mensagem '
    '---------------------------------------------------------------'

    strSQL = ""
    strSQL = strSQL & " SELECT dtmDataBaixa FROM "
    strSQL = strSQL & gstrEconomico
    strSQL = strSQL & " WHERE PKId = " & intPKIdEconomico
    strSQL = strSQL & " AND dtmDataBaixa IS NOT NULL"
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
        With adoRec
            If Not .EOF Then
                If blnMostrarMensagem Then
                    ExibeMensagem " C�lculo cancelado Empresa Baixada em " & !dtmDataBaixa
                End If
                blnEmpresaBaixada = True
            End If
        End With
    End If
    
End Function


Public Function gstrTiraAspaSimples(strTexto As String) As String
    '---------------------------------------------------------------'
    ' FUN��O USADA PARA TROCAR ASPAS SIMPLES POR ACENTO GRAVE.      '
    '---------------------------------------------------------------'
    ' PAR�METRO:                                                    '
    '                                                               '
    ' 1 - Texto(Texto - Tipo String)                                '
    '---------------------------------------------------------------'
    Do While InStr(strTexto, "'") <> 0
        Mid$(strTexto, InStr(strTexto, "'"), 1) = "`"
    Loop
    gstrTiraAspaSimples = Trim(strTexto)
End Function

Public Function gQueryTDB_VencimentoParcelasReceita(intTributo As Integer, intExercicio As Integer) As String
    '------------------------------------------------------------------'
    ' FUN��O USADA PARA MONTAR QUERY QUE BUSCA AS PARCELAS DA RECEITA. '
    '------------------------------------------------------------------'
    ' PAR�METRO:                                                       '
    '                                                                  '
    ' 1 - intTributo(Tipo de Tributo )                                 '
    ' 2 - intExercicio(Ano de Exercicio)                               '
    '------------------------------------------------------------------'

    Dim strSQL As String
    strSQL = "SELECT VP.PKId, VP.intNumero, VP.dtmDataDaParcela " & _
            " FROM " & gstrVencimentosDasParcelas & " VP," & _
            gstrVencimentos & " VC " & _
            " WHERE VP.intNumero >= 0 AND VC.PKId = VP.intVencimento " & _
            " AND VC.intTributo = " & intTributo & _
            " AND VP.intExercicio = " & intExercicio & _
            " ORDER BY intNumero "
    gQueryTDB_VencimentoParcelasReceita = strSQL
End Function

Public Function gBlnVerificaLancamentos(intExercicio As Integer, _
                                        IntComposicao As Integer, _
                                        strComposicao As String, _
                                        intNumeroParcelas As Integer, _
                                        dtmLancamento As String, _
                                        bitTodosMarcados As Byte, _
                                        strInscricaoCadastral As String, _
                                        Optional strInscricaoFinal As String) As Boolean

'******************************************************************************************
' Data: 12/05/2003
' Altera��o: - Substitui��o da chamada direta � stored procedure pela fun��o
'            gstrStoredProcedure.
' Respons�vel: Everton Bianchini
'******************************************************************************************

    '---------------------------------------------------------------------------'
    'USADA PARA VERIFICAR SE EXISTE LAN�AMENTO PARA CONTRIBUINTES SELECIONADOS.
    '---------------------------------------------------------------------------'
    ' PAR�METRO:                                                                '
    '                                                                           '
    ' 1 - intExercicio(Ano de Exercicio)                                        '
    ' 2 - intComposicao (PKId Composi��o da Receita a Ser Pesquisada.           '
    ' 3 - strComposicao (Detalhe da Composi��o da Receita a Ser Pesquisa        '
    ' 4 - bitTodosMarcados (Indica se s�o Todos Contribuinte a Serem Pesquisados'
    ' 5 - strInscricaoCadastral (Inscri��o Cadastral Inicial ou Fixa)           '
    ' 6 - strInscricaoFinal (Inscri��o Final para Determinar um Intervalo       '
    '---------------------------------------------------------------------------'
    
    Dim strSQL As String
    Dim strAux As String
    Dim adoResultado As ADODB.Recordset
    Dim ScrMouse As Integer
    
    ScrMouse = Screen.MousePointer
    strSQL = " FROM " & gstrLancamentoCalculo & _
            " WHERE intExercicio = " & intExercicio & _
            " AND intComposicaoReceita = " & IntComposicao & _
            " AND intNumeroDeParcelas = " & intNumeroParcelas
    If dtmLancamento <> "IPTU" Then
        strSQL = strSQL & " AND dtmLancamento = " & dtmLancamento
    End If
    If bitTodosMarcados = 0 Then
        If (strInscricaoCadastral <> "" And strInscricaoFinal <> "") Then
            If CDbl("0," & strInscricaoCadastral) > CDbl("0," & strInscricaoFinal) Then
                strAux = strInscricaoCadastral
                strInscricaoCadastral = strInscricaoFinal
                strInscricaoFinal = strAux
            End If
            strSQL = strSQL & " AND strInscricaoCadastral BETWEEN '" & _
                    strInscricaoCadastral & "' AND '" & strInscricaoFinal & "'"
        Else
            strSQL = strSQL & " AND strInscricaoCadastral = '" & strInscricaoCadastral & "'"
        End If
    End If
    strAux = "SELECT PKId " & strSQL
    strSQL = "SELECT strInscricaoCadastral " & strSQL & _
            " GROUP BY strInscricaoCadastral"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            strSQL = "Deseja efetuar Novamente o C�lculo de  " & _
            Chr(10) & strComposicao & "," & Chr(10)
            If adoResultado.RecordCount = 1 Then
                strSQL = strSQL & "para o Contribuinte de Inscri��o Cadastral " & Chr(10) & _
                        adoResultado("strInscricaoCadastral")
            Else
                strSQL = strSQL & "para " & adoResultado.RecordCount & " contribuintes "
            End If
            strSQL = strSQL & "," & Chr(10) & "referente ao per�odo de " & intExercicio & " ? "
            Screen.MousePointer = vbNormal
            If MsgBox(strSQL, vbYesNo, "Tribut�rio") = vbYes Then
            Screen.MousePointer = vbHourglass
                Set gobjBanco = Nothing
                Set gobjBanco = New clsBanco
'                gobjBanco.Execute ("sp_RemoveLancamentos '" & Replace(strAux, "'", Chr(34)) & "'")
                gobjBanco.Execute (gstrStoredProcedure("sp_RemoveLancamentos", "'" & Replace(strAux, "'", Chr(34)) & "'"))
                gBlnVerificaLancamentos = True
            End If
            Screen.MousePointer = ScrMouse
        Else
            gBlnVerificaLancamentos = True
        End If
    End If
End Function


Public Sub SelecionaInscricao(ByRef strInscricaoInicial As String, ByRef strInscricaoFinal As String)
    Dim i            As Integer
    Dim strInscricao As String
    'Inscri��o Inicial
    i = 1
    strInscricao = strInscricaoInicial
    strInscricaoInicial = ""
    Do While Mid(strInscricao, i, 1) <> "-"
        strInscricaoInicial = strInscricaoInicial & Mid(strInscricao, i, 1)
        i = i + 1
    Loop
    strInscricaoInicial = RTrim(strInscricaoInicial)
    'Inscri��o Final
    i = 1
    strInscricao = strInscricaoFinal
    strInscricaoFinal = ""
    Do While Mid(strInscricao, i, 1) <> "-" And i <= Len(strInscricao)
        strInscricaoFinal = strInscricaoFinal & Mid(strInscricao, i, 1)
        i = i + 1
    Loop
    strInscricaoFinal = RTrim(strInscricaoFinal)
End Sub

Public Function gblnBaixaCancelamento(intAlfa As Long, IntComposicao As Long, intAno As Integer, intParcela As String, dtmDataBaixa As String, blnMsg As Boolean, blnSimulado As Boolean, Optional PkidMovBancario As Long) As Boolean
    '-------------------------------------------------------------------------------------------'
    'USADA PARA VERIFICA��O DE PARCELAS CANCELADAS E BAIXA DAS MESMAS.                          '
    '-------------------------------------------------------------------------------------------'
    ' PAR�METRO:                                                                                '
    ' 1 - intAlfa (PKId da Tabela tblLacamentoALfa a Ser Pesquisada)                            '
    ' 2 - intComposicao (PKId Tabela tblcomposicaodareceita a Ser Pesquisada)                   '
    ' 3 - IntAno (Ano a ser Pesquisado)                                                         '
    ' 4 - intParcela (intParcela da Tabela tblLancamentoValor a Ser Pesquisada)                 '
    ' 5 - dtmDataBaixa (data que ser� inserida na tabela TblLancamentoPagamento)                '
    ' 6 - blnSimulado (False = Baixa) / (True = Simulado)que vem da tela processamento de baixa '
    '-------------------------------------------------------------------------------------------'

    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim adoParcela      As ADODB.Recordset
    Dim strComposicao   As String
    Dim strAviso        As String
    Dim strEmissao      As String
    
    gblnBaixaCancelamento = False
    
    'Select para busca do Campo strEmissao da Tabela TblLancamentoAlfa
    
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "CR.Strdescricao, "
    strSQL = strSQL & gstrCONVERT(cdt_numeric, "LA.strNumeroAviso") & " strNumeroAviso,"
    strSQL = strSQL & "UPPER(La.strEmissao) As strEmissao "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strSQL = strSQL & "La.Pkid = " & intAlfa
    
    Set gobjBanco = Nothing
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            strComposicao = gstrENulo(adoResultado!strDescricao)
            strAviso = gstrENulo(adoResultado!strNumeroAviso)
            strEmissao = gstrENulo(adoResultado!strEmissao)
        Else
            Exit Function
        End If
    End If

    'Esta Query verifica se exite parcelas canceladas e retorna o n�mero da
    'parcela e o c�digo da baixa que esta na tabela TblFormaPagtoVencimentos
        
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "FPV.intParcela, "
    strSQL = strSQL & "PC.IntCodigoBaixa "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrFormaPagtoVencimentos & " FPV, "
    strSQL = strSQL & "( "
            strSQL = strSQL & "Select "
            strSQL = strSQL & "FPC.Intformapagtovencimentoscancel, "
            strSQL = strSQL & "FPC.INTCODIGOBAIXA "
            strSQL = strSQL & "From "
            strSQL = strSQL & gstrParametroIPTU & " PI, "
            strSQL = strSQL & gstrParametroIPTUPagto & " PIP, "
            strSQL = strSQL & gstrFormaPagtoVencimentos & " FPV, "
            strSQL = strSQL & gstrFormaPagtoCancelamentos & " FPC "
            strSQL = strSQL & "Where "
            strSQL = strSQL & "PI.Pkid = PIP.Intparametroiptu AND "
            strSQL = strSQL & "PIP.Pkid = FPV.Intformapagto AND "
            strSQL = strSQL & "FPV.Pkid = FPC.Intformapagtovencimentos AND "
            strSQL = strSQL & "PI.INTCOMPOSICAODARECEITA = " & IntComposicao
            strSQL = strSQL & " AND PI.INTEXERCICIO = " & intAno
            strSQL = strSQL & " AND PI.Stremissao = " & strEmissao
            strSQL = strSQL & " AND FPV.Intparcela = " & intParcela
    strSQL = strSQL & " ) PC "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "FPV.Pkid = PC.Intformapagtovencimentoscancel"
    
    'Esta Query verifica se exite parcelas na tabela tblLancamentoValor pelo campo intParcela iguais
    'a que vem da tabela TblFormaPagtoVencimentos pelo mesmo campo intParcela
    
    Set gobjBanco = Nothing
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        With adoResultado
            If .RecordCount > 0 Then
                Do While Not .EOF
                    strSQL = ""
                    strSQL = strSQL & "Select "
                    strSQL = strSQL & "LV.Pkid                     As intlancamentovalor, "
                    strSQL = strSQL & "LV.Intparcela               As Parcela "
                    strSQL = strSQL & "From "
                    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
                    strSQL = strSQL & gstrLancamentoValor & " LV "
                    strSQL = strSQL & "Where "
                    strSQL = strSQL & "LA.Pkid = LV.Intlancamentoalfa AND "
                    strSQL = strSQL & "LV.Intlancamentoalfa = " & intAlfa & " And "
                    strSQL = strSQL & "LV.intParcela = " & gstrENulo(!intParcela)
                
                    'Esta Query insere na Tabela TblLancamentoValor o pkid da tabela TblLancamentoValor assim
                    'concretizando a baixa
                    
                    Set gobjBanco = Nothing
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strSQL, 10, adoParcela) Then
                        If .RecordCount > 0 And Not adoParcela.EOF Then
                            If Not IsNull(adoParcela!intLancamentoValor) Then
                                If blnSimulado = False Then
                                    strSQL = ""
                                    strSQL = strSQL & "Insert Into "
                                    strSQL = strSQL & gstrLancamentoPagamento & " ("
                                    strSQL = strSQL & "intlancamentovalor, "
                                    strSQL = strSQL & "dblvalorprincipal, "
                                    strSQL = strSQL & "dblvalormulta, "
                                    strSQL = strSQL & "dblvalorjuros, "
                                    strSQL = strSQL & "dblvalorcorrecao, "
                                    strSQL = strSQL & "dtmdtpagamento, "
                                    strSQL = strSQL & "intcodigobaixa, "
                                    strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr) "
                                    strSQL = strSQL & "VaLues( "
                                    strSQL = strSQL & gstrENulo(adoParcela!intLancamentoValor) & ", "
                                    strSQL = strSQL & "0, "
                                    strSQL = strSQL & "0, "
                                    strSQL = strSQL & "0, "
                                    strSQL = strSQL & "0, "
                                    strSQL = strSQL & gstrConvDtParaSql(Trim(dtmDataBaixa)) & ", "
                                    strSQL = strSQL & gstrENulo(adoResultado!intCodigoBaixa) & ", "
                                    strSQL = strSQL & strGETDATE & ", "
                                    strSQL = strSQL & glngCodUsr
                                    strSQL = strSQL & ") "
                                End If
                            Else
                                If blnMsg = True Then
                                    ExibeMensagem "Receita:" & _
                                                                "   " & strComposicao & _
                                        Chr(13) & "Ano:" & _
                                                                "   " & intAno & _
                                        Chr(13) & "Aviso:" & _
                                                                "   " & strAviso & _
                                        Chr(13) & "Parcela:" & _
                                                                "   " & gstrENulo(adoParcela!PARCELA) & _
                                        Chr(13) & "N�o pode ser baixada "
                                    Exit Function
                                Else
                                    gobjBanco.ExecutaRollbackTrans
                                    strSQL = "INSERT INTO " & gstrCriticaBaixa & " (intMovimentoBancario, intTipoCritica, dtmDtAtualizacao, lngCodUsr)" & _
                                             " VALUES (" & PkidMovBancario & ", 9" & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & ")"
                                    gobjBanco.Execute strSQL
                                    gobjBanco.ExecutaBeginTrans
                                    Exit Function
                                End If
                            End If
                           If gobjBanco.Execute(strSQL) = False Then Exit Function
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If

    gblnBaixaCancelamento = True
    
End Function

Public Function gblnAnaliseDaReceita(PkidLancamentoValor As Long, PkidContaBancaria As Long, PkidComposicaoDaReceita As Long, DBLVALOR As Double, dblValorMulta As Double, dblValorJuros As Double, dblValorCorrecao As Double, dtmDataDaBaixa As Date, intAlfa As Long, blnSimulado As Boolean, blnExibeMsg As Boolean, blnGerarMovimento As Boolean, Optional PkidMovBancario As Long, Optional blnSomenteMovOrcamentario As Boolean = False) As Boolean
Dim adoResultado         As ADODB.Recordset
Dim strSQL               As String
Dim blnEmDividaAtiva     As Boolean

Dim blnMultaOK           As Boolean 'Tipo valor = 1
Dim blnJurosOK           As Boolean 'Tipo valor = 2
Dim blnCorrecaoOK        As Boolean 'Tipo valor = 3

Dim dblValorTotalReceita As Double
Dim dblValorDiferenca    As Double
Dim varAux               As Variant
Dim intFor               As Integer
Dim intForReceitas       As Integer

Dim strUltimaReceita    As String 'Utilizada para jogar a diferenca na receita correspondente

Dim strMsg              As String
Dim strMsgReceitas      As String

Dim bytTipoCritica      As Byte
Dim strDetalheCritica   As String
Dim strComposicao       As String
Dim strAviso            As String
Dim strParcela          As String

On Error GoTo Problema_Na_Rotina
        
       
    'Caso o ultimo registro nao possua parcela vamos forcar o mov no orcamentario
    If blnSomenteMovOrcamentario Then GoTo SomenteMovOrcamentario
    
    'Se o valor for 0 ou nao for informada a conta nao vamos passar pela contabilidade
    If DBLVALOR = 0 Or PkidContaBancaria = 0 Then
        gblnAnaliseDaReceita = True
        Exit Function
    End If
    
    gblnAnaliseDaReceita = False
    
    'Select para busca do Campo strEmissao da Tabela TblLancamentoAlfa
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "CR.Strdescricao, "
    strSQL = strSQL & gstrCONVERT(cdt_numeric, "LA.strNumeroAviso") & " strNumeroAviso "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA "
    strSQL = strSQL & "WHERE "
    strSQL = strSQL & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strSQL = strSQL & "La.Pkid = " & intAlfa
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            strComposicao = gstrENulo(adoResultado!strDescricao)
            strAviso = gstrENulo(adoResultado!strNumeroAviso)
        Else
            Exit Function
        End If
    End If
    
    'Vamos atribuir valores �s variaveis para depois identificar se � preciso serem retornadas na consulta
    blnMultaOK = dblValorMulta = 0
    blnJurosOK = dblValorJuros = 0
    blnCorrecaoOK = dblValorCorrecao = 0
        
    strMsg = ""
    strMsgReceitas = ""
    
    'Select para verificar se as receitas do lancamento constam na tabela de ReceitasExercicio
    strSQL = ""
    strSQL = "SELECT R.strDescricao FROM " & gstrLancamentoReceita & " LR, " & gstrReceita & " R WHERE R.pkid = LR.intReceita AND LR.INTLANCAMENTOVALOR = " & PkidLancamentoValor & " AND "
    strSQL = strSQL & "LR.INTRECEITA NOT IN (SELECT INTRECEITA FROM " & gstrReceitasExercicio & " WHERE INTEXERCICIO = " & Year(dtmDataDaBaixa) & ")"
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        Do While Not adoResultado.EOF
            strMsgReceitas = strMsgReceitas & Chr(13) & "A receita " & Trim(adoResultado("strDescricao").Value) & " n�o foi encontrada para o exerc�cio " & Year(dtmDataDaBaixa) & "."
            adoResultado.MoveNext
        Loop
    End If
    
    'Caso nao seja encontrada a receita na tabela receitasexercicio
    If Len(strMsgReceitas) > 0 Then
        ExibeMensagem strMsgReceitas
        bytTipoCritica = 5
        GoTo ExibeMensagem
    End If
    
    'Vamos verificar na tblLancamentoValor se esta em Divida Ativa
    strSQL = "SELECT Pkid FROM " & gstrLancamentoValor & " WHERE Pkid = " & PkidLancamentoValor & " AND Not intLancamentoAlfaDativa Is Null"
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        blnEmDividaAtiva = adoResultado.RecordCount > 0
    End If
    
    'Vamos buscar as receitas com valores que formam o tributo recebido
    strSQL = "SELECT " & gstrCASEWHEN("LR.dblValor", "0, 1 / (SELECT Count(intReceita) FROM " & gstrLancamentoReceita & " WHERE intLancamentoValor = " & PkidLancamentoValor & ")", "LR.dblValor / (SELECT Sum(dblValor) FROM " & gstrLancamentoReceita & " WHERE intLancamentoValor = " & PkidLancamentoValor & ")") & " Proporcao ," & _
             "LR.dblValor, " & IIf(blnEmDividaAtiva, "R.intDividaAtiva", "R.Pkid") & " intReceita , R.bytTipoReceita, " & gstrCASEWHEN("R.bytTipoReceita", "0, RE.intClassificacaoDaReceita", "RE.intPlanoConta") & " ReceitaContabil, 0 TipoValor, LV.intParcela, R.strSigla " & _
             "FROM " & gstrLancamentoReceita & " LR, " & gstrReceita & " R, " & gstrReceitasExercicio & " RE, " & gstrLancamentoValor & " LV " & _
             "WHERE LR.intLancamentoValor = " & PkidLancamentoValor & " AND R.Pkid = LR.intReceita AND " & _
             "RE.intReceita " & strOUTJOracle & " =" & strOUTJSQLServer & IIf(blnEmDividaAtiva, " R.intDividaAtiva", " R.Pkid") & " AND " & _
             "RE.intExercicio" & strOUTJOracle & " = " & Year(dtmDataDaBaixa) & " AND LV.Pkid = LR.intLancamentoValor"
    strSQL = strSQL & " UNION ALL "
    strSQL = strSQL & "SELECT 1 Proporcao, " & gstrConvVrParaSql(dblValorMulta) & " dblValor, " & IIf(blnEmDividaAtiva, "R.intDividaAtiva", "R.Pkid") & " intReceita, " & _
                      "R.Byttiporeceita, " & gstrCASEWHEN("R.bytTipoReceita", "0, RE.intClassificacaoDaReceita", "RE.intPlanoConta") & " ReceitaContabil, 1 TipoValor, 0 intParcela, R.strSigla " & _
                      "FROM " & gstrParametroAtualizacao & " PA, " & gstrReceita & " R, " & gstrReceitasExercicio & " RE " & _
                      "WHERE PA.intExercicio = " & Year(dtmDataDaBaixa) & " AND PA.intComposicaoReceita = " & PkidComposicaoDaReceita & " AND R.pkid  = PA.IntReceitaMulta AND " & _
                      "RE.intReceita " & strOUTJOracle & " =" & strOUTJSQLServer & IIf(blnEmDividaAtiva, " R.intDividaAtiva", " R.Pkid") & " AND RE.intExercicio" & strOUTJOracle & " = " & Year(dtmDataDaBaixa)
    strSQL = strSQL & " UNION ALL "
    strSQL = strSQL & "SELECT 1 Proporcao, " & gstrConvVrParaSql(dblValorJuros) & " dblValor, " & IIf(blnEmDividaAtiva, "R.intDividaAtiva", "R.Pkid") & " intReceita, " & _
                      "R.Byttiporeceita, " & gstrCASEWHEN("R.bytTipoReceita", "0, RE.intClassificacaoDaReceita", "RE.intPlanoConta") & " ReceitaContabil, 2 TipoValor, 0 intParcela, R.strSigla " & _
                      "FROM " & gstrParametroAtualizacao & " PA, " & gstrReceita & " R, " & gstrReceitasExercicio & " RE " & _
                      "WHERE PA.intExercicio = " & Year(dtmDataDaBaixa) & " AND PA.intComposicaoReceita = " & PkidComposicaoDaReceita & " AND R.pkid  = PA.IntReceitaJuros AND " & _
                      "RE.intReceita " & strOUTJOracle & " =" & strOUTJSQLServer & IIf(blnEmDividaAtiva, " R.intDividaAtiva", " R.Pkid") & " AND RE.intExercicio" & strOUTJOracle & " = " & Year(dtmDataDaBaixa)
    strSQL = strSQL & " UNION ALL "
    strSQL = strSQL & "SELECT 1 Proporcao, " & gstrConvVrParaSql(dblValorCorrecao) & " dblValor, " & IIf(blnEmDividaAtiva, "R.intDividaAtiva", "R.Pkid") & " intReceita, " & _
                      "R.Byttiporeceita, " & gstrCASEWHEN("R.bytTipoReceita", "0, RE.intClassificacaoDaReceita", "RE.intPlanoConta") & " ReceitaContabil, 3 TipoValor, 0 intParcela, R.strSigla " & _
                      "FROM " & gstrParametroAtualizacao & " PA, " & gstrReceita & " R, " & gstrReceitasExercicio & " RE " & _
                      "WHERE PA.intExercicio = " & Year(dtmDataDaBaixa) & " AND PA.intComposicaoReceita = " & PkidComposicaoDaReceita & " AND R.pkid  = PA.IntReceitaCorrecao AND " & _
                      "RE.intReceita " & strOUTJOracle & " =" & strOUTJSQLServer & IIf(blnEmDividaAtiva, " R.intDividaAtiva", " R.Pkid") & " AND RE.intExercicio" & strOUTJOracle & " = " & Year(dtmDataDaBaixa)
                      
    'Vamos preencher o array com todas as receitas encontradas
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                
        If Not adoResultado.EOF Then
            
            'Preenchendo o array de bancos
            With adoResultado
        
                dblValorTotalReceita = 0
                strUltimaReceita = ""
                
                Do While Not .EOF
                    
                    If adoResultado("dblValor") = 0 And adoResultado("TipoValor") <> 0 Then
                        GoTo ProximoMovimento
                    End If
                    
                    'Vamos verificar se � o primeiro movimento a passar pelo array
                    If aAnaliseReceita Is Nothing Then
                        Set aAnaliseReceita = New XArrayDB
                        aAnaliseReceita.ReDim 0, 0, 0, 5
                        intForReceitas = 0
                    'Caso nao seja o primeiro vamos verificar se ja existe a receita com a mesma conta
                    'se existir vamos somar o valor
                    Else
                        intForReceitas = -1
                        For intFor = 0 To aAnaliseReceita.Count(1) - 1
                            If aAnaliseReceita(intFor, 3) = !ReceitaContabil And aAnaliseReceita(intFor, 5) = PkidContaBancaria Then
                                intForReceitas = intFor
                                Exit For
                            End If
                        Next

                        If intForReceitas = -1 Then
                            aAnaliseReceita.ReDim 0, aAnaliseReceita.UpperBound(1) + 1, 0, 5
                            intForReceitas = aAnaliseReceita.UpperBound(1)
                        End If
                    End If
                    
                    'PkId da receita
                    varAux = !intReceita
                    aAnaliseReceita(intForReceitas, 0) = varAux
            
                    If adoResultado("TipoValor") = 0 Then
                        
                        'Valor proporcional
                        varAux = gstrConvVrDoSql(DBLVALOR * !Proporcao)
                        aAnaliseReceita(intForReceitas, 1) = CCur(aAnaliseReceita(intForReceitas, 1)) + varAux
                        
                        dblValorTotalReceita = dblValorTotalReceita + varAux
                        
                        strParcela = !intParcela
                        strUltimaReceita = gstrENulo(!ReceitaContabil)
                        
                    Else
                        
                        'Multa, Juros e Correcao � o valor exato
                        If adoResultado("TipoValor") = 1 Then
                            varAux = gstrConvVrDoSql(dblValorMulta)
                            blnMultaOK = True
                        ElseIf adoResultado("TipoValor") = 2 Then
                            varAux = gstrConvVrDoSql(dblValorJuros)
                            blnJurosOK = True
                        ElseIf adoResultado("TipoValor") = 3 Then
                            varAux = gstrConvVrDoSql(dblValorCorrecao)
                            blnCorrecaoOK = True
                        End If
                        
                        aAnaliseReceita(intForReceitas, 1) = CCur(aAnaliseReceita(intForReceitas, 1)) + varAux
                        
                    End If
                    
                    'Tipo da receita (Orcamentaria = 0, Extra = 1)
                    varAux = !bytTipoReceita
                    aAnaliseReceita(intForReceitas, 2) = varAux
                    
                    'Caso nao seja encontrado a receita contabil de um dos itens, vamos finalizar a function
                    If IsNull(!ReceitaContabil) Then
                        strMsg = "N�o foram encontrados registros correspondentes � Contabilidade em algum dos itens da composi��o da receita."
                        bytTipoCritica = 7
                        strDetalheCritica = !strsigla
                        GoTo ExibeMensagem
                    End If
                    
                    'Pkid de PlanoConta ou PrevisaoDaReceita
                    varAux = !ReceitaContabil
                    aAnaliseReceita(intForReceitas, 3) = varAux
            
                    'Identificador do tipo de valor (Principal = 0, Multa, Juros e Correcao = 1)
                    varAux = !TipoValor
                    aAnaliseReceita(intForReceitas, 4) = varAux
            
                    'Pkid da Conta Bancaria
                    varAux = PkidContaBancaria
                    aAnaliseReceita(intForReceitas, 5) = varAux
                    
ProximoMovimento:

                    .MoveNext
            
                Loop
                
                'Caso o valor do parametro nao seja 0,00 e nao tenha sido encontrado na tblParametroAtualizacao
                If Not blnMultaOK Or Not blnJurosOK Or Not blnCorrecaoOK Then
                    strMsg = "N�o foi(ram) encontrada(s) Receita(s) ou Exerc�cio da Receita correspondente(s) � Multa, Juros ou Corre��o da Composi��o da Receita informada em Par�metros de Atualiza��o."
                    bytTipoCritica = 8
                    GoTo ExibeMensagem
                End If
                
                dblValorDiferenca = CCur(DBLVALOR) - CCur(dblValorTotalReceita)
                
                'Vamos verificar se existe diferenca entre o valor do parametro e o calculado proporcionalmente
                If dblValorDiferenca <> 0 Then
                    If Abs(dblValorDiferenca) >= 1 Then
                        strMsg = "Foi encontrada diferen�a de valores."
                        bytTipoCritica = 3
                        GoTo ExibeMensagem
                    Else
                        'Vamos adicionar a diferenca na primeira receita com valor superior � diferenca
                        For intFor = 0 To aAnaliseReceita.Count(1) - 1
                            'If aAnaliseReceita(intFor, 1) > dblValorDiferenca And aAnaliseReceita(intFor, 4) = 0 And aAnaliseReceita(intFor, 3) = strUltimaReceita And aAnaliseReceita(intFor, 5) = PkidContaBancaria Then
                            If aAnaliseReceita(intFor, 1) > dblValorDiferenca And aAnaliseReceita(intFor, 3) = strUltimaReceita And aAnaliseReceita(intFor, 5) = PkidContaBancaria Then
                                aAnaliseReceita(intFor, 1) = aAnaliseReceita(intFor, 1) + dblValorDiferenca
                                Exit For
                            End If
                        Next
                    End If
                End If
                
            End With
             
            If blnGerarMovimento Then
            
SomenteMovOrcamentario:

'*****************************************************************************************************************
'PARTE QUE GERA MOVIMENTO NA CONTABILIDADE (AQUI O BICHO PEGA!)
                Dim adoAux                As ADODB.Recordset
                Dim strCodigoOrcamentario As String
                Dim lngEvento             As Long
                Dim lngConta              As Long
                Dim lngPlanoConta         As Long
                Dim lngPkidInicial        As Long
                Dim lngPkidUltArrecadacao As Long
                Dim lngUltNumArrecadacao  As Long
                Dim aTipoMovimento()      As Variant
                Dim aContaExtra()         As Variant
                Dim aValor()              As Variant
                
                'for inferior ao registro que inciamos
                lngPkidInicial = glngPegaUltimaChave(gstrArrecadacaoReceita, "Pkid")
                
                For intFor = 0 To aAnaliseReceita.Count(1) - 1
                        
                    If Val(aAnaliseReceita(intFor, 5)) = 0 Then
                        ExibeMensagem "Ocorreram erros durante a grava��o dos movimentos do Evento Cont�bil devido a n�o existir Receita Cont�bil para alguma(s) Receita(s)."
                        Exit Function
                    End If
                    
                    Set adoAux = New ADODB.Recordset
                    
                    'Vamos obter o plano de conta para a gravacao na tabela ArrecadacaDaReceita
                    adoAux.Open "SELECT PC.pkid FROM " & gstrPlanoConta & " PC " & _
                                "WHERE PC.intContaBancaria = " & aAnaliseReceita(intFor, 5), gcncADOMain, adOpenKeyset, adLockOptimistic
                            
                    If Not adoAux.EOF Then
                        lngPlanoConta = adoAux!Pkid
                    Else
                        lngPlanoConta = 0
                    End If
                    
                    adoAux.Close: Set adoAux = Nothing
                    
                    'Vamos verificar se e conta Orcamentaria ou Extra
                    If aAnaliseReceita(intFor, 2) = 0 Then
                        
                        ReDim aContaExtra(1)
                        ReDim aValor(1)
                        ReDim aTipoMovimento(1)
                        
                        aContaExtra(1) = lngPlanoConta
                        aValor(1) = aAnaliseReceita(intFor, 1)
                        
                        Set adoAux = New ADODB.Recordset
                    
                        'Consultas para o retorno do Evento
                        adoAux.Open "SELECT CO.Pkid, CO.strCodigoOrcamentario FROM " & gstrPrevisaoDaReceita & " PR, " & gstrCodigoOrcamentario & " CO " & _
                                    "WHERE PR.Pkid = " & aAnaliseReceita(intFor, 3) & " AND CO.Pkid = PR.intCodigoOrcamentario", gcncADOMain, adOpenKeyset, adLockOptimistic

                        strCodigoOrcamentario = adoAux!strCodigoOrcamentario
                        lngConta = adoAux!Pkid
                        
                        adoAux.Close: Set adoAux = New ADODB.Recordset
                    
                        adoAux.Open "SELECT EC.intEvento, PC.pkid FROM " & gstrEventoContaContabilCredito & " EC, " & gstrPlanoConta & " PC " & _
                                    "WHERE EC.intEvento IN (SELECT pkid FROM " & gstrEvento & " WHERE intTipoEvento = 1) AND " & _
                                    "PC.Pkid = EC.intContaContabil AND " & strSUBSTRING & "(PC.strContaContabil,1,3) = '4" & Mid(strCodigoOrcamentario, 1, 2) & "'", gcncADOMain, adOpenKeyset, adLockOptimistic
                                
                        If Not adoAux.EOF Then
                            lngEvento = adoAux!intEvento
                        Else
                            lngEvento = 0
                        End If
                        
                        aTipoMovimento(1) = 1
                    
                    'Extra
                    Else
                
                        ReDim aContaExtra(2)
                        ReDim aValor(2)
                        ReDim aTipoMovimento(2)
                        
                        aContaExtra(1) = lngPlanoConta
                        aContaExtra(2) = aAnaliseReceita(intFor, 3)
                        
                        aValor(1) = aAnaliseReceita(intFor, 1)
                        aValor(2) = aAnaliseReceita(intFor, 1)
                        
                        aTipoMovimento(1) = 1
                        aTipoMovimento(2) = 0
                        
                        lngConta = aAnaliseReceita(intFor, 3)
                        lngEvento = 0
                        
                    End If
                    
                    'Vamos verificar se ja existe a Receita para o registro atual
                    strSQL = "SELECT Pkid, intNumero FROM " & gstrArrecadacaoReceita & " WHERE dtmData = " & gstrConvDtParaSql(dtmDataDaBaixa) & " AND intContaContabil = " & lngPlanoConta & " AND intEvento = " & lngEvento & " AND Pkid >= " & lngPkidInicial
    
                    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                    
                        If Not adoResultado.EOF Then
                            lngPkidUltArrecadacao = adoResultado!Pkid
                            lngUltNumArrecadacao = adoResultado!INTNUMERO
                        Else
                
                            strSQL = "INSERT INTO " & gstrArrecadacaoReceita & " ("
                            strSQL = strSQL & "intNumero, intExercicio, dtmData, intContaContabil, bytImportacao, intEvento, strHistorico, "
                            strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr) "
                            strSQL = strSQL & "SELECT " & gstrISNULL("MAX(intNumero) + 1", "0" + 1) & " , " & Year(dtmDataDaBaixa) & ", "
                            strSQL = strSQL & gstrConvDtParaSql(dtmDataDaBaixa) & ", " & lngPlanoConta & ", 2, " & lngEvento & ", 'Movimento Banc�rio', "
                            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                            strSQL = strSQL & glngCodUsr & " FROM " & gstrArrecadacaoReceita & " WHERE " & gstrDATEPART(strYEAR, "dtmData") & " = " & Year(dtmDataDaBaixa)
                        
                            If Not gobjBanco.Execute(strSQL) Then
                                strMsg = "Ocorreu um erro ao gravar dados para Arrecada��o de Receita. A opera��o foi cancelada."
                                bytTipoCritica = 4
                                GoTo ExibeMensagem
                            End If
                        
                            lngPkidUltArrecadacao = glngRetornaPkidTabelaPai("seqtblArrecadacaoReceita", gstrArrecadacaoReceita)
                            
                            Set adoAux = New ADODB.Recordset

                            adoAux.Open "SELECT " & gstrISNULL("MAX(intNumero)", "0") & " MaxNumero  FROM " & gstrArrecadacaoReceita, gcncADOMain, adOpenKeyset, adLockOptimistic
                            lngUltNumArrecadacao = adoAux!MaxNumero
    
                            adoAux.Close: Set adoAux = Nothing
                    
                        End If
                    
                    End If
                
                    'Vamos verificar se ja existe a Conta da Receita para o registro atual
                    strSQL = "SELECT Pkid, dblValorOrcamentario FROM " & gstrContaArrecadacaoReceita & " WHERE intArrecadacao = " & lngPkidUltArrecadacao & " AND intConta = " & aAnaliseReceita(intFor, 3)
    
                    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                    
                        If adoResultado.EOF Then
                        
                            'Vamos inserir na tabela de Contas Arrecadacao de Receita
                            strSQL = "INSERT INTO " & gstrContaArrecadacaoReceita & " ("
                            strSQL = strSQL & "intArrecadacao, intConta, dblValorOrcamentario, bytCancelado, dtmDataCancelamento, bytTipo, "
                            strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr) VALUES ("
                            strSQL = strSQL & lngPkidUltArrecadacao & ", " & lngConta & ", "
                            strSQL = strSQL & gstrConvVrParaSql(Abs(aAnaliseReceita(intFor, 1))) & ", " & IIf(aAnaliseReceita(intFor, 1) < 0, 1, 0) & ", " & IIf(aAnaliseReceita(intFor, 1) < 0, gstrConvDtParaSql(dtmDataDaBaixa), "NULL") & ", " & aAnaliseReceita(intFor, 2) & ", "
                            strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                            strSQL = strSQL & glngCodUsr & ")"
                        Else
                    
                            'Vamos atualizar na tabela de Contas Arrecadacao de Receita
                            strSQL = "UPDATE " & gstrContaArrecadacaoReceita & " SET "
                            strSQL = strSQL & "dblValorOrcamentario = " & gstrConvVrParaSql(Abs(adoResultado("dblValorOrcamentario").Value + aAnaliseReceita(intFor, 1))) & ", "
                            strSQL = strSQL & "bytCancelado = " & IIf(adoResultado("dblValorOrcamentario").Value + aAnaliseReceita(intFor, 1) < 0, 1, 0)
                            strSQL = strSQL & "WHERE Pkid = " & adoResultado("Pkid").Value
                        End If
                    
                        If Not gobjBanco.Execute(strSQL) Then
                            strMsg = "Ocorreu um erro ao gravar dados para Contas Arrecada��o de Receita. A opera��o foi cancelada."
                            bytTipoCritica = 4
                            GoTo ExibeMensagem
                        End If
                
                    End If
                
                    If Not GeraMovimentosByEvento(lngEvento, Str(dtmDataDaBaixa), Str(aAnaliseReceita(intFor, 1)), "", Str(lngUltNumArrecadacao), "6", aContaExtra, aTipoMovimento, IIf(UBound(aContaExtra) > 1, True, False), aValor, True) Then
                        strMsg = "Ocorreram erros durante a grava��o dos movimentos do Evento Contabil."
                        bytTipoCritica = 4
                        GoTo ExibeMensagem
                    End If

                Next
            
            End If
'*****************************************************************************************************************

        Else
            strMsg = "N�o foram encontrados registros correspondentes � Contabilidade."
            bytTipoCritica = 1
            GoTo ExibeMensagem
        End If
        
    End If
    
    gblnAnaliseDaReceita = True
        
    Exit Function
    
ExibeMensagem:

    If blnExibeMsg = True Then
        ExibeMensagem strMsg & _
            Chr(13) & "Receita:" & _
                                    "   " & strComposicao & _
            Chr(13) & "Ano:" & _
                                    "   " & Year(dtmDataDaBaixa) & _
            Chr(13) & "Aviso:" & _
                                    "   " & strAviso & _
            Chr(13) & "Parcela:" & _
                                    "   " & strParcela & _
            Chr(13) & "N�o pode ser baixada. "
    Else
        gobjBanco.ExecutaRollbackTrans
        strSQL = "INSERT INTO " & gstrCriticaBaixa & " (intMovimentoBancario, intTipoCritica, strDetalhe, dtmDtAtualizacao, lngCodUsr)" & _
                 " VALUES (" & PkidMovBancario & "," & bytTipoCritica & "," & strDetalheCritica & "," & gstrConvDtParaSql(gstrDataDoSistema) & ", " & glngCodUsr & ")"
        gobjBanco.Execute strSQL
        gobjBanco.ExecutaBeginTrans
    End If
    
    Exit Function
    
Problema_Na_Rotina:
   
   ExibeDetalheErro "Erro na rotina gblnAnaliseDaReceita." & Chr(13) & Err.Number & " - " & Err.Description

End Function

Private Sub MontaMenuDocumentos(ab As ActiveBar2, chd)
   
   Set chd = ab.Bands("bndFormulario").ChildBands.Add("ChildBand")
   
   ab.Bands("bndFormulario").ChildBands.ChildBandCaptionAlignment = ddCACenter
   
   ab.Bands("bndFormulario").ChildBands.BackColor = &HE0E0E0
   
   With chd
      .Name = "chdDocumentos"
      .Caption = "Documentos"
      .GrabHandleStyle = ddGSCaption
      .flags = 127
      .Width = 12000
   End With
   
   'Dim adoRec        As ADODB.Recordset
   'Dim objFile       As Scripting.file
   'Dim stpSQL1       As String
   'Dim stpSQL2       As String
   'Dim objFiles      As Scripting.Files
   'Dim stpFolder     As String
   'Dim objFolder     As Scripting.Folder
   'Dim objFileSystem As Scripting.FileSystemObject
   'Dim intTextoAtual As Integer

   'intTextoAtual = 100

   'stpFolder = gstrDirDocumentos & "\Documentos\" & App.ProductName & "\WordModelos"

   'Set objFileSystem = New Scripting.FileSystemObject

   'If objFileSystem.FolderExists(stpFolder) Then

      'Set objFolder = objFileSystem.GetFolder(stpFolder)

      'Set objFiles = objFolder.Files
   
      'For Each objFile In objFiles
      
         'If UCase$(Right(objFile.Name, 3)) = "DOT" Then
         
            'stpSQL1 = "SELECT DocDescription FROM " & gstrNumeradorDocumentos & " WHERE DocDescription = '" & Left(objFile.Name, Len(objFile.Name) - 4) & "'"
            'stpSQL2 = "INSERT INTO " & gstrNumeradorDocumentos & " (DocDescription) VALUES ('" & Left(objFile.Name, Len(objFile.Name) - 4) & "')"

            'Set adoRec = gcncADOMain.Execute(stpSQL1, , adCmdText)
            
            'If adoRec.EOF Then gcncADOMain.Execute stpSQL2, , adCmdText
            
            'adoRec.Close
            
           'MontaBotoes chd, intTextoAtual + 1000, "mnuDocumentos", Left(objFile.Name, Len(objFile.Name) - 4), , , , Left(objFile.Name, Len(objFile.Name) - 4) & "|" & objFile.Path
            'MontaBotoes chd, intTextoAtual + 1000, "mnuDocumentos", Left(objFile.Name, Len(objFile.Name) - 4), , , , objFile.Path
            
            'intTextoAtual = intTextoAtual + 1
            
         'End If
      'Next
    
   'End If
   
   'Set objFile = Nothing
   
   'Set objFiles = Nothing
   
   'Set objFolder = Nothing
   
   'Set objFileSystem = Nothing
   
   MontaBotoes chd, 1191, "mnuDocumentos", "Alvar� de Funcionamento"
   MontaBotoes chd, 1364, "mnuDocumentos", "Cadastro Mobili�rio para Fins Tribut�rios"
   MontaBotoes chd, 1156, "mnuDocumentos", "Certid�o de Valor Venal"
   MontaBotoes chd, 1186, "mnuDocumentos", "Certid�o Mobili�ria"
   MontaBotoes chd, 1020, "mnuDocumentos", "Documentos Impressos"
   MontaBotoes chd, 1180, "mnuDocumentos", "Termo de Acordo"
   
   
   
   'MontaBotoes chd, 1210, "mnuDocumentos", "Pre�o P�blico - Guia"
   
End Sub
Private Sub MontaSubModelosGravados(ab As ActiveBar2)
Dim chd
Dim objFileSystem       As Scripting.FileSystemObject
Dim objFiles            As Scripting.files
Dim objFolder           As Scripting.Folder
Dim objFile             As Scripting.file
Dim intTextoAtual       As Integer


Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdArqSubModelosGravados"
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTPopup
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

    Set objFileSystem = New Scripting.FileSystemObject
    
    Set objFolder = objFileSystem.GetFolder(gstrDirDocumentos & "\Documentos\" & App.ProductName & "\Modelos")
       
    Set objFiles = objFolder.files
   
    intTextoAtual = 0
    
    For Each objFile In objFiles
        If Right(objFile.Name, 3) = "rpx" Then
            MontaSubArquivosGravados ab, Left(objFile.Name, Len(objFile.Name) - 4)
            MontaBotoes chd, intTextoAtual + 100, "mnuSubModelosGravados", Left(objFile.Name, Len(objFile.Name) - 4), True, "chdArqSubArquivosGravados" & Left(objFile.Name, Len(objFile.Name) - 4)
            chd.Tools(intTextoAtual).TagVariant = gstrDirDocumentos & "\Documentos\Modelos\" & App.ProductName & "\" & objFile.Name
            intTextoAtual = intTextoAtual + 1
        End If
    Next
    
End Sub


' *** TIMTIM - 09/04/2003 ***
Private Sub MontaSubWordModelosGravados(ab As ActiveBar2)
Dim chd           As Band
Dim adoRec        As ADODB.Recordset
Dim stpSQL1       As String
Dim stpSQL2       As String
Dim objFile       As Scripting.file
Dim objFiles      As Scripting.files
Dim stpFolder     As String
Dim objFolder     As Scripting.Folder
Dim objFileSystem As Scripting.FileSystemObject
Dim intTextoAtual As Integer

On Error GoTo Problema_Na_Rotina

   Set chd = ab.Bands.Add("ChildBand")
   
   With chd
       .Name = "chdArqSubWordModelosGravados"
       .DockingArea = ddDAPopup
       .Caption = " "
       .GrabHandleStyle = ddGSNone
       .Type = ddBTPopup
       .Visible = False
       .flags = 127
       .Width = 1200
      '.Left = 15000
   End With
   
   intTextoAtual = 0
   
   stpFolder = gstrDirDocumentos & "\Documentos\" & App.ProductName & "\WordModelos"

   Set objFileSystem = New Scripting.FileSystemObject
   
   If objFileSystem.FolderExists(stpFolder) Then

      Set objFolder = objFileSystem.GetFolder(stpFolder)

      Set objFiles = objFolder.files
    
      For Each objFile In objFiles
         
         If UCase$(Right(objFile.Name, 3)) = "DOT" Then
         
            stpSQL1 = "SELECT DocDescription FROM " & gstrNumeradorDocumentos & " WHERE DocDescription = '" & Left(objFile.Name, Len(objFile.Name) - 4) & "'"
            stpSQL2 = "INSERT INTO " & gstrNumeradorDocumentos & " (DocDescription) VALUES ('" & Left(objFile.Name, Len(objFile.Name) - 4) & "')"
            
            Set adoRec = gcncADOMain.Execute(stpSQL1, , adCmdText)
            
            If adoRec.EOF Then gcncADOMain.Execute stpSQL2, , adCmdText
            
            adoRec.Close
            
           'MontaBotoes chd, intTextoAtual + 2000, "mnuWordTemplate", Left(objFile.Name, Len(objFile.Name) - 4), , , , Left(objFile.Name, Len(objFile.Name) - 4) & "|" & objFile.Path
            MontaBotoes chd, intTextoAtual + 2000, "mnuWordTemplate", Left(objFile.Name, Len(objFile.Name) - 4), , , , objFile.Path
           'MontaBotoes chd, intTextoAtual, "mnuWordTemplate", Left(objFile.Name, Len(objFile.Name) - 4), True, "chdArqSubWordModelos"
            
            intTextoAtual = intTextoAtual + 1
            
         End If
      Next
    
   End If
   
   Set objFile = Nothing
   
   Set objFiles = Nothing
   
   Set objFolder = Nothing
   
   Set objFileSystem = Nothing
   
   Exit Sub

Problema_Na_Rotina:
   
'  If RecoverError("MontaSubWordModelosGravados") Then Resume
   
   ExibeDetalheErro "Erro na rotina MontaSubWordModelosGravados."
    
End Sub


Private Sub MontaSubArquivosGravados(ab As ActiveBar2, strArquivo As String)
    Dim chd                 As Object
Dim objFileSystem       As Scripting.FileSystemObject
Dim objFiles            As Scripting.files
Dim objFolder           As Scripting.Folder
Dim objFile             As Scripting.file
Dim intTextoAtual       As Integer


Set chd = ab.Bands.Add("ChildBand")
With chd
    .Name = "chdArqSubArquivosGravados" & strArquivo
    .DockingArea = ddDAPopup
    .Caption = " "
    .GrabHandleStyle = ddGSNone
    .Type = ddBTPopup
    .Visible = False
    .flags = 127
    .Width = 1200
    '.Left = 15000
End With

    Set objFileSystem = New Scripting.FileSystemObject
    
    Set objFolder = objFileSystem.GetFolder(gstrDirDocumentos & "\Documentos\" & App.ProductName & "\Gravados")
       
    Set objFiles = objFolder.files
   
    intTextoAtual = 0
   
    For Each objFile In objFiles
        If Right(objFile.Name, 3) = "rpx" And Left(objFile.Name, Len(strArquivo)) = strArquivo Then
            MontaBotoes chd, intTextoAtual + 100, "mnuSubArquivosGravados", Replace(Replace(Right(Left(objFile.Name, Len(objFile.Name) - 4), Len(objFile.Name) - Len(strArquivo) - 4), "#", ":"), "-", "/")
            chd.Tools(intTextoAtual).TagVariant = gstrDirDocumentos & "\Documentos\" & App.ProductName & "\Gravados\" & objFile.Name
            intTextoAtual = intTextoAtual + 1
        End If
    Next
    
    If intTextoAtual = 0 Then ab.Bands.Remove "chdArqSubArquivosGravados" & strArquivo
    
End Sub

Public Sub OpenWordDocumentCertidaoNegativa(strInscricao As String, strNumero As String, strLogradouro As String, STRBAIRRO As String, strLote As String, strQuadra As String, strVila As String, strProprietario As String, strIPTU As String, strVencimento As String, strData As String, strInscricaoAuxiliar As String, intUtilizacao As Integer)
                                 
                Const MODELO        As String = "CERTID�O NEGATIVA"
                Dim blpMsg          As Boolean
                Dim intFor          As Integer
                Dim stpDocument     As String
                Dim stpTemplate     As String
                Dim objFileSystem   As Scripting.FileSystemObject
                Dim stpTemplatePath As String
                Dim stpDocumentPath As String
                Dim strSQL          As String
                    
    stpTemplate = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & MODELO & ".dot"
    stpTemplatePath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\"
    stpDocumentPath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordGravados\"
   
    Set objFileSystem = New Scripting.FileSystemObject
    
    If objFileSystem.FolderExists(stpTemplatePath) Then
        If objFileSystem.FileExists(stpTemplate) Then
            If objFileSystem.FolderExists(stpDocumentPath) Then
                ReDim Documentos(1)
                stpDocument = stpDocumentPath & MODELO & "_" & strInscricao & "_" & Year(gstrDataDoSistema) & "_" & Format$(strNumero, "0000000000") & ".doc"
                                
                blpMsg = True
                
                If objFileSystem.FileExists(stpDocument) Then
                    blpMsg = (MsgBox("A " & MODELO & "_" & strInscricao & "_" & Year(gstrDataDoSistema) & "_" & Format$(strNumero, "0000000000") & " j� existe. Deseja atualiz�-la a partir do modelo original ?", vbYesNo + vbInformation, "Mensagem ao usu�rio") = vbYes)
                    If blpMsg Then
                        objFileSystem.DeleteFile stpDocument, True
                    End If
                End If
                            
                Set Documentos(1) = New cWordWrapper
                Documentos(1).GetContainer
                Documentos(1).DocumentTemplatePath = stpTemplate
                Documentos(1).DocumentPath = stpDocument
                Documentos(1).DocumentFormat = WORDOPENFORMATDOCUMENT
                Documentos(1).DocumentOpen
                                                
                'Substitui��o dos Campos
                
                Documentos(1).DocumentReplaceField "|Inscri��o|", gstrFormataInscricao(Right(strInscricao, gintRetornaTamanhoMascara(CByte(intUtilizacao))), intUtilizacao)
                Documentos(1).DocumentReplaceField "|Inscri��o Auxiliar|", strInscricaoAuxiliar
                Documentos(1).DocumentReplaceField "|Numero|", Format$(strNumero, "000000000")
                Documentos(1).DocumentReplaceField "|Logradouro|", strLogradouro
                Documentos(1).DocumentReplaceField "|Bairro|", STRBAIRRO
                Documentos(1).DocumentReplaceField "|Lote|", strLote
                Documentos(1).DocumentReplaceField "|Quadra|", strQuadra
                Documentos(1).DocumentReplaceField "|Vila|", strVila
                Documentos(1).DocumentReplaceField "|Propriet�rio|", strProprietario
                Documentos(1).DocumentReplaceField "|IPTU|", strIPTU
                Documentos(1).DocumentReplaceField "|Vencimento|", DateAdd("d", 30, strVencimento)
                Documentos(1).DocumentReplaceField "|Vencimento60|", DateAdd("d", 60, strVencimento)
                Documentos(1).DocumentReplaceField "|Vencimento90|", DateAdd("d", 90, strVencimento)
                Documentos(1).DocumentReplaceField "|Data|", gstrDataPorExtenso(strData)
                Documentos(1).DocumentSave
            Else
                MsgBox "A pasta de documentos : " & stpDocumentPath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
            End If
        
        Else
            MsgBox "O modelo de documento do Microsoft Word : " & stpTemplate & " n�o foi localizado. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
        End If
    
    Else
       MsgBox "A pasta de modelos de documentos : " & stpTemplatePath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
    End If
    
    Set objFileSystem = Nothing

End Sub

'Feito por Hugo
Public Sub ImprimeGuiaFebraban(vetGuias() As String, _
                               strQuadra As String, _
                               strLote As String, _
                               dtmDataVencimento As String, _
                               vetParecelas() As String)
                               
Dim intFor               As Integer

Dim strCodBarras         As String
Dim adoResultado         As ADODB.Recordset
Dim adoCommand           As ADODB.Command

Dim strSQL               As String

Dim lngGuias             As Long

Dim intFebraban          As Integer
Dim INTNUMERO            As Long
Dim strNumeroBoleto      As String
               
Dim vetGuiaArrecadacao() As String
    
On Error GoTo Problema_Na_Rotina

    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans

    ReDim vetGuiaArrecadacao(19, 0)
    
    'Query utilizada para pegar o Codigo Febraban da tblEmpresa
    strSQL = ""
    strSQL = strSQL & "Select * From " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            If gstrENulo(adoResultado!intFebraban) <> "" Then
                intFebraban = gstrENulo(adoResultado!intFebraban)
            Else
                ExibeMensagem "C�digo Febraban n�o encontrado."
            End If
        Else
            ExibeMensagem "C�digo Febraban n�o encontrado."
            Exit Sub
        End If
    End If
    
ProximoNumeroGuia:
        
    INTNUMERO = glngRetornaProximoNumeroGuia
    If Val(INTNUMERO) = 0 Then
        Exit Sub
    End If
        
    'Vamos montar o codigo de barras
    strCodBarras = gstrMontaCodigoBarras(FEBRABAN, Val(vetGuias(12, 1)), vetGuias(9, 1), dtmDataVencimento, intFebraban, INTNUMERO, False, True)
    If Len(strCodBarras) = 0 Then Exit Sub
    
    strNumeroBoleto = gstrMontaLinhaDigitavel(FEBRABAN, strCodBarras)
    
    'Vamos inserir a guia na tabela TblGuias
    strSQL = ""
    'strSQL = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    strSQL = strSQL & "Insert Into " & gstrGuias & "("
    'strSQL = strSQL & "Pkid, "
    strSQL = strSQL & "Intcontabancaria, "
    strSQL = strSQL & "Intnumero, "
    strSQL = strSQL & "Dtmdtemissao, "
    strSQL = strSQL & "Dblvalor, "
    strSQL = strSQL & "Strcodbarra, "
    strSQL = strSQL & "Dtmdtatualizacao, "
    strSQL = strSQL & "Lngcodusr, "
    strSQL = strSQL & "Dtmdtvencimento "
    strSQL = strSQL & ") Values("
    'strSQL = strSQL & lngGuias & ", "
    strSQL = strSQL & "Null, "
    strSQL = strSQL & INTNUMERO & ", "
    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSQL = strSQL & gstrConvVrParaSql(vetGuias(9, 1)) & ", '"
    strSQL = strSQL & strCodBarras & "', "
    strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
    strSQL = strSQL & glngCodUsr & ", "
    strSQL = strSQL & gstrConvDtParaSql(dtmDataVencimento)
    strSQL = strSQL & ")"
    'strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " ; ", "")
    
    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = gcncADOMain
    adoCommand.CommandText = strSQL
    adoCommand.Execute strSQL, , adExecuteNoRecords
    
    lngGuias = glngRetornaPkidTabelaPai("seqTblGuias", gstrGuias)
    
    strSQL = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    
    'Vamos inserir as parcelas na tabela TblLancamentoGuias
    For intFor = 1 To UBound(vetParecelas(), 2)
        If Len(vetParecelas(13, intFor)) = 0 Then
            strSQL = strSQL & "Insert Into " & gstrLancamentoGuias & "("
            strSQL = strSQL & "intlancamentovalor, "
            strSQL = strSQL & "intguias, "
            strSQL = strSQL & "dblvalorprincipal, "
            strSQL = strSQL & "dblvalormulta, "
            strSQL = strSQL & "dblvalorjuros, "
            strSQL = strSQL & "dblvalorcorrecao, "
            strSQL = strSQL & "dblvalordesconto, "
            strSQL = strSQL & "dtmdtatualizacao, "
            strSQL = strSQL & "lngcodusr) "
            strSQL = strSQL & "Values ("
            strSQL = strSQL & vetParecelas(0, intFor) & ", "
            strSQL = strSQL & lngGuias & ","
            strSQL = strSQL & gstrConvVrParaSql(vetParecelas(1, intFor)) & ", "
            strSQL = strSQL & gstrConvVrParaSql(vetParecelas(2, intFor)) & ", "
            strSQL = strSQL & gstrConvVrParaSql(vetParecelas(3, intFor)) & ", "
            strSQL = strSQL & gstrConvVrParaSql(vetParecelas(4, intFor)) & ", "
            strSQL = strSQL & gstrConvVrParaSql("0") & ", "
            strSQL = strSQL & strGETDATE & ", "
            strSQL = strSQL & glngCodUsr & ") "
            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), ";", "")
        End If
    Next
    
    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
    
    If gobjBanco.Execute(strSQL, True) Then
        gobjBanco.ExecutaCommitTrans
    Else
        ExibeMensagem "Erro na grava��o dos lan�amentos da guia."
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If
    
    vetGuiaArrecadacao(0, 0) = INTNUMERO & "/" & Year(gstrDataDoSistema)
    vetGuiaArrecadacao(1, 0) = dtmDataVencimento
    vetGuiaArrecadacao(2, 0) = vetGuias(0, 1)
    vetGuiaArrecadacao(3, 0) = vetGuias(1, 1)
    vetGuiaArrecadacao(4, 0) = vetGuias(2, 1)
    vetGuiaArrecadacao(5, 0) = strQuadra
    vetGuiaArrecadacao(6, 0) = vetGuias(11, 1)
    vetGuiaArrecadacao(7, 0) = vetGuias(3, 1)
    vetGuiaArrecadacao(8, 0) = vetGuias(4, 1)
    vetGuiaArrecadacao(9, 0) = vetGuias(10, 1)
    vetGuiaArrecadacao(10, 0) = gstrConvVrDoSql(vetGuias(5, 1))
    vetGuiaArrecadacao(11, 0) = gstrConvVrDoSql(vetGuias(8, 1))
    vetGuiaArrecadacao(12, 0) = gstrConvVrDoSql(vetGuias(6, 1))
    vetGuiaArrecadacao(13, 0) = gstrConvVrDoSql(vetGuias(7, 1))
    vetGuiaArrecadacao(14, 0) = gstrConvVrDoSql(vetGuias(9, 1))
    vetGuiaArrecadacao(15, 0) = gstrDataDoSistema
    vetGuiaArrecadacao(16, 0) = gstrLoginUser
    vetGuiaArrecadacao(17, 0) = dtmDataVencimento
    vetGuiaArrecadacao(18, 0) = strNumeroBoleto
    vetGuiaArrecadacao(19, 0) = strCodBarras
        
    'Vamos imprimir o relatorio de guia de arrecadacao
    If Not IsNull(vetGuiaArrecadacao(0, 0)) Then
        ImprimeRelatorioPorArray rptGuiaDeArrecadacao, vetGuiaArrecadacao, "Guia de Arrecada��o"
    End If
    
    Exit Sub
    
Problema_Na_Rotina:
   
  If InStr(1, UCase(Err.Description), "UK_TBLGUIAS_INTNUMERODTEMISSAO") > 0 Then
      GoTo ProximoNumeroGuia
  Else
      ExibeDetalheErro "Erro na rotina ImprimeGuiaFebraban."
      gobjBanco.ExecutaRollbackTrans
  End If
   
End Sub

Public Sub ImprimeGuiaFichaCompensacao(vetGuias() As String, _
                                        strQuadra As String, _
                                        strLote As String, _
                                        dtmDataVencimento As String, _
                                        vetParecelas() As String)
                               
Dim intForGuia           As Integer
Dim intFor               As Integer

Dim strCodBarras         As String
Dim adoResultado         As ADODB.Recordset
Dim adoCommand           As ADODB.Command

Dim strSQL               As String

Dim lngGuias             As Long

Dim intFebraban          As Integer
Dim INTNUMERO            As Long
Dim strNumeroBoleto      As String
Dim strNossoNumero       As String

Dim vetGuiaArrecadacao() As String
    
On Error GoTo Problema_Na_Rotina

    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans

    ReDim vetGuiaArrecadacao(21, 0)
    
    'Query utilizada para pegar o Codigo Febraban da tblEmpresa
    strSQL = ""
    strSQL = strSQL & "Select * From " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount > 0 Then
            If gstrENulo(adoResultado!intFebraban) <> "" Then
                intFebraban = gstrENulo(adoResultado!intFebraban)
            Else
                ExibeMensagem "C�digo Febraban n�o encontrado."
            End If
        Else
            ExibeMensagem "C�digo Febraban n�o encontrado."
            Exit Sub
        End If
    End If
    
    For intForGuia = 1 To UBound(vetGuias(), 2)
        
        If Len(vetGuiaArrecadacao(0, 0)) > 0 Then
            ReDim Preserve vetGuiaArrecadacao(21, UBound(vetGuiaArrecadacao, 2) + 1)
        End If
        
ProximoNumeroGuia:
        
        INTNUMERO = glngRetornaProximoNumeroGuia
        If Val(INTNUMERO) = 0 Then
            Exit Sub
        End If
        
        strCodBarras = gstrMontaCodigoBarras(FICHA_COMPENSACAO, Val(vetGuias(12, intForGuia)), vetGuias(9, intForGuia), dtmDataVencimento, intFebraban, INTNUMERO, False, True)
        If Len(strCodBarras) = 0 Then Exit Sub
        
        strNumeroBoleto = gstrMontaLinhaDigitavel(FICHA_COMPENSACAO, strCodBarras)

        strNossoNumero = gstrMontaNossoNumero(Val(vetGuias(12, intForGuia)), INTNUMERO)

        'Vamos inserir a guia na tabela TblGuias
        strSQL = ""
        'strSQL = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
        strSQL = strSQL & "Insert Into " & gstrGuias & "("
        'strSQL = strSQL & "Pkid, "
        strSQL = strSQL & "Intcontabancaria, "
        strSQL = strSQL & "Intnumero, "
        strSQL = strSQL & "Dtmdtemissao, "
        strSQL = strSQL & "Dblvalor, "
        strSQL = strSQL & "Strcodbarra, "
        strSQL = strSQL & "Dtmdtatualizacao, "
        strSQL = strSQL & "Lngcodusr, "
        strSQL = strSQL & "Dtmdtvencimento "
        strSQL = strSQL & ") Values("
        'strSQL = strSQL & lngGuias & ", "
        strSQL = strSQL & Val(vetGuias(12, intForGuia)) & ", "
        strSQL = strSQL & INTNUMERO & ", "
        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
        strSQL = strSQL & gstrConvVrParaSql(vetGuias(9, intForGuia)) & ", '"
        strSQL = strSQL & strCodBarras & "', "
        strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
        strSQL = strSQL & glngCodUsr & ", "
        strSQL = strSQL & gstrConvDtParaSql(dtmDataVencimento)
        strSQL = strSQL & ")"
        'strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " ; ", "")
        
        Set adoCommand = New ADODB.Command
        Set adoCommand.ActiveConnection = gcncADOMain
        adoCommand.CommandText = strSQL
        adoCommand.Execute strSQL, , adExecuteNoRecords
        
        lngGuias = glngRetornaPkidTabelaPai("seqTblGuias", gstrGuias)
        
        strSQL = IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
        
        'Vamos inserir as parcelas na tabela TblLancamentoGuias
        For intFor = 1 To UBound(vetParecelas(), 2)
            If vetGuias(12, intForGuia) = vetParecelas(13, intFor) Then
                strSQL = strSQL & "Insert Into " & gstrLancamentoGuias & "("
                strSQL = strSQL & "intlancamentovalor, "
                strSQL = strSQL & "intguias, "
                strSQL = strSQL & "dblvalorprincipal, "
                strSQL = strSQL & "dblvalormulta, "
                strSQL = strSQL & "dblvalorjuros, "
                strSQL = strSQL & "dblvalorcorrecao, "
                strSQL = strSQL & "dblvalordesconto, "
                strSQL = strSQL & "dtmdtatualizacao, "
                strSQL = strSQL & "lngcodusr) "
                strSQL = strSQL & "Values ("
                strSQL = strSQL & vetParecelas(0, intFor) & ", "
                strSQL = strSQL & lngGuias & ","
                strSQL = strSQL & gstrConvVrParaSql(vetParecelas(1, intFor)) & ", "
                strSQL = strSQL & gstrConvVrParaSql(vetParecelas(2, intFor)) & ", "
                strSQL = strSQL & gstrConvVrParaSql(vetParecelas(3, intFor)) & ", "
                strSQL = strSQL & gstrConvVrParaSql(vetParecelas(4, intFor)) & ", "
                strSQL = strSQL & gstrConvVrParaSql("0") & ", "
                strSQL = strSQL & strGETDATE & ", "
                strSQL = strSQL & glngCodUsr & ") "
                strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), ";", "")
            End If
        Next
        
        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
        
        If gobjBanco.Execute(strSQL, True) Then
            gobjBanco.ExecutaCommitTrans
        Else
            ExibeMensagem "Erro na grava��o dos lan�amentos da guia."
            gobjBanco.ExecutaRollbackTrans
            Exit Sub
        End If
        
        vetGuiaArrecadacao(0, intForGuia - 1) = INTNUMERO & "/" & Year(gstrDataDoSistema)
        vetGuiaArrecadacao(1, intForGuia - 1) = dtmDataVencimento
        vetGuiaArrecadacao(2, intForGuia - 1) = vetGuias(0, intForGuia)
        vetGuiaArrecadacao(3, intForGuia - 1) = vetGuias(1, intForGuia)
        vetGuiaArrecadacao(4, intForGuia - 1) = vetGuias(2, intForGuia)
        vetGuiaArrecadacao(5, intForGuia - 1) = strQuadra
        vetGuiaArrecadacao(6, intForGuia - 1) = vetGuias(11, intForGuia)
        vetGuiaArrecadacao(7, intForGuia - 1) = vetGuias(3, intForGuia)
        vetGuiaArrecadacao(8, intForGuia - 1) = vetGuias(4, intForGuia)
        vetGuiaArrecadacao(9, intForGuia - 1) = vetGuias(10, intForGuia)
        vetGuiaArrecadacao(10, intForGuia - 1) = gstrConvVrDoSql(vetGuias(5, intForGuia))
        vetGuiaArrecadacao(11, intForGuia - 1) = gstrConvVrDoSql(vetGuias(8, intForGuia))
        vetGuiaArrecadacao(12, intForGuia - 1) = gstrConvVrDoSql(vetGuias(6, intForGuia))
        vetGuiaArrecadacao(13, intForGuia - 1) = gstrConvVrDoSql(vetGuias(7, intForGuia))
        vetGuiaArrecadacao(14, intForGuia - 1) = gstrConvVrDoSql(vetGuias(9, intForGuia))
        vetGuiaArrecadacao(15, intForGuia - 1) = gstrDataDoSistema
        vetGuiaArrecadacao(16, intForGuia - 1) = gstrLoginUser
        vetGuiaArrecadacao(17, intForGuia - 1) = dtmDataVencimento
        vetGuiaArrecadacao(18, intForGuia - 1) = strNumeroBoleto
        vetGuiaArrecadacao(19, intForGuia - 1) = strCodBarras
        vetGuiaArrecadacao(20, intForGuia - 1) = vetGuias(12, intForGuia)
        vetGuiaArrecadacao(21, intForGuia - 1) = strNossoNumero
        
    Next
    
    'Vamos imprimir o relatorio de guia de arrecadacao
    If Not IsNull(vetGuiaArrecadacao(0, 0)) Then
        ImprimeRelatorioPorArray rptGuiaFichaDeArrecadacao, vetGuiaArrecadacao, "Guia de Arrecada��o"
    End If
    
    Exit Sub
    
Problema_Na_Rotina:
   
  If InStr(1, UCase(Err.Description), "UK_TBLGUIAS_INTNUMERODTEMISSAO") > 0 Then
      GoTo ProximoNumeroGuia
  Else
      ExibeDetalheErro "Erro na rotina ImprimeGuiaFichaCompensacao."
      gobjBanco.ExecutaRollbackTrans
  End If
   
End Sub

'Feito por Hugo
Public Sub OpenWordDocumentCertidaoPositiva(strInscricao As String, strAtividade As String, strNumeroProcesso As String, strNumero As String, strLogradouro As String, strProprietario As String, strVencimento As String, XArrayTabela As XArrayDB, XArrayAlinhaColunas As XArrayDB, dblTotal As Double, strInscricaoAuxiliar As String, intUtilizacao As Integer)

Const MODELO        As String = "CERTID�O POSITIVA"
Dim blpMsg          As Boolean
Dim intFor          As Integer
Dim stpDocument     As String
Dim stpTemplate     As String
Dim objFileSystem   As Scripting.FileSystemObject
Dim stpTemplatePath As String
Dim stpDocumentPath As String
Dim strSQL          As String
                                    
    stpTemplate = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & MODELO & ".dot"
    stpTemplatePath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\"
    stpDocumentPath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordGravados\"
   
    Set objFileSystem = New Scripting.FileSystemObject
    
    If objFileSystem.FolderExists(stpTemplatePath) Then
        If objFileSystem.FileExists(stpTemplate) Then
            If objFileSystem.FolderExists(stpDocumentPath) Then
                ReDim Documentos(1)
                stpDocument = stpDocumentPath & MODELO & "_" & strInscricao & "_" & Year(gstrDataDoSistema) & "_" & Format$(strNumero, "0000000000") & ".doc"
                                
                blpMsg = True
                
                If objFileSystem.FileExists(stpDocument) Then
                    blpMsg = (MsgBox("A " & MODELO & "_" & strInscricao & "_" & Year(gstrDataDoSistema) & "_" & Format$(strNumero, "0000000000") & " j� existe. Deseja atualiz�-la a partir do modelo original ?", vbYesNo + vbInformation, "Mensagem ao usu�rio") = vbYes)
                    If blpMsg Then
                        objFileSystem.DeleteFile stpDocument, True
                    End If
                End If
                            
                Set Documentos(1) = New cWordWrapper
                Documentos(1).GetContainer
                Documentos(1).DocumentTemplatePath = stpTemplate
                Documentos(1).DocumentPath = stpDocument
                Documentos(1).DocumentFormat = WORDOPENFORMATDOCUMENT
                Documentos(1).DocumentOpen
                                                
                'Substitui��o dos Campos
                Documentos(1).DocumentReplaceField "|Inscri��o Municipal|", gstrFormataInscricao(Right(strInscricao, gintRetornaTamanhoMascara(CByte(intUtilizacao))), intUtilizacao)
                Documentos(1).DocumentReplaceField "|Inscri��o Auxiliar|", strInscricaoAuxiliar 'Tri0842
                Documentos(1).DocumentReplaceField "|Contribuinte|", strProprietario
                Documentos(1).DocumentReplaceField "|N�mero Certid�o|", strNumero
                Documentos(1).DocumentReplaceField "|Processo|", strNumeroProcesso
                Documentos(1).DocumentReplaceField "|Data|", Day(gstrDataDoSistema) & " de " & gstrNomeDoMes(Month(gstrDataDoSistema)) & " de " & Year(gstrDataDoSistema)
                Documentos(1).DocumentReplaceField "|Local|", strLogradouro
                Documentos(1).DocumentReplaceField "|Atividade|", strAtividade
                Documentos(1).DocumentReplaceField "|Validade|", gstrDataFormatada(DateAdd("D", 30, gstrDataDoSistema))
                Documentos(1).DocumentReplaceField "|Validade60|", gstrDataFormatada(DateAdd("D", 60, gstrDataDoSistema))
                Documentos(1).DocumentReplaceField "|Validade90|", gstrDataFormatada(DateAdd("D", 90, gstrDataDoSistema))
                Documentos(1).DocumentInsert "|Tabela|", , XArrayTabela, XArrayAlinhaColunas
                Documentos(1).DocumentReplaceField "|Total|", gstrConvVrDoSql(dblTotal) & " ( " & gstrExtenso(dblTotal, 0) & " )."
                Documentos(1).DocumentSave
            Else
                MsgBox "A pasta de documentos : " & stpDocumentPath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
            End If
        
        Else
            MsgBox "O modelo de documento do Microsoft Word : " & stpTemplate & " n�o foi localizado. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
        End If
    
    Else
       MsgBox "A pasta de modelos de documentos : " & stpTemplatePath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
    End If
    
    Set objFileSystem = Nothing

End Sub

Public Sub OpenWordDocumentCertidaoPositivaNegativo(strInscricao As String, strNumeroProcesso As String, strNumero As String, strLogradouro As String, STRBAIRRO As String, STRMUNICIPIO As String, STRUF As String, strProprietario As String, strVencimento As String, strInscricaoAuxiliar As String, intUtilizacao As Integer)
Const MODELO        As String = "CERTID�O POSITIVA NEGATIVO"
Dim blpMsg          As Boolean
Dim intFor          As Integer
Dim stpDocument     As String
Dim stpTemplate     As String
Dim objFileSystem   As Scripting.FileSystemObject
Dim stpTemplatePath As String
Dim stpDocumentPath As String
Dim strSQL          As String
                                    
    stpTemplate = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & MODELO & ".dot"
    stpTemplatePath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\"
    stpDocumentPath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordGravados\"
   
    Set objFileSystem = New Scripting.FileSystemObject
    
    If objFileSystem.FolderExists(stpTemplatePath) Then
        If objFileSystem.FileExists(stpTemplate) Then
            If objFileSystem.FolderExists(stpDocumentPath) Then
                ReDim Documentos(1)
                stpDocument = stpDocumentPath & MODELO & "_" & strInscricao & "_" & Year(gstrDataDoSistema) & "_" & Format$(strNumero, "0000000000") & ".doc"
                                
                blpMsg = True
                
                If objFileSystem.FileExists(stpDocument) Then
                    blpMsg = (MsgBox("A " & MODELO & "_" & strInscricao & "_" & Year(gstrDataDoSistema) & "_" & Format$(strNumero, "0000000000") & " j� existe. Deseja atualiz�-la a partir do modelo original ?", vbYesNo + vbInformation, "Mensagem ao usu�rio") = vbYes)
                    If blpMsg Then
                        objFileSystem.DeleteFile stpDocument, True
                    End If
                End If
                            
                Set Documentos(1) = New cWordWrapper
                Documentos(1).GetContainer
                Documentos(1).DocumentTemplatePath = stpTemplate
                Documentos(1).DocumentPath = stpDocument
                Documentos(1).DocumentFormat = WORDOPENFORMATDOCUMENT
                Documentos(1).DocumentOpen
                                                
                'Substitui��o dos Campos
                Documentos(1).DocumentReplaceField "|Inscri��o Municipal|", gstrFormataInscricao(Right(strInscricao, gintRetornaTamanhoMascara(CByte(intUtilizacao))), intUtilizacao)
                Documentos(1).DocumentReplaceField "|Inscri��o Auxiliar|", strInscricaoAuxiliar
                Documentos(1).DocumentReplaceField "|Contribuinte|", strProprietario
                Documentos(1).DocumentReplaceField "|Processo|", strNumeroProcesso
                Documentos(1).DocumentReplaceField "|Data|", Day(gstrDataDoSistema) & " de " & gstrNomeDoMes(Month(gstrDataDoSistema)) & " de " & Year(gstrDataDoSistema)
                Documentos(1).DocumentReplaceField "|Local|", strLogradouro
                Documentos(1).DocumentReplaceField "|Local2|", STRBAIRRO & " - " & STRMUNICIPIO & " - " & STRUF
                Documentos(1).DocumentReplaceField "|Validade|", gstrDataFormatada(DateAdd("D", 30, gstrDataDoSistema))
                Documentos(1).DocumentReplaceField "|Validade60|", gstrDataFormatada(DateAdd("D", 60, gstrDataDoSistema))
                Documentos(1).DocumentReplaceField "|Validade90|", gstrDataFormatada(DateAdd("D", 90, gstrDataDoSistema))
                Documentos(1).DocumentSave
            Else
                MsgBox "A pasta de documentos : " & stpDocumentPath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
            End If
        
        Else
            MsgBox "O modelo de documento do Microsoft Word : " & stpTemplate & " n�o foi localizado. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
        End If
    
    Else
       MsgBox "A pasta de modelos de documentos : " & stpTemplatePath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
    End If
    
    Set objFileSystem = Nothing

End Sub

'Feito por Hugo
Public Sub OpenWordDocumentTermoDeAcordo(strInscricaoAcordo As String, _
                                        strAnoAcordo As String, _
                                        strDiaAcordo As String, _
                                        strContribuinteAcordo As String, _
                                        strCPFAcordo As String, _
                                        strRGAcordo As String, _
                                        strEndContribuinteAcordo As String, _
                                        strInscricaoRepresentante As String, _
                                        strContribuinteRepresentante As String, _
                                        strEndContribuinteRepresentante As String, _
                                        strMoeda As String, _
                                        strUsuario As String, _
                                        Strindexador As String, _
                                        dblvlIndexador As Double, _
                                        DblvlParcela As Double, _
                                        intQtdeParcelasAcordo As Integer, _
                                        DblvlTotal As Double, _
                                        XArrayTabela As XArrayDB, _
                                        XArrayAlinhaColunas As XArrayDB)

Const MODELO        As String = "TERMO DE ACORDO"
Dim blpMsg          As Boolean
Dim stpDocument     As String
Dim stpTemplate     As String
Dim objFileSystem   As Scripting.FileSystemObject
Dim stpTemplatePath As String
Dim stpDocumentPath As String
Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
Dim strComposicoes  As String
Dim intFor          As Integer

    stpTemplate = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & MODELO & ".dot"
    stpTemplatePath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\"
    stpDocumentPath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordGravados\"
   
    Set objFileSystem = New Scripting.FileSystemObject
    
    If objFileSystem.FolderExists(stpTemplatePath) Then
        If objFileSystem.FileExists(stpTemplate) Then
            If objFileSystem.FolderExists(stpDocumentPath) Then
                ReDim Documentos(1)
                stpDocument = stpDocumentPath & MODELO & "_" & Trim(Replace(strInscricaoAcordo, "/", "")) & "_" & (strAnoAcordo) & ".doc"
                blpMsg = True
                
                If objFileSystem.FileExists(stpDocument) Then
                
                    blpMsg = (MsgBox("O Termo de Acordo: " & MODELO & "_" & Trim(strInscricaoAcordo) & "_" & (strAnoAcordo) & " j� existe. Deseja atualiz�-la a partir do modelo original ?", vbYesNo + vbInformation, "Mensagem ao usu�rio") = vbYes)
                
                    If blpMsg Then
                        objFileSystem.DeleteFile stpDocument, True
                    End If
                End If
                            
                Set Documentos(1) = New cWordWrapper
                Documentos(1).GetContainer
                Documentos(1).DocumentTemplatePath = stpTemplate
                Documentos(1).DocumentPath = stpDocument
                Documentos(1).DocumentFormat = WORDOPENFORMATDOCUMENT
                Documentos(1).DocumentOpen
                                                
                'Substitui��o dos Campos
                Documentos(1).DocumentReplaceField "|Inscri��oAno|", strInscricaoAcordo
                Documentos(1).DocumentReplaceField "|Dia do Acordo|", Day(strDiaAcordo) & " dia(s) do m�s de " & gstrNomeDoMes(Month(strDiaAcordo)) & " de " & Year(strDiaAcordo)
                Documentos(1).DocumentReplaceField "|Contribuinte|", strContribuinteAcordo
                Documentos(1).DocumentReplaceField "|CPF|", strCPFAcordo
                Documentos(1).DocumentReplaceField "|RG|", strRGAcordo
                Documentos(1).DocumentReplaceField "|Endere�oAcordo|", strEndContribuinteAcordo
                Documentos(1).DocumentReplaceField "|Inscri��o|", strInscricaoRepresentante
                Documentos(1).DocumentReplaceField "|Empresa|", strContribuinteRepresentante
                Documentos(1).DocumentReplaceField "|Endere�o|", strEndContribuinteRepresentante
                Documentos(1).DocumentReplaceField "|Sigla|", strMoeda
                Documentos(1).DocumentReplaceField "|NumeroParcelas|", intQtdeParcelasAcordo
                
                If Trim(Strindexador) <> "" Then
                    Documentos(1).DocumentReplaceField "|ValorParcela|", gstrConvVrDoSql(CDbl(DblvlParcela) / dblvlIndexador, 4, , True)
                    Documentos(1).DocumentReplaceField "|IndexadorMoeda|", Strindexador
                    Documentos(1).DocumentReplaceField "|ValorTotal|", gstrConvVrDoSql(CDbl(DblvlTotal) / dblvlIndexador, 4, , True)
                Else
                    Documentos(1).DocumentReplaceField "|ValorParcela|", gstrConvVrDoSql(DblvlParcela)
                    Documentos(1).DocumentReplaceField "|IndexadorMoeda|", strMoeda
                    Documentos(1).DocumentReplaceField "|ValorTotal|", gstrConvVrDoSql(DblvlTotal)
                End If
                
                                
                Documentos(1).DocumentReplaceField "|ValorTotal1|", gstrConvVrDoSql(DblvlTotal)
                Documentos(1).DocumentReplaceField "|ValorParcela1|", gstrConvVrDoSql(DblvlParcela)
                
                strSQL = ""
                strSQL = strSQL & "SELECT EP.PKId, MU.strDescricao, EP.strCGC "
                strSQL = strSQL & "FROM " & gstrCidade & " MU, "
                strSQL = strSQL & gstrEmpresa & " EP "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & " MU.PKId = EP.intCidade "
                
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    With adoResultado
                        Do While Not .EOF
                            Documentos(1).DocumentReplaceField "|data|", gstrENulo(!strDescricao) + ", " + gstrDataPorExtenso(gstrDataDoSistema) + "."
                            Documentos(1).DocumentReplaceField "|CGCEmpresa|", gstrENulo(!strCGC)
                            .MoveNext
                        Loop
                    End With
                End If
                
                strSQL = ""
                strSQL = strSQL & "SELECT LV.dtmDtVencimento "
                strSQL = strSQL & "FROM " & gstrLancamentoAlfa & " LA, "
                strSQL = strSQL & gstrLancamentoValor & " LV "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & " LA.strInscricao = '" & String(gintLenInscricao - Len(Replace(strInscricaoAcordo, "/", "")), "0") & Replace(strInscricaoAcordo, "/", "") & "'"
                strSQL = strSQL & " AND LV.intLancamentoAlfa = La.Pkid "
                strSQL = strSQL & " ORDER BY LV.intParcela DESC "
                
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    With adoResultado
                        Documentos(1).DocumentReplaceField "|VenctoUltimaParcela|", gstrENulo(!Dtmdtvencimento)
                    End With
                End If
                
                For intFor = 1 To XArrayTabela.UpperBound(1)
                    If InStr(1, strComposicoes, XArrayTabela(intFor, 0) & "/" & XArrayTabela(intFor, 1)) = 0 Then
                        strComposicoes = strComposicoes & " " & XArrayTabela(intFor, 0) & "/" & XArrayTabela(intFor, 1)
                    End If
                Next
                Documentos(1).DocumentReplaceField "|Composicoes|", strComposicoes
                
                Documentos(1).DocumentReplaceField "|Usu�rio|", strUsuario
                Documentos(1).DocumentReplaceField "|Usu�rioImpressao|", gstrNomeUsuario
                Documentos(1).DocumentReplaceField "|DataAtual|", Day(gstrDataDoSistema) & " de " & gstrNomeDoMes(Month(gstrDataDoSistema)) & " de " & Year(gstrDataDoSistema)
                Documentos(1).DocumentInsert "|Tabela|", , XArrayTabela, XArrayAlinhaColunas
                Documentos(1).DocumentSave
            Else
                MsgBox "A pasta de documentos : " & stpDocumentPath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
            End If
        
        Else
            MsgBox "O modelo de documento do Microsoft Word : " & stpTemplate & " n�o foi localizado. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
        End If
    
    Else
       MsgBox "A pasta de modelos de documentos : " & stpTemplatePath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
    End If
    
    Set objFileSystem = Nothing

End Sub

Public Sub OpenWordDocumentCertidaoMobiliario(strContMobiliario As String, _
                                        strInscricaoCadastral As String, _
                                        strRazaoSocial As String, _
                                        strDataCadastroContribuintes As String, _
                                        strSiglaLogradouro As String, _
                                        strLogradouro As String, _
                                        INTNUMERO As String, _
                                        STRBAIRRO As String, _
                                        strNumeroProcesso As String, _
                                        strdetalheHorario As String, _
                                        XArrayTabela As XArrayDB, _
                                        XArrayAlinhaColunas As XArrayDB)
                                        
                Const MODELO                As String = "CERTIDAO CADASTRO MOBILIARIO"
                Dim blpMsg                  As Boolean
                Dim stpDocument             As String
                Dim stpTemplate             As String
                Dim objFileSystem           As Scripting.FileSystemObject
                Dim stpTemplatePath         As String
                Dim stpDocumentPath         As String
                Dim strSQL                  As String

                                    
    stpTemplate = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & MODELO & ".dot"
    stpTemplatePath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\"
    stpDocumentPath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordGravados\"
   
    Set objFileSystem = New Scripting.FileSystemObject
    If objFileSystem.FolderExists(stpTemplatePath) Then
        If objFileSystem.FileExists(stpTemplate) Then
            If objFileSystem.FolderExists(stpDocumentPath) Then
                ReDim Documentos(1)
                stpDocument = stpDocumentPath & MODELO & "_" & gstrFormataInscricao(strInscricaoCadastral, TYP_ECONOMICA) & "_" & Year(gstrDataDoSistema) & "_" & Format$(strContMobiliario, "0000000000") & ".doc"
                blpMsg = True
                If objFileSystem.FileExists(stpDocument) Then
                    blpMsg = (MsgBox("A Certid�o Mobiliaria:" & MODELO & gstrFormataInscricao(strInscricaoCadastral, TYP_ECONOMICA) & "_" & Year(gstrDataDoSistema) & "_" & Format$(strContMobiliario, "0000000000") & " j� existe. Deseja atualiz�-la a partir do modelo original ?", vbYesNo + vbInformation, "Mensagem ao usu�rio") = vbYes)
                    If blpMsg Then
                        objFileSystem.DeleteFile stpDocument, True
                    End If
                End If
                                        
                Set Documentos(1) = New cWordWrapper
                Documentos(1).GetContainer
                Documentos(1).DocumentTemplatePath = stpTemplate
                Documentos(1).DocumentPath = stpDocument
                Documentos(1).DocumentFormat = WORDOPENFORMATDOCUMENT
                Documentos(1).DocumentOpen
                    
                Documentos(1).DocumentReplaceField "|Inscri��o Cadastral|", gstrFormataInscricao(strInscricaoCadastral, TYP_ECONOMICA)
                Documentos(1).DocumentReplaceField "|Nome do Contribuinte|", strRazaoSocial
                Documentos(1).DocumentReplaceField "|Data Cadastro|", strDataCadastroContribuintes
                Documentos(1).DocumentReplaceField "|Tipo Logradouro|", strSiglaLogradouro
                Documentos(1).DocumentReplaceField "|Logradouro|", strLogradouro
                Documentos(1).DocumentReplaceField "|numero|", INTNUMERO
                Documentos(1).DocumentReplaceField "|Bairro|", STRBAIRRO
                Documentos(1).DocumentReplaceField "|Numero Processo|", strNumeroProcesso
                Documentos(1).DocumentReplaceField "|Descri��o hor�rio|", strdetalheHorario
                Documentos(1).DocumentReplaceField "|Data Atual|", gstrDataPorExtenso
                Documentos(1).DocumentInsert "|Tabela|", , XArrayTabela, XArrayAlinhaColunas
                Documentos(1).DocumentSave
            Else
                MsgBox "A pasta de documentos : " & stpDocumentPath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
            End If

        Else
            MsgBox "O modelo de documento do Microsoft Word : " & stpTemplate & " n�o foi localizado. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
        End If
    Else
        MsgBox "A pasta de modelos de documentos : " & stpTemplatePath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
    End If
    
    Set objFileSystem = Nothing

End Sub

Public Sub OpenWordDocumentAlvaraFuncionamento(strContAlvaraFuncionamento As String, _
                                        strDataDeAbertura As String, _
                                        strDataVencimento As String, _
                                        strInscricaoCadastral As String, _
                                        strRazaoSocial As String, _
                                        strSiglaLogradouro As String, _
                                        strLogradouro As String, _
                                        INTNUMERO As String, _
                                        STRBAIRRO As String, _
                                        intNumeroEmpregados As Integer, _
                                        dblAreaOcupada As Double, _
                                        strdetalheHorario As String, _
                                        strNumeroProcesso As String, _
                                        STRCNPJCPF As String, _
                                        strIE As String, _
                                        strNumeroJucesp As String, _
                                        strObservacao As String, _
                                        XArrayTabela As XArrayDB, _
                                        XArrayAlinhaColunas As XArrayDB, BitDefinitivo As Byte, _
                                        strDtmrazaoinicio As String, _
                                        strDtmenderecoinicio As String, _
                                        strOcorrenciaProcesso As String, _
                                        XArrayTabelaHorario As XArrayDB, _
                                        XArrayAlinhaColunasHorario As XArrayDB)

                Dim MODELO                  As String
                Dim blpMsg                  As Boolean
                Dim stpDocument             As String
                Dim stpTemplate             As String
                Dim objFileSystem           As Scripting.FileSystemObject
                Dim stpTemplatePath         As String
                Dim stpDocumentPath         As String
                Dim strSQL                  As String

    If BitDefinitivo = 0 Then
        MODELO = "INSCRICAO PROVISORIA"
    Else
        MODELO = "ALVARA DE FUNCIONAMENTO"
    End If
                                    
    stpTemplate = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & MODELO & ".dot"
    stpTemplatePath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\"
    stpDocumentPath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordGravados\"
   
    Set objFileSystem = New Scripting.FileSystemObject
    ReDim Documentos(1)
    
    If objFileSystem.FolderExists(stpTemplatePath) Then
        If objFileSystem.FileExists(stpTemplate) Then
            If objFileSystem.FolderExists(stpDocumentPath) Then
                blpMsg = True
                stpDocument = stpDocumentPath & MODELO & "_" & gstrFormataInscricao(strInscricaoCadastral, TYP_ECONOMICA) & "_" & Year(gstrDataDoSistema) & "_" & Format$(strContAlvaraFuncionamento, "0000000000") & ".doc"
                
                If objFileSystem.FileExists(stpDocument) Then
                    blpMsg = (MsgBox("O Alvar� de Funcionamento:" & MODELO & "_" & gstrFormataInscricao(strInscricaoCadastral, TYP_ECONOMICA) & "_" & Year(gstrDataDoSistema) & "_" & Format$(strContAlvaraFuncionamento, "0000000000") & " j� existe. Deseja atualiz�-la a partir do modelo original ?", vbYesNo + vbInformation, "Mensagem ao usu�rio") = vbYes)
                    If blpMsg Then
                        objFileSystem.DeleteFile stpDocument, True
                    End If
                End If
                                        
                Set Documentos(1) = New cWordWrapper
                
                Documentos(1).GetContainer
                Documentos(1).DocumentTemplatePath = stpTemplate
                Documentos(1).DocumentPath = stpDocument
                Documentos(1).DocumentFormat = WORDOPENFORMATDOCUMENT
                Documentos(1).DocumentOpen
                
                
                
                Documentos(1).DocumentReplaceField "|Data Validade|", strDataVencimento
                Documentos(1).DocumentReplaceField "|Data Abertura Razao|", strDtmrazaoinicio
                Documentos(1).DocumentReplaceField "|Inscri��o Cadastral|", gstrFormataInscricao(strInscricaoCadastral, TYP_ECONOMICA)
                Documentos(1).DocumentReplaceField "|Nome do Contribuinte|", strRazaoSocial
                Documentos(1).DocumentReplaceField "|Tipo Logradouro|", strSiglaLogradouro
                Documentos(1).DocumentReplaceField "|Logradouro|", strLogradouro
                Documentos(1).DocumentReplaceField "|numero|", INTNUMERO
                Documentos(1).DocumentReplaceField "|Bairro|", STRBAIRRO
                Documentos(1).DocumentReplaceField "|Data Logradouro|", strDtmenderecoinicio
                Documentos(1).DocumentReplaceField "|N. Empregados|", intNumeroEmpregados
                Documentos(1).DocumentReplaceField "|Area Ocupada|", dblAreaOcupada
                Documentos(1).DocumentReplaceField "|Descri��o hor�rio|", strdetalheHorario
                Documentos(1).DocumentReplaceField "|N. Processo Adm|", strNumeroProcesso
                Documentos(1).DocumentReplaceField "|Ocorrencia do Processo|", strOcorrenciaProcesso
                Documentos(1).DocumentReplaceField "|CNPJ/RG|", STRCNPJCPF
                Documentos(1).DocumentReplaceField "|IE|", strIE
                Documentos(1).DocumentReplaceField "|N. Jucesp|", strNumeroJucesp
                Documentos(1).DocumentReplaceField "|Observa��o|", strObservacao
                Documentos(1).DocumentReplaceField "|Data Atual|", gstrDataPorExtenso
                Documentos(1).DocumentInsert "|Tabela|", , XArrayTabela, XArrayAlinhaColunas
                Documentos(1).DocumentInsert "|Horario|", , XArrayTabelaHorario, XArrayAlinhaColunasHorario, 3
                Documentos(1).DocumentSave
                    
            Else
                MsgBox "A pasta de documentos : " & stpDocumentPath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
            End If

        Else
            MsgBox "O modelo de documento do Microsoft Word : " & stpTemplate & " n�o foi localizado. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
        End If
    Else
        MsgBox "A pasta de modelos de documentos : " & stpTemplatePath & " n�o foi localizada. A opera��o n�o pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usu�rio"
    End If
    
    Set objFileSystem = Nothing

End Sub

Public Sub ImprimirTermo(lngPkid As Long)
Dim strSQL                    As String
Dim adoResultado              As ADODB.Recordset
Dim adoResultadoParcela       As ADODB.Recordset
Dim strInscricaoAcordo        As String
Dim strInscricaoRepresentante As String
Dim strProprietarioAcordo     As String
Dim StrCnpjCpfAcordo          As String
Dim strIdentidadeAcordo       As String
Dim strLogradouroAcordo       As String
Dim strDataAcordo             As String
Dim strMoedaAcordo            As String
Dim strSqlSub                 As String
Dim strSqlSub2                As String
Dim Strindexador              As String
Dim dblvlIndexador            As Double
Dim dblValorParcela           As Double
Dim dblValorTotal             As Double
Dim intQtdeParcelasAcordo     As Integer
    
    Screen.MousePointer = vbHourglass

    strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " strIdentificacao " & _
                "FROM " & gstrAcordoDebitos & " AD WHERE intAcordo = AC.Pkid " & _
                "ORDER BY strIdentificacao, strComposicaoDaReceita, intExercicio"
    strSqlSub = gstrTOPnOracle(strSqlSub, 1, "intAcordo", "AC.Pkid", "strIdentificacao")
    
    strSqlSub2 = "SELECT " & gstrTOPnSQLServer(1) & " intUtilizacao " & _
                "FROM " & gstrAcordoDebitos & " AD WHERE intAcordo = AC.Pkid " & _
                "ORDER BY strIdentificacao, strComposicaoDaReceita, intExercicio"
    strSqlSub2 = gstrTOPnOracle(strSqlSub2, 1, "intAcordo", "AC.Pkid", "intUtilizacao")
    
    'Vamos pegar o Contribuinte requerente do " Acordo "
    strSQL = "Select "
    strSQL = strSQL & "LA.strInscricao AS strInscricaoAcordo, "
    strSQL = strSQL & "LA.STRNOMEPROPRIETARIO, "
    strSQL = strSQL & "LA.Strcnpjcpf, "
    strSQL = strSQL & "LA.STRIDENTIDADE, "
    strSQL = strSQL & "Ltrim(Rtrim(LA.STRLOGRADOUROC)) " & strCONCAT & " ',' " & strCONCAT & " Ltrim(Rtrim(LA.STRNUMEROC)) " & strCONCAT & " ' ' " & strCONCAT & " "
    strSQL = strSQL & "Ltrim(Rtrim(LA.StrcomplementoC)) " & strCONCAT & " ' ' " & strCONCAT & " Ltrim(Rtrim(LA.STRBAIRROC)) " & strCONCAT & " ' ' " & strCONCAT & " "
    strSQL = strSQL & "Ltrim(Rtrim(LA.STRMUNICIPIOC)) " & strCONCAT & " ' ' " & strCONCAT & " Ltrim(Rtrim(LA.StrufC)) " & strCONCAT & " ' CEP: ' " & strCONCAT & " Ltrim(Rtrim(LA.INTCEPC)) AS strLogradouro, "
    strSQL = strSQL & "AC.dtmData DataAcordo, "
    strSQL = strSQL & "(" & strSqlSub & ") strIdentificacaoRepresentante,"
    strSQL = strSQL & "(" & strSqlSub2 & ") intUtilizacaoRepresentante,"
    strSQL = strSQL & "ME.Strabreviatura strMoeda, "
    strSQL = strSQL & "US.strNome strUsuario, "
    strSQL = strSQL & "LA.Strindexador , "
    strSQL = strSQL & "LA.Dblvlindexador "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrMoedas & " ME, "
    strSQL = strSQL & gstrAcordo & " AC, "
    strSQL = strSQL & gstrUsuarios & " US "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "AC.intLancamentoAlfa = LA.Pkid AND "
    strSQL = strSQL & "ME.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " AC.Intmoedas AND "
    strSQL = strSQL & "US.Pkid " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.lngCodUsr AND "
    strSQL = strSQL & "LA.Pkid = " & lngPkid
    
    Set gobjBanco = New clsBanco
    Set adoResultado = New ADODB.Recordset
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                strInscricaoAcordo = gstrFormataInscricao(Right(adoResultado("strInscricaoAcordo").Value, gintRetornaTamanhoMascara(TYP_ACORDO)), TYP_ACORDO)
                strProprietarioAcordo = gstrENulo(!strnomeproprietario)
                StrCnpjCpfAcordo = gstrCGCCPFFormatado(gstrENulo(!STRCNPJCPF))
                strIdentidadeAcordo = gstrENulo(!STRIDENTIDADE)
                strLogradouroAcordo = gstrENulo(!strLogradouro)
                strDataAcordo = gstrENulo(!DataAcordo)
                strMoedaAcordo = gstrENulo(!strMoeda)
                strInscricaoRepresentante = gstrFormataInscricao(Right(adoResultado("strIdentificacaoRepresentante").Value, gintRetornaTamanhoMascara(adoResultado("intUtilizacaoRepresentante").Value)), adoResultado("intUtilizacaoRepresentante").Value)
                Strindexador = gstrENulo(!Strindexador)
                dblvlIndexador = IIf(IsNull(!dblvlIndexador), 0, gstrENulo(!dblvlIndexador))
                If blnPreencheParcela(lngPkid) Then
                    Set gobjBanco = New clsBanco
                    Set adoResultadoParcela = New ADODB.Recordset
                    strSQL = "SELECT (SELECT sum(dblvalor) FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & lngPkid & ") ValorTotal, dblvalor FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & lngPkid & " Order By intparcela asc "
                    If gobjBanco.CriaADO(strSQL, 5, adoResultadoParcela) Then
                        If Not adoResultadoParcela.EOF Then
                            dblValorParcela = gstrENulo(adoResultadoParcela!DBLVALOR)
                            dblValorTotal = gstrENulo(adoResultadoParcela!ValorTotal)
                            intQtdeParcelasAcordo = adoResultadoParcela.RecordCount
                        End If
                    End If
                    Alinhamento
                    OpenWordDocumentTermoDeAcordo strInscricaoAcordo, Right$(Trim(strInscricaoAcordo), 4), _
                    strDataAcordo, strProprietarioAcordo, StrCnpjCpfAcordo, strIdentidadeAcordo, strLogradouroAcordo, _
                    strInscricaoRepresentante, "", "", strMoedaAcordo, gstrENulo(!strUsuario), Strindexador, dblvlIndexador, dblValorParcela, intQtdeParcelasAcordo, dblValorTotal, XParcelas, XArrayAlinhaColunas
                Else
                    ExibeMensagem "N�o foi poss�vel imprimir o Termo de Acordo, pois n�o foi retornada nenhuma parcela."
                End If
                
            End With
        End If
    End If
    Set gobjBanco = Nothing
    Set adoResultado = Nothing
    
    'Vamos pegar o restante dos dados
'    strSQL = "SELECT LA.strInscricao AS strInscricaoRepresentante, "
'    strSQL = strSQL & "LA.strNomeProprietario NomeProprietario, "
'    strSQL = strSQL & "Ltrim(Rtrim(LA.STRLOGRADOURO)) " & strCONCAT & " ',' " & strCONCAT & " Ltrim(Rtrim(LA.STRNUMERO)) " & strCONCAT & " ' ' " & strCONCAT & " "
'    strSQL = strSQL & "Ltrim(Rtrim(LA.Strcomplemento)) " & strCONCAT & " ' ' " & strCONCAT & " Ltrim(Rtrim(LA.STRBAIRRO)) " & strCONCAT & " ' ' " & strCONCAT & " "
'    strSQL = strSQL & "Ltrim(Rtrim(LA.STRMUNICIPIO)) " & strCONCAT & " ' ' " & strCONCAT & " Ltrim(Rtrim(LA.Struf)) " & strCONCAT & " ' CEP: ' " & strCONCAT & " Ltrim(Rtrim(LA.INTCEP)) AS strLogradouro "
'    strSQL = strSQL & " FROM "
'    strSQL = strSQL & gstrLancamentoAlfa & " LA "
'    strSQL = strSQL & "WHERE "
'    strSQL = strSQL & "LA.Pkid = " & lngPkid
'    strSQL = strSQL & " ORDER BY LA.strInscricao"
'
'    Set gobjBanco = New clsBanco
'    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
'        If Not adoResultado.EOF Then
'            With adoResultado
'                If blnPreencheParcela(lngPkid) Then
'                    Alinhamento
'                    OpenWordDocumentTermoDeAcordo Mid(strInscricaoAcordo, 1, Len(strInscricaoAcordo) - 4), Right$(Trim(strInscricaoAcordo), 4), _
'                    strDataAcordo, strProprietarioAcordo, StrCnpjCpfAcordo, strIdentidadeAcordo, strLogradouroAcordo, _
'                    gstrENulo(!strInscricaoRepresentante), gstrENulo(!NomeProprietario), gstrENulo(!strLogradouro), strMoedaAcordo, XParcelas, XArrayAlinhaColunas
'                Else
'                    ExibeMensagem "N�o foi poss�vel imprimir o Termo de Acordo, pois n�o foi retornada nenhuma parcela."
'                End If
'            End With
'        End If
'    End If
'    Set gobjBanco = Nothing

    Screen.MousePointer = vbDefault
    
End Sub

Private Function blnPreencheParcela(lngPkid As Long) As Boolean
Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
Dim intPosition     As Integer
Dim varAux          As Variant
    
Dim dblOriginal     As Double
Dim dblPrincipal    As Double
Dim dblMulta        As Double
Dim dblJuros        As Double
Dim dblCorrecao     As Double
Dim dblTotal        As Double
    
    blnPreencheParcela = False
    
    Set XParcelas = New XArrayDB
    XParcelas.Clear
    XParcelas.ReDim 0, 0, 0, 14
        
    'Vamos adicionar o cabe��rio da coluna
    varAux = "C. Rec."
    XParcelas(0, 0) = varAux
    varAux = "Exer."
    XParcelas(0, 1) = varAux
    varAux = "Identifica��o"
    XParcelas(0, 2) = varAux
    varAux = "Aviso"
    XParcelas(0, 3) = varAux
    varAux = "Parcela"
    XParcelas(0, 4) = varAux
    varAux = "Vencimento"
    XParcelas(0, 5) = varAux
    varAux = ""
    XParcelas(0, 6) = varAux
    varAux = "Vl. Original"
    XParcelas(0, 7) = varAux
    varAux = "Vl. Principal"
    XParcelas(0, 8) = varAux
    varAux = "Vl. Multa"
    XParcelas(0, 9) = varAux
    varAux = "Vl. Juros"
    XParcelas(0, 10) = varAux
    varAux = "Vl. Corre��o"
    XParcelas(0, 11) = varAux
    varAux = "Vl. Total"
    XParcelas(0, 12) = varAux
    varAux = "Certid�o"
    XParcelas(0, 13) = varAux
    varAux = "Executivo"
    XParcelas(0, 14) = varAux
    
    'Vamos trazer as parcelas originais
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "ad.STRCOMPOSICAODARECEITA as Trib, "
    strSQL = strSQL & "ad.intUtilizacao , "
    strSQL = strSQL & "ad.Intexercicio as Exercicio, "
    strSQL = strSQL & "ad.stridentificacao as Ident, "
    strSQL = strSQL & "ad.STRNUMEROAVISO as Aviso, "
    strSQL = strSQL & "ad.INTPARCELA as Parcela, "
    strSQL = strSQL & "ad.Dtmdtvencimento as Vencimento, "
    strSQL = strSQL & "ad.strprincipaloriginalmoeda as strMoedaOrig, "
    strSQL = strSQL & "ad.dblprincipaloriginal as dblValororig , "
    strSQL = strSQL & "ad.Dblprincipal as DblValor, "
    strSQL = strSQL & "ad.DBLMULTA as DblMulta, "
    strSQL = strSQL & "ad.DBLJUROS as DblJuros, "
    strSQL = strSQL & "ad.DBLCORRECAOMONETARIA as dblCorrecao, "
    strSQL = strSQL & "ad.intcertidao as Certidao, "
    strSQL = strSQL & gstrCONVERT(CDT_VARCHAR, "ad.intexecutivonumero") & strCONCAT & " '/' " & strCONCAT & gstrCONVERT(CDT_VARCHAR, "ad.intexecutivoserie") & " as Executivo, "
    strSQL = strSQL & "(" & gstrISNULL("ad.Dblprincipal", "0") & " + " & gstrISNULL("ad.DBLMULTA", "0") & " + "
    strSQL = strSQL & gstrISNULL("ad.DBLJUROS", "0") & " + " & gstrISNULL("ad.DBLCORRECAOMONETARIA", "0") & ") as dblTotal "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrAcordo & " A, "
    strSQL = strSQL & gstrAcordoDebitos & " AD "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "LA.Pkid = A.Intlancamentoalfa AND "
    strSQL = strSQL & "A.Pkid = AD.Intacordo AND "
    strSQL = strSQL & "LA.Pkid = " & lngPkid
    
    intPosition = 1
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                Do While Not .EOF
                    XParcelas.ReDim 0, intPosition, 0, 14
                    
                    dblOriginal = dblOriginal + gstrConvVrDoSql(gstrENulo(!dblValororig), , , True)
                    dblPrincipal = dblPrincipal + gstrConvVrDoSql(gstrENulo(!DBLVALOR), , , True)
                    dblMulta = dblMulta + gstrConvVrDoSql(gstrENulo(!dblMulta), , , True)
                    dblJuros = dblJuros + gstrConvVrDoSql(gstrENulo(!dblJuros), , , True)
                    dblCorrecao = dblCorrecao + gstrConvVrDoSql(gstrENulo(!dblCorrecao), , , True)
                    dblTotal = dblTotal + gstrConvVrDoSql(gstrENulo(!dblTotal), , , True)
                    
                    varAux = gstrENulo(!Trib)
                    XParcelas(intPosition, 0) = varAux
                    '
                    varAux = gstrENulo(!Exercicio)
                    XParcelas(intPosition, 1) = varAux
                    '
                    varAux = Space$(0) & gstrFormataInscricao(Right(adoResultado("ident").Value, gintRetornaTamanhoMascara(adoResultado("intUtilizacao").Value)), adoResultado("intUtilizacao").Value)
                    XParcelas(intPosition, 2) = varAux
                    '
                    varAux = gstrENulo(!Aviso)
                    XParcelas(intPosition, 3) = varAux
                    '
                    varAux = gstrENulo(!PARCELA)
                    XParcelas(intPosition, 4) = varAux
                    '
                    varAux = gstrDataFormatada(gstrENulo(!Vencimento))
                    XParcelas(intPosition, 5) = varAux
                    '
                    varAux = gstrENulo(!strMoedaOrig)
                    XParcelas(intPosition, 6) = varAux
                    '
                    varAux = gstrConvVrDoSql(gstrENulo(!dblValororig), , , True)
                    XParcelas(intPosition, 7) = varAux
                    '
                    varAux = gstrConvVrDoSql(gstrENulo(!DBLVALOR), , , True)
                    XParcelas(intPosition, 8) = varAux
                    '
                    varAux = gstrConvVrDoSql(gstrENulo(!dblMulta), , , True)
                    XParcelas(intPosition, 9) = varAux
                    '
                    varAux = gstrConvVrDoSql(gstrENulo(!dblJuros), , , True)
                    XParcelas(intPosition, 10) = varAux
                    '
                    varAux = gstrConvVrDoSql(gstrENulo(!dblCorrecao), , , True)
                    XParcelas(intPosition, 11) = varAux
                    '
                    varAux = gstrConvVrDoSql(gstrENulo(!dblTotal), , , True)
                    XParcelas(intPosition, 12) = varAux
                    '
                    varAux = gstrENulo(!Certidao)
                    XParcelas(intPosition, 13) = varAux
                    '
                    varAux = gstrENulo(!EXECUTIVO)
                    XParcelas(intPosition, 14) = varAux

                    intPosition = intPosition + 1
                    .MoveNext
                Loop
                    'Vamos adicionar uma linha em branco
                    XParcelas.ReDim 0, intPosition, 0, 14
                    varAux = " "
                    XParcelas(intPosition, 0) = varAux
                    XParcelas(intPosition, 1) = varAux
                    XParcelas(intPosition, 2) = varAux
                    XParcelas(intPosition, 3) = varAux
                    XParcelas(intPosition, 4) = varAux
                    XParcelas(intPosition, 5) = varAux
                    XParcelas(intPosition, 6) = varAux
                    varAux = "--------"
                    XParcelas(intPosition, 7) = varAux
                    XParcelas(intPosition, 8) = varAux
                    XParcelas(intPosition, 9) = varAux
                    XParcelas(intPosition, 10) = varAux
                    XParcelas(intPosition, 11) = varAux
                    XParcelas(intPosition, 12) = varAux
                    varAux = " "
                    XParcelas(intPosition, 13) = varAux
                    XParcelas(intPosition, 14) = varAux

                    intPosition = intPosition + 1
                    
                    'Vamos inserir os totais
                    XParcelas.ReDim 0, intPosition, 0, 14
                    varAux = " "
                    XParcelas(intPosition, 0) = varAux
                    XParcelas(intPosition, 1) = varAux
                    XParcelas(intPosition, 2) = varAux
                    XParcelas(intPosition, 3) = varAux
                    XParcelas(intPosition, 4) = varAux
                    XParcelas(intPosition, 5) = varAux
                    XParcelas(intPosition, 6) = varAux
                    varAux = gstrConvVrDoSql(dblOriginal)
                    XParcelas(intPosition, 7) = varAux
                    varAux = gstrConvVrDoSql(dblPrincipal)
                    XParcelas(intPosition, 8) = varAux
                    varAux = gstrConvVrDoSql(dblMulta)
                    XParcelas(intPosition, 9) = varAux
                    varAux = gstrConvVrDoSql(dblJuros)
                    XParcelas(intPosition, 10) = varAux
                    varAux = gstrConvVrDoSql(dblCorrecao)
                    XParcelas(intPosition, 11) = varAux
                    varAux = gstrConvVrDoSql(dblTotal)
                    XParcelas(intPosition, 12) = varAux
                    varAux = " "
                    XParcelas(intPosition, 13) = varAux
                    XParcelas(intPosition, 14) = varAux
                    
                    blnPreencheParcela = True
            End With
        End If
    End If
    
End Function

Public Sub Alinhamento()

    Set XArrayAlinhaColunas = New XArrayDB
    
    With XArrayAlinhaColunas 'Alinhamento
        .Clear
        .ReDim 0, 0, 0, 14
        .Value(0, 0) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 1) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 2) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 3) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 4) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 5) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 6) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 7) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 8) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 9) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 10) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 11) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 12) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 13) = WORDALIGNPARAGRAPHCENTER
        .Value(0, 14) = WORDALIGNPARAGRAPHCENTER
        
    End With
End Sub

Public Function gstrQueryCarneAcordo(strAcordoInicial As String, _
                                    strAcordoFinal As String, _
                                    strParcelas As String, _
                                    Optional blnTodosAcordos As Boolean = False, _
                                    Optional blnExercicio As Boolean = True)
Dim strSQL As String
Dim strOpcao As String
  
  strOpcao = ""
  If blnTodosAcordos = False Then
     If strAcordoInicial <> "" And strAcordoFinal <> "" Then
        strOpcao = "LA.strInscricao BETWEEN '" & String(gintLenInscricao - Len(Trim(strAcordoInicial)), "0") & Trim(strAcordoInicial) & "' AND '"
        strOpcao = strOpcao & String(gintLenInscricao - Len(Trim(strAcordoFinal)), "0") & Trim(strAcordoFinal) & "' AND "
     Else
        If strAcordoInicial <> "" Then
           If blnExercicio Then
              strOpcao = "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strAcordoInicial)), "0") & Trim(strAcordoInicial) & "' AND "
           Else
              strOpcao = "LA.strInscricao LIKE '" & String(gintLenInscricao - Len(Trim(strAcordoInicial)) - 4, "0") & Trim(strAcordoInicial) & "%' AND "
           End If
        Else
           If blnExercicio Then
              strOpcao = "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strAcordoFinal)), "0") & Trim(strAcordoFinal) & "' AND "
           Else
              strOpcao = "LA.strInscricao LIKE '" & String(gintLenInscricao - Len(Trim(strAcordoFinal)) - 4, "0") & Trim(strAcordoFinal) & "%' AND "
           End If
        End If
     End If
  End If
  
  strParcelas = "LV.intParcela IN (" & strParcelas & ") "
  
  strSQL = ""
  strSQL = strSQL & "SELECT "
  strSQL = strSQL & "LA.strNomeProprietario strContribuinte, "
  strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ACORDO)) & " strInscricao, "
  strSQL = strSQL & "LA.intExercicio intExercicio, "
  
  strSQL = strSQL & "CASE WHEN AC.strCodigoProcesso IS NOT NULL THEN ( "
  strSQL = strSQL & "AC.strCodigoProcesso " & strCONCAT & "'/' " & strCONCAT & gstrCONVERT(CDT_VARCHAR, "AC.intExercicioProcesso ")
  strSQL = strSQL & strCONCAT & "'-' " & strCONCAT & gstrCONVERT(CDT_VARCHAR, "AC.bitDigitoProcesso ") & ") ELSE NULL END "
  strSQL = strSQL & "strProcesso , "
  strSQL = strSQL & strSUBSTRING & "(LA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ") strAcordo, "
  strSQL = strSQL & "LA.strComposicaoDaReceita strComposicao, LA.intComposicaoDaReceita intComposicao, "
  strSQL = strSQL & "CR.strSigla strSigla, "
  strSQL = strSQL & "LA.strEmissao strEmissao, "
  strSQL = strSQL & "LA.strNomeProprietario strProprietario, "
  strSQL = strSQL & "LA.strLogradouroC strLogradouroC, "
  strSQL = strSQL & "LA.strNumeroC strNumeroC, LA.strComplementoC strComplementoC, "
  strSQL = strSQL & "LA.strBairroC strBairroC, LA.strMunicipioC strMunicipioC, "
  strSQL = strSQL & "LA.strUFC strUFC, LA.intCEPC intCEPC, "
  strSQL = strSQL & "LA.strLogradouro strLogradouro, "
  strSQL = strSQL & "LA.strNumero strNumero, LA.strComplemento strComplemento, "
  strSQL = strSQL & "LA.strBairro strBairro, LA.strMunicipio strMunicipio, "
  strSQL = strSQL & "LA.strUF strUF, LA.intCEP intCEP, "
  
  strSQL = strSQL & "LA.dblvlIndexador dblvlIndexador, "
  strSQL = strSQL & "LA.strIndexador, "

  strSQL = strSQL & "AD.dblPrincipal dblPrincipal, "
  strSQL = strSQL & "AD.dblCorrecaoMonetaria dblCorrecaoMonetaria, "
  strSQL = strSQL & "AD.dblMulta dblMulta, "
  strSQL = strSQL & "AD.dblJuros dblJuros, "
  strSQL = strSQL & "AD.dblTotal dblTotal, "
  
  strSQL = strSQL & "CASE WHEN LA.dblvlIndexador <> 0 AND LA.dblvlIndexador IS NOT NULL AND LA.strIndexador IS NOT NULL THEN "
  strSQL = strSQL & "(AD.dblTotal / LA.dblvlIndexador) ELSE AD.dblTotal END dblTotalFMP, "
  
  strSQL = strSQL & "LV.dblValor dblPrimeiraParcela, "
  strSQL = strSQL & "LV.dtmdtVencimento dtmdtPrimeiroVencimento, "
  
  strSQL = strSQL & "CASE WHEN LA.dblvlIndexador <> 0 AND LA.dblvlIndexador IS NOT NULL AND LA.strIndexador IS NOT NULL THEN "
  strSQL = strSQL & "(LV.dblValor / LA.dblvlIndexador) ELSE 0 END dblQuantidadeFMP, "
  
  strSQL = strSQL & "LV.intNumeroParcelas intNumeroParcelas "
    
  'FROM
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrAcordo & " AC, "
  strSQL = strSQL & gstrComposicaoDaReceita & " CR, "
  
  'ACORDO D�BITOS
  strSQL = strSQL & "( "
  strSQL = strSQL & "SELECT "
  strSQL = strSQL & "AD.intAcordo intAcordo, "
  strSQL = strSQL & "COUNT(AD.intAcordo) intNumeroParcelas, "
  strSQL = strSQL & "SUM(AD.dblPrincipal) dblPrincipal, "
  strSQL = strSQL & "SUM(AD.dblCorrecaoMonetaria) dblCorrecaoMonetaria, "
  strSQL = strSQL & "SUM(AD.dblMulta) dblMulta, "
  strSQL = strSQL & "SUM(AD.dblJuros) dblJuros, "
  strSQL = strSQL & "SUM(AD.dblPrincipal) + SUM(AD.dblCorrecaoMonetaria) + SUM(AD.dblMulta) + SUM(AD.dblJuros) dblTotal "
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrAcordoDebitos & " AD, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrAcordo & " AC "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & strOpcao
  strSQL = strSQL & "AC.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & "AD.intAcordo = AC.Pkid "
  strSQL = strSQL & "GROUP BY "
  strSQL = strSQL & "AD.intAcordo "
  strSQL = strSQL & ") AD, "
  
  'VALOR DA PRIMEIRA PARCELA
  strSQL = strSQL & "( "
  strSQL = strSQL & "SELECT "
  strSQL = strSQL & "LV.intLancamentoAlfa, LV.dblValor dblValor , LV.dtmdtVencimento dtmdtVencimento, "
  strSQL = strSQL & "LVS.intNumeroParcelas "
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrLancamentoValor & " LV, "
  strSQL = strSQL & "(SELECT MIN(LV.intParcela) intParcela, LV.intLancamentoAlfa, "
  strSQL = strSQL & "COUNT(LV.intLancamentoAlfa) intNumeroParcelas "
  strSQL = strSQL & "FROM " & gstrLancamentoValor & " LV, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrAcordo & " AC "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & IIf(strOpcao <> "", strOpcao & "LV.intLancamentoAlfa = LA.pkID AND ", "")
  strSQL = strSQL & "AC.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & "LV.bitParcelaValida  = 1 AND "
  strSQL = strSQL & "LV.intParcela > 0 "
  strSQL = strSQL & "GROUP BY LV.intLancamentoAlfa) LVS "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & "LV.intLancamentoAlfa = LVS.intLancamentoAlfa AND "
  strSQL = strSQL & "LV.intParcela = LVS.intParcela "
  strSQL = strSQL & ") LV, "
  
  'VERIFICA SE CONT�M AS PARCELAS SELECIONADAS
  strSQL = strSQL & "( "
  strSQL = strSQL & "SELECT "
  strSQL = strSQL & "LV.intLancamentoAlfa "
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrLancamentoValor & " LV, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrAcordo & " AC "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & IIf(strOpcao <> "", strOpcao & " LV.intLancamentoAlfa = LA.pkID AND ", "")
  strSQL = strSQL & "AC.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & strParcelas
  strSQL = strSQL & "GROUP BY "
  strSQL = strSQL & "LV.intLancamentoAlfa "
  strSQL = strSQL & ") LVG "
  
  'WHERE
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & "LA.pkID = LV.intLancamentoAlfa AND "
  strSQL = strSQL & "LVG.intLancamentoAlfa = LV.intLancamentoAlfa AND "
  strSQL = strSQL & "AC.intLancamentoAlfa = LV.intLancamentoAlfa AND "
  strSQL = strSQL & "AD.intAcordo " & strOUTJOracle & "=" & strOUTJSQLServer & " AC.pkID AND "
  strSQL = strSQL & "CR.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.intComposicaoDaReceita "
  
  'ORDER BY
  strSQL = strSQL & "ORDER BY strInscricao, intExercicio DESC "
  
  gstrQueryCarneAcordo = strSQL
End Function

Public Function gstrQueryCarneAcordoAtualizadas(strAcordoInicial As String, _
                                    strAcordoFinal As String, _
                                    intExercicio As Integer, _
                                    Optional blnTodosAcordos As Boolean = False, _
                                    Optional blnExercicio As Boolean = True)
Dim strSQL As String
Dim strOpcao As String
  
  strOpcao = ""
  If blnTodosAcordos = False Then
     If strAcordoInicial <> "" And strAcordoFinal <> "" Then
        strOpcao = strSUBSTRING & "(LA.strInscricao,17,4) " & strCONCAT & " " & strSUBSTRING & "(LA.strInscricao,1,16) BETWEEN '" & Right(String(gintLenInscricao - Len(Trim(strAcordoInicial)), "0") & Trim(strAcordoInicial), 4) & Left(String(gintLenInscricao - Len(Trim(strAcordoInicial)), "0") & Trim(strAcordoInicial), 16) & "' AND '"
        strOpcao = strOpcao & Right(String(gintLenInscricao - Len(Trim(strAcordoFinal)), "0") & Trim(strAcordoFinal), 4) & Left(String(gintLenInscricao - Len(Trim(strAcordoFinal)), "0") & Trim(strAcordoFinal), 16) & "' AND "
'        strOpcao = " LA.Pkid BETWEEN " & strAcordoInicial & " AND " & strAcordoFinal & " AND "
     Else
        If strAcordoInicial <> "" Then
           If blnExercicio Then
              strOpcao = "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strAcordoInicial)), "0") & Trim(strAcordoInicial) & "' AND "
           Else
              strOpcao = "LA.strInscricao LIKE '" & String(gintLenInscricao - Len(Trim(strAcordoInicial)) - 4, "0") & Trim(strAcordoInicial) & "%' AND "
           End If
        Else
           If blnExercicio Then
              strOpcao = "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strAcordoFinal)), "0") & Trim(strAcordoFinal) & "' AND "
           Else
              strOpcao = "LA.strInscricao LIKE '" & String(gintLenInscricao - Len(Trim(strAcordoFinal)) - 4, "0") & Trim(strAcordoFinal) & "%' AND "
           End If
        End If
     End If
  End If
  
  strSQL = ""
  strSQL = strSQL & "SELECT "
  strSQL = strSQL & gstrISNULL("LA.strPromissario", "LA.strNomeProprietario", "La.strPromissario") & " strContribuinte, "
  strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_ACORDO)) & " strInscricao, "
  strSQL = strSQL & "LA.intExercicio intExercicio, "
  
  strSQL = strSQL & "CASE WHEN AC.strCodigoProcesso IS NOT NULL THEN ( "
  strSQL = strSQL & "AC.strCodigoProcesso " & strCONCAT & "'/' " & strCONCAT & gstrCONVERT(CDT_VARCHAR, "AC.intExercicioProcesso ")
  strSQL = strSQL & strCONCAT & "'-' " & strCONCAT & gstrCONVERT(CDT_VARCHAR, "AC.bitDigitoProcesso ") & ") ELSE NULL END "
  strSQL = strSQL & "strProcesso , "
  strSQL = strSQL & strSUBSTRING & "(LA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & ", " & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ") strAcordo, "
  strSQL = strSQL & "LA.strComposicaoDaReceita strComposicao, LA.intComposicaoDaReceita intComposicao, "
  strSQL = strSQL & "CR.strSigla strSigla, "
  strSQL = strSQL & "LA.strEmissao strEmissao, "
  strSQL = strSQL & "LA.strNomeProprietario strProprietario, "
  strSQL = strSQL & "LA.strLogradouroC strLogradouroC, "
  strSQL = strSQL & "LA.strNumeroC strNumeroC, LA.strComplementoC strComplementoC, "
  strSQL = strSQL & "LA.strBairroC strBairroC, LA.strMunicipioC strMunicipioC, "
  strSQL = strSQL & "LA.strUFC strUFC, LA.intCEPC intCEPC, "
  strSQL = strSQL & "LA.strLogradouro strLogradouro, "
  strSQL = strSQL & "LA.strNumero strNumero, LA.strComplemento strComplemento, "
  strSQL = strSQL & "LA.strBairro strBairro, LA.strMunicipio strMunicipio, "
  strSQL = strSQL & "LA.strUF strUF, LA.intCEP intCEP, "
  
  strSQL = strSQL & "LA.dblvlIndexador dblvlIndexador, "
  strSQL = strSQL & "LA.strIndexador, "

  strSQL = strSQL & "AD.dblPrincipal dblPrincipal, "
  strSQL = strSQL & "AD.dblCorrecaoMonetaria dblCorrecaoMonetaria, "
  strSQL = strSQL & "AD.dblMulta dblMulta, "
  strSQL = strSQL & "AD.dblJuros dblJuros, "
  strSQL = strSQL & "AD.dblTotal dblTotal, "
  
  strSQL = strSQL & "CASE WHEN LA.dblvlIndexador <> 0 AND LA.dblvlIndexador IS NOT NULL AND LA.strIndexador IS NOT NULL THEN "
  strSQL = strSQL & "(AD.dblTotal / LA.dblvlIndexador) ELSE AD.dblTotal END dblTotalFMP, "
  
  strSQL = strSQL & "LV.dblValor dblPrimeiraParcela, "
  strSQL = strSQL & "LV.dtmdtVencimento dtmdtPrimeiroVencimento, "
  
  strSQL = strSQL & "CASE WHEN LA.dblvlIndexador <> 0 AND LA.dblvlIndexador IS NOT NULL AND LA.strIndexador IS NOT NULL THEN "
  strSQL = strSQL & "(LV.dblValor / LA.dblvlIndexador) ELSE 0 END dblQuantidadeFMP, "
  
  strSQL = strSQL & "LV.intNumeroParcelas intNumeroParcelas "
    
  'FROM
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrAcordo & " AC, "
  strSQL = strSQL & gstrComposicaoDaReceita & " CR, "
  
  'ACORDO D�BITOS
  strSQL = strSQL & "( "
  strSQL = strSQL & "SELECT "
  strSQL = strSQL & "AD.intAcordo intAcordo, "
  strSQL = strSQL & "COUNT(AD.intAcordo) intNumeroParcelas, "
  strSQL = strSQL & "SUM(AD.dblPrincipal) dblPrincipal, "
  strSQL = strSQL & "SUM(AD.dblCorrecaoMonetaria) dblCorrecaoMonetaria, "
  strSQL = strSQL & "SUM(AD.dblMulta) dblMulta, "
  strSQL = strSQL & "SUM(AD.dblJuros) dblJuros, "
  strSQL = strSQL & "SUM(AD.dblPrincipal) + SUM(AD.dblCorrecaoMonetaria) + SUM(AD.dblMulta) + SUM(AD.dblJuros) dblTotal "
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrAcordoDebitos & " AD, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrAcordo & " AC "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & strOpcao
  strSQL = strSQL & "AC.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & "AD.intAcordo = AC.Pkid "
  strSQL = strSQL & "GROUP BY "
  strSQL = strSQL & "AD.intAcordo "
  strSQL = strSQL & ") AD, "
  
  'VALOR DA PRIMEIRA PARCELA
  strSQL = strSQL & "( "
  strSQL = strSQL & "SELECT "
  strSQL = strSQL & "LV.intLancamentoAlfa, LV.dblValor dblValor , LV.dtmdtVencimento dtmdtVencimento, "
  strSQL = strSQL & "LVS.intNumeroParcelas "
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrLancamentoValor & " LV, "
  strSQL = strSQL & "(SELECT MIN(LV.intParcela) intParcela, LV.intLancamentoAlfa, "
  strSQL = strSQL & "COUNT(LV.intLancamentoAlfa) intNumeroParcelas "
  strSQL = strSQL & "FROM " & gstrLancamentoValor & " LV, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrAcordo & " AC "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & IIf(strOpcao <> "", strOpcao & "LV.intLancamentoAlfa = LA.pkID AND ", "")
  strSQL = strSQL & "AC.intLancamentoAlfa = LA.pkID AND "
  strSQL = strSQL & "LV.bitParcelaValida  = 1 AND "
  strSQL = strSQL & gstrDATEPART(strYEAR, "LV.dtmDtVencimento") & " = " & intExercicio & " AND "
  strSQL = strSQL & "LV.intParcela > 0 "
  strSQL = strSQL & "GROUP BY LV.intLancamentoAlfa) LVS "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & "LV.intLancamentoAlfa = LVS.intLancamentoAlfa AND "
  strSQL = strSQL & "LV.intParcela = LVS.intParcela "
  strSQL = strSQL & ") LV, "
  
  'VERIFICA SE CONT�M AS PARCELAS SELECIONADAS
  strSQL = strSQL & "( "
  strSQL = strSQL & "SELECT "
  strSQL = strSQL & "LV.intLancamentoAlfa "
  strSQL = strSQL & "FROM "
  strSQL = strSQL & gstrLancamentoValor & " LV, "
  strSQL = strSQL & gstrLancamentoAlfa & " LA, "
  strSQL = strSQL & gstrAcordo & " AC "
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & IIf(strOpcao <> "", strOpcao & " LV.intLancamentoAlfa = LA.pkID AND ", "")
  strSQL = strSQL & "AC.intLancamentoAlfa = LA.pkID "
  strSQL = strSQL & "GROUP BY "
  strSQL = strSQL & "LV.intLancamentoAlfa "
  strSQL = strSQL & ") LVG "
  
  'WHERE
  strSQL = strSQL & "WHERE "
  strSQL = strSQL & "LA.pkID = LV.intLancamentoAlfa AND "
  strSQL = strSQL & "LA.dtmDtCancelamento Is null AND "
  strSQL = strSQL & "LVG.intLancamentoAlfa = LV.intLancamentoAlfa AND "
  strSQL = strSQL & "AC.intLancamentoAlfa = LV.intLancamentoAlfa AND "
  strSQL = strSQL & "AD.intAcordo " & strOUTJOracle & "=" & strOUTJSQLServer & " AC.pkID AND "
  strSQL = strSQL & "CR.pkID " & strOUTJOracle & "=" & strOUTJSQLServer & " LA.intComposicaoDaReceita "
  
  'ORDER BY
  strSQL = strSQL & "ORDER BY intExercicio, strInscricao "
  
  gstrQueryCarneAcordoAtualizadas = strSQL

End Function

Public Sub gQuitacaoDeAcordos(lngLancamentoAlfa As Long, dtmDataPagamento As Date, Optional dtmDataMovimento As Date)
Dim adoResultado As ADODB.Recordset
Dim adoPagamento As ADODB.Recordset
Dim strSQL       As String
Dim strSqlSub    As String

    'Vamos consultar os pagamentos deste acordo
    strSQL = "SELECT LP.DTMDTPAGAMENTO FROM " & gstrLancamentoPagamento & " LP, " & gstrLancamentoValor & " LV " & _
             "WHERE LP.INTLANCAMENTOVALOR " & strOUTJOracle & " =" & strOUTJSQLServer & " LV.pkid AND LV.INTLANCAMENTOALFA = " & lngLancamentoAlfa & " AND LP.DTMDTPAGAMENTO Is Null"
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        With adoResultado
        
            'Caso nao falte nenhuma parcela, vamos dar baixa em todas as parcelas que originaram o acordo
            If .EOF Then
                                
                strSqlSub = "SELECT " & gstrTOPnSQLServer(1) & " Pkid FROM " & gstrCodigoDeBaixa & " WHERE bytTipo = 2 ORDER BY Pkid"
                strSqlSub = gstrTOPnOracle(strSqlSub, 1, , , "Pkid")
    
                strSQL = ""
                strSQL = "SELECT LV.* , (" & strSqlSub & ") intCodigo FROM " & gstrLancamentoValor & " LV WHERE LV.INTLANCAMENTOALFAACORDO = " & lngLancamentoAlfa
                
                If gobjBanco.CriaADO(strSQL, 10, adoPagamento) Then
                    With adoPagamento
                        
                        Do While Not .EOF
                            
                            strSQL = ""
                            strSQL = "INSERT INTO tblLancamentoPagamento(intLancamentoValor, "
                            strSQL = strSQL & " dblValorPrincipal, "
                            strSQL = strSQL & " dblValorMulta, "
                            strSQL = strSQL & " dblValorJuros, "
                            strSQL = strSQL & " dblValorCorrecao, "
                            strSQL = strSQL & " dblValorCorreto, "
                            strSQL = strSQL & " dtmDtPagamento, "
                            strSQL = strSQL & " dtmDtMovimento, "
                            strSQL = strSQL & " dtmDtAtualizacao, "
                            strSQL = strSQL & " lngCodUsr, "
                            strSQL = strSQL & " intCodigoBaixa, "
                            strSQL = strSQL & " strObservacao "
                            strSQL = strSQL & ") "
                        
                            strSQL = strSQL & "VaLues( "
                            strSQL = strSQL & gstrENulo(!Pkid) & ", "
                            strSQL = strSQL & "0.00, "
                            strSQL = strSQL & "0.00, "
                            strSQL = strSQL & "0.00, "
                            strSQL = strSQL & "0.00, "
                            strSQL = strSQL & "0.00, "
                            strSQL = strSQL & gstrConvDtParaSql(dtmDataPagamento) & ", "
                            strSQL = strSQL & gstrConvDtParaSql(dtmDataMovimento) & ", "
                            strSQL = strSQL & strGETDATE & ", "
                            strSQL = strSQL & glngCodUsr & ", "
                            strSQL = strSQL & gstrENulo(!intCodigo) & ", "
                            strSQL = strSQL & "'Quitado pelo Acordo.') "

                            Set gobjBanco = New clsBanco
                            If Not gobjBanco.Execute(strSQL) Then
                                ExibeMensagem "Ocorreu um erro na grava��o dos registros em Lan�amento de Pagamento, na quita��o do acordo, a opera��o n�o foi conclu�da."
                                gobjBanco.ExecutaRollbackTrans
                                Exit Sub
                            End If
                            
                            adoPagamento.MoveNext
                            
                        Loop
                        
                    End With
                End If
                
            End If
            
        End With
    End If
    
End Sub

Public Function ExcluiAcordo(strInscricaoAcordo As String)
Dim adoResultado           As New ADODB.Recordset
Dim adoAcordo              As New ADODB.Recordset
Dim strSQL                 As String
Dim strAcordosParaConsulta As String
Dim strInscricoes          As String

On Error GoTo Problema_Na_Rotina

    'Vamos obter os valores das parcelas da inscricao selecionada
    Set gobjBanco = New clsBanco
        
    'Vamos obter os Pkids das inscricoes para fazer consulta de acordos
    strSQL = "SELECT  LA.Pkid PkidLA, AC.Pkid PkidAC " & _
             "FROM " & gstrLancamentoAlfa & " LA, " & gstrAcordo & " AC " & _
             "WHERE LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strInscricaoAcordo)), "0") & UCase(strInscricaoAcordo) & "' AND " & _
             "LA.INTUTILIZACAO = " & TYP_ACORDO & " AND " & _
             "LA.Pkid = AC.intLancamentoAlfa "

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Do While Not adoResultado.EOF
                strAcordosParaConsulta = strAcordosParaConsulta & adoResultado("PkidLA").Value & ","
            '     strInscricoes = strInscricoes & adoResultado("Pkid").Value & ","
                adoResultado.MoveNext
            Loop
            strAcordosParaConsulta = Mid(strAcordosParaConsulta, 1, Len(strAcordosParaConsulta) - 1)
        Else
            ExibeMensagem "N�o foi encontrado nenhum acordo com esta Inscri��o."
            Exit Function
        End If
    End If
    
    'Verifica se o acordo informado faz parte de outro acordo
    strSQL = "SELECT LV.pkid " & _
             "FROM " & gstrLancamentoValor & " LV " & _
             "WHERE LV.intLancamentoAlfa IN (" & strAcordosParaConsulta & ") AND LV.intLancamentoAlfaAcordo IS NOT NULL "
    If gobjBanco.CriaADO(strSQL, 5, adoAcordo) Then
        If Not adoAcordo.EOF Then
            ExibeMensagem "O acordo informado est� dentro de outro acordo."
            Exit Function
        End If
    End If
    
'ConsultarAcordos:

    'Vamos obter os acordos, caso exista, para exibir no grid Pai
    'strSql = "SELECT  LV.intLancamentoAlfaAcordo " & _
             "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA " & _
             "WHERE LV.intLancamentoAlfa = LA.pkid AND " & _
             "LA.Pkid IN (" & strAcordosParaConsulta & ") AND Not LV.intLancamentoAlfaAcordo Is Null " & _
             "GROUP BY LV.intLancamentoAlfaAcordo "
    
    'If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
    '    If Not adoResultado.EOF Then
    '        strAcordosParaConsulta = Space$(0)
    '        Do While Not adoResultado.EOF
    '            strInscricoes = strInscricoes & adoResultado("intlancamentoalfaacordo").Value & ","
    '            strAcordosParaConsulta = strAcordosParaConsulta & adoResultado("intlancamentoalfaacordo").Value & ","
    '            adoResultado.MoveNext
    '        Loop
    '        strAcordosParaConsulta = Mid(strAcordosParaConsulta, 1, Len(strAcordosParaConsulta) - 1)
    '        GoTo ConsultarAcordos
    '    End If
    'End If
            
    'strInscricoes = Mid(strInscricoes, 1, Len(strInscricoes) - 1)
    
    'Vamos obter os acordos e acordos vinculados
    'strSql = " SELECT LA.Pkid PkidLA, AC.Pkid PkidAC " & _
    '         " FROM " & gstrLancamentoAlfa & " LA, " & gstrAcordo & " AC " & _
    '         " WHERE LA.Pkid IN (" & strAcordosParaConsulta & ") AND LA.Pkid = AC.intLancamentoAlfa "
    
    adoResultado.MoveFirst
    gobjBanco.ExecutaBeginTrans
    Do While Not adoResultado.EOF
    
        'Vamos excluir da tabela Lancamento Receita
        gobjBanco.Execute " DELETE FROM " & gstrLancamentoReceita & " WHERE intLancamentoValor IN (SELECT Pkid FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & adoResultado("PkidLA").Value & ")"
        
        'Vamos excluir da tabela Lancamento Guias
        gobjBanco.Execute " DELETE FROM " & gstrLancamentoGuias & " WHERE intLancamentoValor IN (SELECT Pkid FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & adoResultado("PkidLA").Value & ")"
        
        'Vamos excluir da tabela Lancamento Pagamento
        gobjBanco.Execute " DELETE FROM " & gstrLancamentoPagamento & " WHERE intLancamentoValor IN (SELECT Pkid FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & adoResultado("PkidLA").Value & ")"
        
        'Vamos excluir da tabela Lancamento Valor
        gobjBanco.Execute " DELETE FROM " & gstrLancamentoValor & " WHERE intLancamentoAlfa = " & adoResultado("PkidLA").Value
        
        'Vamos excluir da tabela Acordo Debitos
        gobjBanco.Execute " DELETE FROM " & gstrAcordoDebitos & " WHERE intAcordo = " & adoResultado("PkidAC").Value
        
        'Vamos excluir da tabela Acordos
        gobjBanco.Execute " DELETE FROM " & gstrAcordo & " WHERE Pkid = " & adoResultado("PkidAC").Value
        
        'Vamos desvincular o acordo da tabela Lancamento Valor
        gobjBanco.Execute " UPDATE " & gstrLancamentoValor & " SET intLancamentoAlfaAcordo = Null WHERE intLancamentoAlfaAcordo = " & adoResultado("PkidLA").Value
        
        'Vamos excluir da tabela Lancamento Alfa
        gobjBanco.Execute " DELETE FROM " & gstrLancamentoAlfa & " WHERE Pkid = " & adoResultado("PkidLA").Value
        
        adoResultado.MoveNext
    
    Loop
    gobjBanco.ExecutaCommitTrans

    Exit Function

Problema_Na_Rotina:
    ExibeMensagem "N�o foi poss�vel concluir a opera��o."
    gobjBanco.ExecutaRollbackTrans
    
End Function

Public Function strQueryCarneISSConstrucao(lngPkid As Long) As String
    Dim strSQL  As String
    Dim strSql1 As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "LI.Pkid, "
    strSQL = strSQL & "LA.Pkid as IntLancamentoAlfa, "
    strSQL = strSQL & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao, "
    strSQL = strSQL & "LA.strComposicaoDaReceita, "
    strSQL = strSQL & "LA.intExercicio, "
    strSQL = strSQL & gstrRIGHT("LA.strNumeroAviso", gintLenNumAviso) & " strNumeroAviso, "
    strSQL = strSQL & gstrRIGHT("LA.strEmissao", gintLenEmissao) & " strEmissao, "
    strSQL = strSQL & "LA.strNomeProprietario, "
    strSQL = strSQL & "LA.Strpromissario, "
    strSQL = strSQL & "LA.strInscricao, "
    strSQL = strSQL & "LA.strLogradouro, "
    strSQL = strSQL & "LA.strNumero, "
    strSQL = strSQL & "LA.strComplemento, "
    strSQL = strSQL & "LA.strBairro, "
    strSQL = strSQL & "LA.strMunicipio, "
    strSQL = strSQL & "LA.strUf, "
    strSQL = strSQL & "LA.intCep, "
    strSQL = strSQL & "LA.strLogradouroC, "
    strSQL = strSQL & "LA.strNumeroC, "
    strSQL = strSQL & "LA.strComplementoC, "
    strSQL = strSQL & "LA.strBairroC, "
    strSQL = strSQL & "LA.strMunicipioC, "
    strSQL = strSQL & "LA.strUfC, "
    strSQL = strSQL & "LA.intCepC, "
    strSQL = strSQL & "LA.Strindexador, "
    strSQL = strSQL & "LA.Dblvlindexador, "
    strSQL = strSQL & "LI.strCodigoProcesso" & strCONCAT & "'/'" & strCONCAT & "LI.intExercicioProcesso" & strCONCAT & "'-'" & strCONCAT & "LI.bitDigitoProcesso as strProcesso, "
    strSQL = strSQL & "LI.strObservacoes, "
    strSQL = strSQL & "LI.dtmLancamento, "
    strSQL = strSQL & "LV.TotParcela, "
    strSQL = strSQL & "LV1.dbl1valor, "
    strSQL = strSQL & "LV1.dtmdtvencimentoParcela, "
    strSQL = strSQL & "LIC1.dblPorcDemolicao, "
    strSQL = strSQL & "LIC1.dblarealancada, "
    strSQL = strSQL & "LIC1.dblvalorm2, "
    strSQL = strSQL & "LIC1.dblvalorservico, "
    strSQL = strSQL & "LIC1.dblaliquotaiss, "
    strSQL = strSQL & "LIC1.dblvalorlancto, "
    strSQL = strSQL & "LIC1.dblvalorabatido, "
    strSQL = strSQL & "LIC1.dblSaldo "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLanctoIssConstrucao & " LI, "
    
    'Select para trazer somat�ria da tabela de pr�dios ISS
    strSQL = strSQL & "(Select "
    strSQL = strSQL & "Sum(LIC.dblPorcDemolicao) as dblPorcDemolicao, "
    strSQL = strSQL & "Sum(LIC.dblarealancada) as dblarealancada, "
    strSQL = strSQL & "Sum(LIC.dblvalorm2) as dblvalorm2, "
    strSQL = strSQL & "Sum(LIC.dblvalorservico) as dblvalorservico, "
    strSQL = strSQL & "LIC.dblaliquotaiss, "
    strSQL = strSQL & "Sum(LIC.dblvalorlancto) as dblvalorlancto, "
    strSQL = strSQL & "Sum(LIC.dblvalorabatido) as dblvalorabatido, "
    'strSQL = strSQL & "(Sum(LIC.dblvalorlancto)  -  Sum(LIC.dblvalorabatido)) dblSaldo "
    strSQL = strSQL & "((CASE WHEN SUM(LIC.dblValorLancto) IS NULL THEN 0 ELSE SUM(LIC.dblValorLancto) END) - "
    strSQL = strSQL & "(CASE WHEN SUM(LIC.dblValorAbatido)IS NULL THEN 0 ELSE SUM(LIC.dblValorAbatido) END)) dblSaldo "
    strSQL = strSQL & "From " & gstrLanctoIssConstrucao & " LI," & gstrLanctoIssConstrucaoPredios & " LIC "
    strSQL = strSQL & "Where LI.Pkid" & strOUTJOracle & "=" & strOUTJSQLServer & "LIC.INTLANCTOISSCONSTRUCAO AND LI.Intlancamentoalfa = " & lngPkid & " Group by LIC.dblaliquotaiss) LIC1, "
    
    'Select para trazer Qtde de parcelas
    strSQL = strSQL & "(Select Count(intParcela) as TotParcela From "
    strSQL = strSQL & gstrLancamentoValor & " Where Intlancamentoalfa =" & lngPkid & " ) LV, "
    
    'Select para trazer 1� Vencimento e 1� Valor de parcela
    strSql1 = ""
    strSql1 = strSql1 & "Select " & gstrTOPnSQLServer(1) & "dblvalor as dbl1valor, dtmdtvencimento as dtmdtvencimentoParcela From "
    strSql1 = strSql1 & gstrLancamentoValor & " Where intLancamentoAlfa = " & lngPkid & " Order by dtmdtvencimento"
    strSql1 = "(" & gstrTOPnOracle(strSql1, 1) & ") LV1 "
    
    strSQL = strSQL & strSql1

    strSQL = strSQL & "WHERE LA.Pkid = " & lngPkid & " AND LI.intLancamentoAlfa = LA.Pkid "
    
    strQueryCarneISSConstrucao = strSQL

End Function

Public Function blnAtualizacaoDeDebitos(strInscricao As String, ByRef vetParcelas() As String, Optional lngComposicao As Long = 0) As Boolean
    Dim adoResultado            As ADODB.Recordset
    Dim adoParcelas             As ADODB.Recordset
    Dim strSQL                  As String
    Dim strAcordosParaConsulta  As String
    Dim strInscricoes           As String
    Dim intFor                  As Integer
    
    blnAtualizacaoDeDebitos = False
    
    ReDim vetParcelas(21, 0)
    
    'Vamos obter os Pkids das inscricoes para fazer consulta de acordos
    strSQL = "SELECT  LA.Pkid " & _
             "FROM " & gstrLancamentoAlfa & " LA " & _
             "WHERE LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & UCase(strInscricao) & "' AND " & _
             "(LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
             If lngComposicao > 0 Then
                 strSQL = strSQL & " AND LA.intComposicaoDaReceita = " & lngComposicao
             End If
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Do While Not adoResultado.EOF
                strAcordosParaConsulta = strAcordosParaConsulta & adoResultado("Pkid").Value & ","
                adoResultado.MoveNext
            Loop
            strAcordosParaConsulta = Mid(strAcordosParaConsulta, 1, Len(strAcordosParaConsulta) - 1)
        End If
    Else
        Exit Function
    End If

ConsultarAcordos:

    'Vamos obter os acordos, caso exista, para exibir no grid Pai
    strSQL = "SELECT  LV.intLancamentoAlfaAcordo " & _
             "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA " & _
             "WHERE LV.intLancamentoAlfa = LA.pkid AND " & _
             "LA.Pkid IN (" & strAcordosParaConsulta & ") AND Not LV.intLancamentoAlfaAcordo Is Null " & _
             "GROUP BY LV.intLancamentoAlfaAcordo "
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            strAcordosParaConsulta = Space$(0)
            Do While Not adoResultado.EOF
                strInscricoes = strInscricoes & adoResultado("intlancamentoalfaacordo").Value & ","
                strAcordosParaConsulta = strAcordosParaConsulta & adoResultado("intlancamentoalfaacordo").Value & ","
                adoResultado.MoveNext
            Loop
            strAcordosParaConsulta = Mid(strAcordosParaConsulta, 1, Len(strAcordosParaConsulta) - 1)
            GoTo ConsultarAcordos
        End If
    End If


    strSQL = "SELECT LV.Pkid PkidLV, LV.bitParcelaValida, LA.intExercicio, LV.intLancamentoAlfa, LV.intParcela, "
    strSQL = strSQL & "LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.intLancamentoAlfaAcordo, LV.intLancamentoAlfaDAtiva, "
    strSQL = strSQL & "LA.strInscricao, " & gstrCONVERT(cdt_numeric, "LA.strNumeroAviso") & " strNumeroAviso, "
    strSQL = strSQL & "LA.intComposicaoDaReceita, LA.strComposicaoDaReceita, " & strSUBSTRING & "(LAA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & " " & gstrRIGHT("LAA.strInscricao", 4) & " Acordo, LA.intUtilizacao "
    strSQL = strSQL & "FROM " & gstrLancamentoValor & " LV, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrAcordo & " AC, "
    strSQL = strSQL & gstrLancamentoAlfa & " LAA, "
    strSQL = strSQL & gstrLancamentoPagamento & " LP "
    strSQL = strSQL & "WHERE LV.intLancamentoAlfa = LA.pkid AND "
    strSQL = strSQL & "LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= AC.intLancamentoAlfa " & strOUTJOracle & " And "
    strSQL = strSQL & "LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= LAA.Pkid " & strOUTJOracle & " And "
    strSQL = strSQL & "LV.Pkid" & strOUTJSQLServer & "= LP.Intlancamentovalor " & strOUTJOracle & " And "
    strSQL = strSQL & "LP.Intlancamentovalor Is Null And "
    strSQL = strSQL & "LA.strInscricao = '" & String(gintLenInscricao - Len(Trim(strInscricao)), "0") & UCase(strInscricao) & "' And "
    strSQL = strSQL & "(LA.INTUTILIZACAO = " & TYP_ACORDO & " OR (LA.INTUTILIZACAO <> " & TYP_ACORDO & " AND LA.bytNaoInscreveda = 0)) "
    If lngComposicao > 0 Then
        strSQL = strSQL & " AND LA.intComposicaoDaReceita = " & lngComposicao
    End If
             
    'Consulta que retorna os acordos
    If Len(strInscricoes) > 0 Then
        
        strInscricoes = Mid(strInscricoes, 1, Len(strInscricoes) - 1)
        
        strSQL = strSQL & " UNION ALL "
        strSQL = strSQL & "SELECT LV.Pkid PkidLV, LV.bitParcelaValida, LA.intExercicio, LV.intLancamentoAlfa, LV.intParcela, LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda, LV.intLancamentoAlfaAcordo, LV.intLancamentoAlfaDAtiva, " & _
                 "LA.strInscricao, " & gstrCONVERT(cdt_numeric, "LA.strNumeroAviso") & " strNumeroAviso, LA.intComposicaoDaReceita, LA.strComposicaoDaReceita, " & strSUBSTRING & "(LAA.strInscricao, " & gintLenInscricao - gintRetornaTamanhoMascara(TYP_ACORDO) + 1 & "," & gintRetornaTamanhoMascara(TYP_ACORDO) - 4 & ")" & strCONCAT & " '/' " & strCONCAT & " " & gstrRIGHT("LAA.strInscricao", 4) & " Acordo, LA.intUtilizacao " & _
                 "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA, " & gstrAcordo & " AC, " & gstrLancamentoAlfa & " LAA " & _
                 "WHERE LV.intLancamentoAlfa = LA.pkid AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= AC.intLancamentoAlfa " & strOUTJOracle
                 strSQL = strSQL & " AND LV.intLancamentoAlfaAcordo " & strOUTJSQLServer & "= LAA.Pkid " & strOUTJOracle
                 strSQL = strSQL & " AND LV.Pkid not in(Select Intlancamentovalor From " & gstrLancamentoPagamento & ")" & _
                                   " AND LA.Pkid IN (" & strInscricoes & ") "
    End If

    If bytDBType = EDatabases.Oracle Then
       strSQL = strSQL & " ORDER BY intLancamentoAlfa, intParcela"
    Else
       strSQL = strSQL & " ORDER BY LV.intLancamentoAlfa, LV.intParcela"
    End If

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
            
                For intFor = 0 To adoResultado.RecordCount - 1
                    strSQL = gstrStoredProcedure("sp_AtualizaParcela", !intComposicaoDaReceita & ", " & !intExercicio & ", " & !intParcela & ", " & gstrConvDtParaSql(!Dtmdtvencimento) & ", " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & gstrConvVrParaSql(!ValorOrig) & ", " & !intMoeda, True)
                    
                    Set gobjBanco = New clsBanco
                    If gobjBanco.CriaADO(strSQL, 80, adoParcelas) Then
                    
                        ReDim Preserve vetParcelas(21, intFor)
                        
                        vetParcelas(0, intFor) = Space$(0) & adoResultado("PkidLV").Value
                        vetParcelas(1, intFor) = Space$(0) & adoResultado("intLancamentoAlfa").Value
                        vetParcelas(2, intFor) = Space$(0) & adoResultado("intParcela").Value
                        vetParcelas(3, intFor) = Space$(0) & CCur(gstrConvVrDoSql(adoResultado("ValorOrig").Value))
                        vetParcelas(4, intFor) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value))
                        vetParcelas(5, intFor) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value))
                        vetParcelas(6, intFor) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value))
                        vetParcelas(7, intFor) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                        vetParcelas(8, intFor) = Space$(0) & CCur(gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)) + CCur(gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value))
                        vetParcelas(9, intFor) = Space$(0) & adoResultado("Acordo").Value
                        vetParcelas(10, intFor) = IsNull(adoResultado("intLancamentoAlfaAcordo").Value)
                        vetParcelas(11, intFor) = Space$(0) & gstrDataFormatada(adoResultado("dtmDtVencimento").Value)
                        vetParcelas(12, intFor) = Space$(0) & gstrFormataInscricao(Right(adoResultado("strInscricao").Value, gintRetornaTamanhoMascara(adoResultado("intUtilizacao").Value)), adoResultado("intUtilizacao").Value)
                        vetParcelas(13, intFor) = Space$(0) & adoResultado("strNumeroAviso").Value
                        vetParcelas(14, intFor) = Space$(0) & adoResultado("intExercicio").Value
                        vetParcelas(15, intFor) = Space$(0) & adoResultado("strComposicaoDaReceita").Value
                        vetParcelas(16, intFor) = Space$(0) & adoResultado("strInscricao").Value
                        vetParcelas(17, intFor) = Space$(0) & adoResultado("intUtilizacao").Value
                        vetParcelas(18, intFor) = Space$(0) & IIf(IsNull(adoResultado("intLancamentoAlfaDAtiva").Value), "N�o", "Sim")
                        vetParcelas(19, intFor) = Space$(0) & adoResultado("bitParcelaValida").Value
                        vetParcelas(20, intFor) = Space$(0) & adoResultado("intMoeda").Value
                        vetParcelas(21, intFor) = Space$(0) & adoResultado("intComposicaoDaReceita").Value
                        adoResultado.MoveNext
                    End If
                Next
            End With
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    
    blnAtualizacaoDeDebitos = True

End Function

Public Function gstrINSTR(strString As String, strCampo As String, intPosicao As Integer, Optional intOcorrencia As Integer = 0) As String

'******************************************************************************************
' Data: 28/05/2003
' Descri��o: - strString --> String que a ser pesquisada.
'            - strCampo --> Campo para busca
'            - intPosicao --> Inicio da busca
' Respons�vel: Anderson
'******************************************************************************************

    If (bytDBType = EDatabases.Oracle) Then
        gstrINSTR = " INSTR(" & strCampo & ", '" & strString & "'," & intPosicao
        If intOcorrencia > 0 Then gstrINSTR = gstrINSTR & ", " & intPosicao
        gstrINSTR = gstrINSTR & ")"
    Else
        gstrINSTR = " CHARINDEX('" & strString & "', " & strCampo & ", " & intPosicao & ")"
    End If

End Function

Public Function gstrCertidaoPorExecutivo(ByVal strExecutivoPKID As String) As String
    Dim strSQL As String
    Dim adoResultado  As ADODB.Recordset
    Dim i As Integer
    
    
    strSQL = ""
    strSQL = strSQL & " SELECT DISTINCT da.intcertidao "
    strSQL = strSQL & " FROM " & gstrDativa & " DA "
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " DA.INTEXECUTIVO = " & strExecutivoPKID
        
    Set gobjBanco = New clsBanco
   
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
        While Not adoResultado.EOF
            If Len(Trim(adoResultado!intCertidao)) > 0 Then
               gstrCertidaoPorExecutivo = gstrCertidaoPorExecutivo & gstrConvVrDoSql(adoResultado!intCertidao, 0) & ", "
            End If
            
            adoResultado.MoveNext
        Wend
    End If
    
    If Trim(gstrCertidaoPorExecutivo) <> "" Then
        gstrCertidaoPorExecutivo = Trim(Mid(Trim(gstrCertidaoPorExecutivo), 1, Len(gstrCertidaoPorExecutivo) - 2))
        For i = Len(gstrCertidaoPorExecutivo) To 1 Step -1
            If Mid(gstrCertidaoPorExecutivo, i, 1) = "," Then
                gstrCertidaoPorExecutivo = Mid(gstrCertidaoPorExecutivo, 1, i - 1) & " e" & Mid(gstrCertidaoPorExecutivo, i + 1)
                Exit For
            End If
        Next
    End If
    
End Function

Public Function dblCalculaEncargos(bitTipoEncargo As Byte, dblValorTotal As Double, Optional strExecutivo As String, Optional intComposicaoDaReceita As Long) As Double
Dim strSQL As String
Dim adoResultado As ADODB.Recordset
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "dblPorcHonorarios "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & " tblParametrosTributario PT "
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            Select Case bitTipoEncargo
                Case Is = BIT_HONORARIOS
                    'Caso exista executivo vamos calcular Honorarios e Custas
                    If Len(Trim(strExecutivo)) <> 0 Then
                        dblCalculaEncargos = FormatCurrency(dblValorTotal * adoResultado("dblPorcHonorarios").Value, 2)
                    End If
            End Select
        Else
            dblCalculaEncargos = 0
        End If
    End If

End Function

Private Sub MontaSubBandSaldoDividaAtiva(ab As ActiveBar2)
    Dim chd
    
    Set chd = ab.Bands.Add("ChildBand")
    With chd
        .Name = "chdSubSaldoDividaAtiva"
        .DockingArea = ddDAPopup
        .Caption = " "
        .GrabHandleStyle = ddGSNone
        .Type = ddBTNormal
        .Visible = False
        .flags = 127
        .Width = 1200
    End With
    
    MontaBotoes chd, 1427, "mnuSubSaldoDividaAtiva", "Geral                        "
    MontaBotoes chd, 1428, "mnuSubSaldoDividaAtiva", "Per�odo De Inscri��o"
End Sub

